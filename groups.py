#!/usr/bin/env python3
"""
Bible Study Group Attendance Analysis - Fixed Calendar Years Version
Analyzes attendance for members of Bible Study Groups
Reports on 'Last Calendar Year' and 'This Calendar Year'
Follow-up list: First shows zero attendance this year, then last year
"""

import subprocess
import sys

# Auto-install required packages
def install_packages():
    """
    Auto-install required packages.
    - plotly: charting
    - kaleido: static image export for plotly.write_image()
    - beautifulsoup4: HTML parsing
    - requests: API/HTTP
    """
    packages = ['beautifulsoup4', 'plotly', 'kaleido', 'requests']
    for package in packages:
        try:
            if package == 'beautifulsoup4':
                import bs4  # noqa: F401
            else:
                __import__(package)
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

# Install packages first
install_packages()

import requests
import json
from datetime import datetime, timedelta
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from bs4 import BeautifulSoup
import re
import webbrowser
import os

# Global variable for API key
# Get API key from config file
try:
    from config import ELVANTO_API_KEY
    API_KEY = ELVANTO_API_KEY
    print("✅ API key loaded from config.py")
except ImportError:
    print("❌ Error: config.py file not found!")
    print("Please create config.py with your ELVANTO_API_KEY")
    sys.exit(1)
except AttributeError:
    print("❌ Error: ELVANTO_API_KEY not found in config.py")
    print("Please add ELVANTO_API_KEY = 'your_key_here' to config.py")
    sys.exit(1)
BASE_URL = "https://api.elvanto.com/v1/"

def get_api_key():
    """Get API key from user input"""
    global API_KEY
    if not API_KEY:
        API_KEY = input("Enter your Elvanto API key: ").strip()
    return API_KEY

def make_request(endpoint, params=None):
    """Make authenticated request to Elvanto API with better debugging"""
    api_key = get_api_key()
    if not api_key:
        print("No API key provided")
        return None
    
    url = f"{BASE_URL}{endpoint}.json"
    auth = (api_key, '')
    
    try:
        response = requests.post(url, json=params, auth=auth, timeout=30)
        print(f"   API Call: {endpoint} -> Status: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            if data.get('status') == 'ok':
                return data
            else:
                error_info = data.get('error', {})
                print(f"   API Error: {error_info.get('message', 'Unknown error')}")
                print(f"   Error Code: {error_info.get('code', 'No code')}")
                return None
        else:
            print(f"   HTTP Error {response.status_code}: {response.text[:200]}")
            return None
    except Exception as e:
        print(f"   Request failed: {e}")
        return None

def is_bible_study_group(group):
    """Check if group has the 'Bible Study Groups_' category"""
    if group.get('categories') and group['categories'].get('category'):
        categories = group['categories']['category']
        if not isinstance(categories, list):
            categories = [categories] if categories else []
        for cat in categories:
            if isinstance(cat, dict):
                cat_name = cat.get('name', '')
            else:
                cat_name = str(cat)

            if cat_name == 'Bible Study Groups_':
                return True
    return False

def test_api_connection():
    """Test API connection"""
    print("Testing API connection...")
    data = make_request('groups/getAll', {'page_size': 1000, 'page': 1})
    if data:
        print("API connection successful!")
        return True
    else:
        print("API connection failed!")
        return False

def normalize_name(name):
    """Normalize a name for comparison"""
    if not name:
        return ""
    return re.sub(r'[^a-zA-Z0-9\s]', '', name.lower()).strip()

def find_attendance_report_groups():
    """Find attendance report groups (reports are stored as groups in Elvanto)"""
    print("Searching for attendance report groups...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_year_group = None
    last_year_group = None
    current_year = datetime.now().year

    print(f"Scanning {len(groups)} groups for attendance reports...")
    
    for group in groups:
        group_name = group.get('name', '').lower()
        
        # Look for the correct group name patterns
        if 'report of group individual attendance' in group_name:
            # This should be the current year report
            current_year_group = group
            print(f"Found current year report: {group.get('name')}")
            
        elif 'report of last year group individual attendance' in group_name:
            # This should be the last year report
            last_year_group = group
            print(f"Found last year report: {group.get('name')}")
        
        # Also check for service attendance reports (alternative naming)
        elif ('service individual attendance' in group_name or 
              'individual service attendance' in group_name):
            
            if str(current_year) in group_name:
                current_year_group = group
                print(f"Found current year service report: {group.get('name')}")
            elif str(current_year - 1) in group_name:
                last_year_group = group
                print(f"Found last year service report: {group.get('name')}")

    return current_year_group, last_year_group

def download_group_attendance_data(group):
    """Download attendance data from a group (following successful pattern)"""
    if not group:
        return None
        
    group_name = group.get('name', 'Unknown')
    print(f"Getting attendance data from group: {group_name}")
    
    # Extract URL from group location fields (using original group object)
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break
    
    if not report_url:
        print("No URL found in group location fields")
        return None
    
    print("Found report URL in group")
    
    # Fetch and parse HTML from the URL
    try:
        import requests
        response = requests.get(report_url, timeout=60)
        if response.status_code == 200:
            content = response.text
            print(f"Downloaded {len(content)} characters from report URL")
            return content
        else:
            print(f"Failed to fetch report: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error fetching report: {e}")
        return None

def extract_groups_from_attendance_data(html_content):
    """Extract group names and attendance from HTML report data"""
    print("Parsing attendance data...")
    
    if not html_content:
        print("No content to parse")
        return {}
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find the main table
    table = soup.find('table')
    if not table:
        print("No table found in report")
        return {}
    
    rows = table.find_all('tr')
    if not rows:
        print("No rows found in table")
        return {}
    
    print(f"Found {len(rows)} rows in attendance table")
    
    # Extract header to find date columns
    header_row = rows[0] if rows else None
    date_columns = []
    date_info = {}  # column_index: date_string
    
    if header_row:
        cells = header_row.find_all(['th', 'td'])
        for i, cell in enumerate(cells):
            cell_text = cell.get_text().strip()
            # Look for date patterns (dd/mm format)
            if re.match(r'\d{1,2}/\d{1,2}', cell_text):
                date_columns.append(i)
                date_info[i] = cell_text
        print(f"Found {len(date_columns)} date columns: {date_columns}")
    
    group_data = {}
    current_group = None
    group_rows = []  # Store all rows for current group to analyze meeting patterns
    
    for row_idx, row in enumerate(rows[1:], 1):  # Skip header
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        # Convert cells to text for easier processing
        row_data = [cell.get_text().strip() for cell in cells]
        
        # Check if this is a group header row (black background, contains group name)
        first_cell = cells[0]
        is_group_header = False
        
        # Check for styling that indicates group header
        style = first_cell.get('style', '')
        class_attr = first_cell.get('class', [])
        
        if ('background' in style.lower() and 'black' in style.lower()) or \
           any('header' in str(cls).lower() or 'group' in str(cls).lower() for cls in class_attr):
            is_group_header = True
        
        # Also check if row contains typical group names
        first_cell_text = row_data[0] if row_data else ""
        if any(keyword in first_cell_text.lower() for keyword in 
               ['bible study', 'youth group', 'ever attended', 'kids club', 'small group', 'home group']):
            is_group_header = True
        
        if is_group_header:
            # Process previous group data
            if current_group and group_rows:
                group_data[current_group] = process_group_attendance(current_group, group_rows, date_columns)
            
            # Start new group
            current_group = first_cell_text
            group_rows = []
            print(f"Found group header: '{current_group}'")
            continue
        
        # Collect rows for current group
        if current_group and row_data:
            # Check if first column contains a person's name
            first_cell_value = row_data[0].strip()
            
            if first_cell_value and ',' in first_cell_value:
                group_rows.append(row_data)
    
    # Don't forget the last group
    if current_group and group_rows:
        group_data[current_group] = process_group_attendance(current_group, group_rows, date_columns)
    
    return group_data

def process_group_attendance(group_name, group_rows, date_columns):
    """Process attendance data for a specific group"""
    group_attendees = []
    group_all_people = []
    group_recent_missed = []
    
    # Find which date columns this group actually used (had meetings)
    group_meeting_dates = []
    for col_idx in date_columns:
        has_data = False
        for row_data in group_rows:
            if col_idx < len(row_data):
                cell = str(row_data[col_idx]).strip().upper()
                if cell in ['Y', 'N']:  # This date had a meeting for this group
                    has_data = True
                    break
        if has_data:
            group_meeting_dates.append(col_idx)
    
    # Get the last 3 meeting dates for this specific group
    last_3_meetings = group_meeting_dates[-3:] if len(group_meeting_dates) >= 3 else group_meeting_dates
    print(f"   {group_name}: Found {len(group_meeting_dates)} meeting dates, analyzing last {len(last_3_meetings)} meetings")
    
    # Process each person in the group
    for row_data in group_rows:
        first_cell_value = row_data[0].strip()
        
        if not first_cell_value or not ',' in first_cell_value:
            continue
        
        # Parse person's name
        name_parts = first_cell_value.split(',')
        if len(name_parts) >= 2:
            last_name = name_parts[0].strip()
            first_name = name_parts[1].strip()
            
            # Remove role information in parentheses
            if '(' in first_name:
                first_name = first_name.split('(')[0].strip()
            
            full_name = f"{first_name} {last_name}".strip()
            
            # Add to all_people list (everyone who appears in report)
            group_all_people.append(full_name)
            
            # Check if this person has ANY attendance in any group meeting dates
            attended_any = False
            for col_idx in group_meeting_dates:
                if col_idx < len(row_data):
                    cell = str(row_data[col_idx]).strip().upper()
                    if cell == 'Y':  # Attended
                        attended_any = True
                        break
            
            # Only add to attendees if they actually attended
            if attended_any:
                group_attendees.append(full_name)
            
            # Check if this person missed the last 3 meetings of THIS group
            if last_3_meetings:
                attended_recent = 0
                had_opportunity = 0
                
                for col_idx in last_3_meetings:
                    if col_idx < len(row_data):
                        cell = str(row_data[col_idx]).strip().upper()
                        if cell in ['Y', 'N']:  # They had opportunity to attend
                            had_opportunity += 1
                            if cell == 'Y':
                                attended_recent += 1
                
                # If they had opportunity to attend recent meetings but attended none
                if had_opportunity > 0 and attended_recent == 0:
                    group_recent_missed.append(full_name)
    
    unique_attendees = list(set(group_attendees))
    unique_all_people = list(set(group_all_people))
    unique_recent_missed = list(set(group_recent_missed))
    
    print(f"   {group_name}: {len(unique_attendees)} attendees, {len(unique_recent_missed)} missed last {len(last_3_meetings)} meetings")
    
    return {
        'attendees': unique_attendees,
        'all_people': unique_all_people,
        'recent_missed': unique_recent_missed
    }

def fetch_all_groups_from_api():
    """Fetch all groups from API"""
    print("Fetching groups with categories and people...")
    all_groups = []
    page = 1
    while True:
        print(f"Groups page {page}...", end=" ")
        response = make_request('groups/getAll', {
            'page': page, 
            'page_size': 1000,
            'fields': ['people', 'categories']
        })
        if not response or not response.get('groups'):
            print("Done")
            break
        groups = response['groups'].get('group', [])
        if not isinstance(groups, list):
            groups = [groups] if groups else []
        all_groups.extend(groups)
        print(f"({len(groups)} groups)")
        if len(groups) < 1000:  # KEY: Break when fewer than page_size returned
            break
        page += 1
    print(f"Total groups: {len(all_groups)}")
    return all_groups

def extract_monthly_attendance_data(html_content, year_label):
    """Extract monthly attendance data from HTML report"""
    print(f"Extracting monthly attendance data for {year_label}...")
    
    if not html_content:
        return {}
    
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    if not table:
        return {}
    
    rows = table.find_all('tr')
    if not rows:
        return {}
    
    # Extract header to find date columns and determine months
    header_row = rows[0] if rows else None
    date_columns = {}  # column_index: month_number
    
    if header_row:
        cells = header_row.find_all(['th', 'td'])
        for i, cell in enumerate(cells):
            cell_text = cell.get_text().strip()
            # Look for date patterns (dd/mm format)
            if re.match(r'\d{1,2}/\d{1,2}', cell_text):
                try:
                    day, month = cell_text.split('/')
                    date_columns[i] = int(month)
                except:
                    continue
    
    print(f"   Found {len(date_columns)} date columns for {year_label}")
    
    monthly_data = {}  # group_name: {month: attendance_count}
    current_group = None
    
    for row_idx, row in enumerate(rows[1:], 1):  # Skip header
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        row_data = [cell.get_text().strip() for cell in cells]
        
        # Check if this is a group header row
        first_cell = cells[0]
        is_group_header = False
        
        style = first_cell.get('style', '')
        class_attr = first_cell.get('class', [])
        
        if ('background' in style.lower() and 'black' in style.lower()) or \
           any('header' in str(cls).lower() or 'group' in str(cls).lower() for cls in class_attr):
            is_group_header = True
        
        first_cell_text = row_data[0] if row_data else ""
        if any(keyword in first_cell_text.lower() for keyword in 
               ['bible study', 'youth group', 'ever attended', 'kids club', 'small group', 'home group', 'iff', 'international']):
            is_group_header = True
        
        if is_group_header:
            current_group = first_cell_text
            if current_group not in monthly_data:
                monthly_data[current_group] = {}
            continue
        
        # Process individual attendance row
        if current_group and row_data:
            first_cell_value = row_data[0].strip()
            
            if not first_cell_value or not ',' in first_cell_value:
                continue
            
            # Count attendances by month for this person
            for col_idx, month in date_columns.items():
                if col_idx < len(row_data):
                    cell = str(row_data[col_idx]).strip().upper()
                    if cell == 'Y':  # Attended
                        if month not in monthly_data[current_group]:
                            monthly_data[current_group][month] = 0
                        monthly_data[current_group][month] += 1
    
    return monthly_data

def create_progressive_attendance_charts(current_monthly_data, last_year_monthly_data, current_year, last_year):
    """Create progressive attendance charts with consistent colors, 3-across layout,
    save a PNG in the working folder, and auto-open it."""
    import os, math, subprocess, sys
    from datetime import datetime
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    print("Creating progressive monthly attendance charts...")

    # ----- Fixed colors so they are consistent across ALL subplots -----
    COLOR_THIS_YEAR  = '#3B82F6'  # blue-500
    COLOR_LAST_YEAR  = '#F97316'  # orange-500
    COLOR_BENCHMARK  = '#22C55E'  # green-500

    # Categorize groups
    def categorize_group(group_name):
        name_lower = group_name.lower()
        if 'kids club' in name_lower or 'youth group' in name_lower:
            return 'kids_youth'
        elif 'iff' in name_lower or 'international food' in name_lower:
            return 'iff'
        else:
            return 'regular_bible_studies'

    # Build category totals for each year
    categories = {'all': {}, 'kids_youth': {}, 'iff': {}, 'regular_bible_studies': {}}
    for group_name, monthly_counts in (current_monthly_data or {}).items():
        cat = categorize_group(group_name)
        for m, cnt in monthly_counts.items():
            categories[cat][m] = categories[cat].get(m, 0) + cnt
            categories['all'][m] = categories['all'].get(m, 0) + cnt

    last_year_categories = {'all': {}, 'kids_youth': {}, 'iff': {}, 'regular_bible_studies': {}}
    for group_name, monthly_counts in (last_year_monthly_data or {}).items():
        cat = categorize_group(group_name)
        for m, cnt in monthly_counts.items():
            last_year_categories[cat][m] = last_year_categories[cat].get(m, 0) + cnt
            last_year_categories['all'][m] = last_year_categories['all'].get(m, 0) + cnt

    def create_cumulative_data(monthly_data, year):
        # last_year = full 12 months; this_year = up to current month
        from datetime import datetime
        months = list(range(1, 13)) if year == last_year else list(range(1, datetime.now().month + 1))
        running = 0
        cumulative = []
        for m in months:
            running += int(monthly_data.get(m, 0))
            cumulative.append(running)
        return months, cumulative

    # Individual group list
    individual_groups = sorted(set(list((current_monthly_data or {}).keys()) +
                                   list((last_year_monthly_data or {}).keys())))
    num_individual = len(individual_groups)

    # ----- Grid: 3 across (portrait overall) -----
    total_charts = 4 + num_individual
    cols = 3
    rows = math.ceil(total_charts / cols)

    month_names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    # Titles: 4 summary + all individuals
    titles = [
        'All Groups Combined',
        'Kids Club + Youth Group',
        'International Food & Friends (IFF)',
        'Regular Bible Study Groups',
        *individual_groups
    ]

    fig = make_subplots(
        rows=rows, cols=cols,
        subplot_titles=titles,
        vertical_spacing=0.06,
        horizontal_spacing=0.06
    )

    # Helper to place traces sequentially in 3-across grid
    def cell_from_index(idx):
        r = (idx // cols) + 1
        c = (idx % cols) + 1
        return r, c

    # ----- First 4 category charts -----
    category_keys = ['all', 'kids_youth', 'iff', 'regular_bible_studies']
    for i, cat in enumerate(category_keys):
        row, col = cell_from_index(i)

        last_m, last_cum = create_cumulative_data(last_year_categories[cat], last_year)
        this_m, this_cum = create_cumulative_data(categories[cat], current_year)

        last_labels = [month_names[m-1] for m in last_m]
        this_labels = [month_names[m-1] for m in this_m]

        # 10% benchmark line based on last year total
        benchmark_target = (last_cum[-1] if last_cum else 0) * 1.10
        benchmark_line = [(benchmark_target * (i+1) / 12.0) for i in range(len(this_m))]

        # Last year
        fig.add_trace(
            go.Scatter(x=last_labels, y=last_cum, mode='lines+markers',
                       name=f'{last_year}', line=dict(color=COLOR_LAST_YEAR, width=3),
                       marker=dict(size=7), showlegend=(i == 0), legendgroup='last'),
            row=row, col=col
        )
        # This year
        fig.add_trace(
            go.Scatter(x=this_labels, y=this_cum, mode='lines+markers',
                       name=f'{current_year}', line=dict(color=COLOR_THIS_YEAR, width=3),
                       marker=dict(size=7), showlegend=(i == 0), legendgroup='this'),
            row=row, col=col
        )
        # Benchmark
        if benchmark_line:
            fig.add_trace(
                go.Scatter(x=this_labels, y=benchmark_line, mode='lines',
                           name='10% Growth Target', line=dict(color=COLOR_BENCHMARK, width=2, dash='dash'),
                           showlegend=(i == 0), legendgroup='bench'),
                row=row, col=col
            )

    # ----- Individual group charts -----
    start_idx = 4
    for offset, group_name in enumerate(individual_groups):
        idx = start_idx + offset
        row, col = cell_from_index(idx)

        this_data = (current_monthly_data or {}).get(group_name, {}) or {}
        last_data = (last_year_monthly_data or {}).get(group_name, {}) or {}

        this_m, this_cum = create_cumulative_data(this_data, current_year)
        last_m, last_cum = create_cumulative_data(last_data, last_year)

        this_labels = [month_names[m-1] for m in this_m]
        last_labels = [month_names[m-1] for m in last_m]

        benchmark_target = (last_cum[-1] if last_cum else 0) * 1.10
        benchmark_line = [(benchmark_target * (i+1) / 12.0) for i in range(len(this_m))]

        if last_cum:
            fig.add_trace(
                go.Scatter(x=last_labels, y=last_cum, mode='lines+markers',
                           line=dict(color=COLOR_LAST_YEAR, width=3),
                           marker=dict(size=6), showlegend=False, legendgroup='last'),
                row=row, col=col
            )
        if this_cum:
            fig.add_trace(
                go.Scatter(x=this_labels, y=this_cum, mode='lines+markers',
                           line=dict(color=COLOR_THIS_YEAR, width=3),
                           marker=dict(size=6), showlegend=False, legendgroup='this'),
                row=row, col=col
            )
        if benchmark_line and this_cum:
            fig.add_trace(
                go.Scatter(x=this_labels, y=benchmark_line, mode='lines',
                           line=dict(color=COLOR_BENCHMARK, width=2, dash='dash'),
                           showlegend=False, legendgroup='bench'),
                row=row, col=col
            )

    # Axes labels
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            fig.update_yaxes(title_text="Cumulative Attendance", row=r, col=c, title_font_size=12)
            fig.update_xaxes(title_text="Month", row=r, col=c, title_font_size=12)

    # Portrait-ish proportions
    width = 1800
    height = max(1100, 380 * rows)

    fig.update_layout(
        title=f'Group Attendance Progress (Cumulative)<br><span style="font-size:12px">Last Year ({last_year}) vs This Year ({current_year})</span>',
        showlegend=True,
        height=height,
        width=width,
        font=dict(size=11),
        legend=dict(orientation="h", yanchor="bottom", y=1.03, xanchor="center", x=0.5)
    )

    # ----- Save both HTML and PNG to the current working folder -----
    html_name = f"bible_study_progressive_attendance_calendar_years_{current_year}_{last_year}.html"
    png_name  = f"bible_study_progressive_attendance_calendar_years_{current_year}_{last_year}.png"

    try:
        fig.write_html(html_name)
        print(f"✅ Saved HTML: {os.path.abspath(html_name)}")
    except Exception as e:
        print(f"⚠️ Could not save HTML: {e}")

    # PNG requires kaleido
    try:
        fig.write_image(png_name, width=width, height=height, scale=2)  # high-res
        png_path = os.path.abspath(png_name)
        print(f"✅ Saved PNG: {png_path}")

        # ----- Auto-open the PNG (cross-platform) -----
        try:
            if sys.platform.startswith('darwin'):
                subprocess.Popen(['open', png_path])
            elif os.name == 'nt':
                os.startfile(png_path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(['xdg-open', png_path])
            print("✅ PNG auto-opened.")
        except Exception as e:
            print(f"⚠️ PNG saved but could not auto-open: {e}")
            print(f"   Please open manually: {png_path}")
    except Exception as e:
        print(f"❌ PNG save failed (is 'kaleido' installed?): {e}")

def create_charts(follow_up_this_year, follow_up_last_year, current_year, last_year):
    """Create bar charts showing number of people needing follow-up per group"""
    
    print("Creating charts with data:")
    print(f"   This year groups: {len(follow_up_this_year)} groups")
    print(f"   Last year groups: {len(follow_up_last_year)} groups")
    
    # Prepare data for charts
    this_year_groups = list(follow_up_this_year.keys())
    this_year_counts = [len(follow_up_this_year[group]) for group in this_year_groups]
    
    last_year_groups = list(follow_up_last_year.keys())
    last_year_counts = [len(follow_up_last_year[group]) for group in last_year_groups]
    
    # Create subplots - SIDE BY SIDE instead of top/bottom
    fig = make_subplots(
        rows=1, cols=2,  # Side by side layout
        subplot_titles=[
            f'Zero Attendance This Calendar Year ({current_year})',
            f'Zero Attendance Last Calendar Year ({last_year})'
        ],
        horizontal_spacing=0.1
    )
    
    # This year chart (left side)
    fig.add_trace(
        go.Bar(
            x=this_year_groups,
            y=this_year_counts,
            name=f'This Year ({current_year})',
            marker_color='rgba(239, 68, 68, 0.8)',
            text=this_year_counts,
            textposition='auto',
        ),
        row=1, col=1
    )
    
    # Last year chart (right side)
    fig.add_trace(
        go.Bar(
            x=last_year_groups,
            y=last_year_counts,
            name=f'Last Year ({last_year})',
            marker_color='rgba(249, 115, 22, 0.8)',
            text=last_year_counts,
            textposition='auto',
        ),
        row=1, col=2
    )
    
    # Update layout
    fig.update_layout(
        title=f'Bible Study Group Follow-up Analysis - Calendar Years {last_year} & {current_year}',
        showlegend=False,
        height=600,  # Reduced height since side-by-side
        width=1400,  # Increased width for side-by-side
        font=dict(size=12)
    )
    
    # Update x-axis labels
    fig.update_xaxes(tickangle=45)
    fig.update_yaxes(title_text="Number of People", row=1, col=1)
    fig.update_yaxes(title_text="Number of People", row=1, col=2)
    
    # Save chart with absolute path info
    import os
    filename = f"bible_study_followup_calendar_years_{current_year}_{last_year}.html"
    current_dir = os.getcwd()
    full_path = os.path.join(current_dir, filename)
    
    # Remove existing file if it exists to ensure we can overwrite
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            print(f"   Removed existing file: {filename}")
        except Exception as e:
            print(f"   Warning: Could not remove existing file: {e}")
    
    try:
        fig.write_html(filename)
        print("Chart saved successfully!")
        print(f"   Directory: {current_dir}")
        print(f"   Filename: {filename}")
        print(f"   Full path: {full_path}")
        print(f"   File exists: {os.path.exists(full_path)}")
        if os.path.exists(full_path):
            file_size = os.path.getsize(full_path)
            print(f"   File size: {file_size} bytes")
    except Exception as e:
        print(f"Error saving chart: {e}")
        import traceback
        traceback.print_exc()

def create_followup_member_list(follow_up_recent_missed, follow_up_this_year, follow_up_last_year, current_year, last_year):
    """Create detailed HTML report of members needing follow-up"""
    
    print("Creating member list with data:")
    print(f"   Recent missed follow-ups: {sum(len(members) for members in follow_up_recent_missed.values()) if follow_up_recent_missed else 0}")
    print(f"   This year follow-ups: {sum(len(members) for members in follow_up_this_year.values()) if follow_up_this_year else 0}")
    print(f"   Last year follow-ups: {sum(len(members) for members in follow_up_last_year.values()) if follow_up_last_year else 0}")
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bible Study Follow-up List - Calendar Years {last_year} & {current_year}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            color: #e2e8f0;
            margin: 0;
            padding: 20px;
            line-height: 1.6;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: rgba(15, 23, 42, 0.9);
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            backdrop-filter: blur(20px);
        }}
        
        .header {{
            background: linear-gradient(135deg, #dc2626, #b91c1c);
            color: white;
            padding: 30px;
            border-radius: 20px 20px 0 0;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5rem;
            font-weight: bold;
            margin-bottom: 10px;
        }}
        
        .summary {{
            padding: 30px;
            background: rgba(30, 41, 59, 0.5);
            border-bottom: 1px solid rgba(71, 85, 105, 0.3);
        }}
        
        .summary h2 {{
            color: #60a5fa;
            margin-bottom: 20px;
            font-size: 1.8rem;
        }}
        
        .summary-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }}
        
        .stat-card {{
            background: rgba(30, 41, 59, 0.4);
            padding: 15px;
            border-radius: 10px;
            text-align: center;
        }}
        
        .stat-number {{
            font-size: 2rem;
            font-weight: bold;
            color: #60a5fa;
        }}
        
        .stat-label {{
            color: #94a3b8;
            font-size: 0.9rem;
        }}
        
        .section-title {{
            background: linear-gradient(135deg, #dc2626, #b91c1c);
            color: white;
            padding: 20px 30px;
            font-size: 1.5rem;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 0;
        }}
        
        .section-title.urgent {{
            background: linear-gradient(135deg, #dc2626, #7f1d1d);
            animation: pulse 2s infinite;
        }}
        
        .section-title.priority {{
            background: linear-gradient(135deg, #ea580c, #c2410c);
        }}
        
        .section-title.last-year {{
            background: linear-gradient(135deg, #0891b2, #0e7490);
        }}
        
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.8; }}
        }}
        
        .group-section {{
            padding: 30px;
            border-bottom: 1px solid rgba(71, 85, 105, 0.3);
        }}
        
        .group-title {{
            color: #60a5fa;
            font-size: 1.4rem;
            font-weight: bold;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid rgba(96, 165, 250, 0.3);
        }}
        
        .member-list {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 15px;
        }}
        
        .member-card {{
            background: rgba(51, 65, 85, 0.4);
            padding: 15px;
            border-radius: 10px;
            border-left: 4px solid #60a5fa;
            transition: all 0.3s ease;
        }}
        
        .member-card:hover {{
            background: rgba(51, 65, 85, 0.6);
            transform: translateY(-2px);
        }}
        
        .member-name {{
            font-weight: bold;
            font-size: 1.1rem;
            margin-bottom: 5px;
            color: #f1f5f9;
        }}
        
        .member-role {{
            color: #94a3b8;
            font-size: 0.9rem;
        }}
        
        .no-followup {{
            text-align: center;
            color: #10b981;
            font-style: italic;
            padding: 20px;
        }}
        
        .urgent-note {{
            background: rgba(220, 38, 38, 0.1);
            border: 1px solid rgba(220, 38, 38, 0.3);
            border-radius: 10px;
            padding: 15px;
            margin: 20px 30px;
            color: #fca5a5;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Bible Study Follow-up List</h1>
            <p>Calendar Year Analysis: {last_year} & {current_year}</p>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
        """
    
    # Calculate summary statistics - NOW INCLUDING RECENT MISSED
    total_recent_missed = sum(len(members) for members in follow_up_recent_missed.values()) if follow_up_recent_missed else 0
    total_this_year = sum(len(members) for members in follow_up_this_year.values()) if follow_up_this_year else 0
    total_last_year = sum(len(members) for members in follow_up_last_year.values()) if follow_up_last_year else 0
    
    # Add summary section
    html_content += f"""
        <div class="summary">
            <h2>Summary</h2>
            <p>Analysis of Bible Study Group members requiring follow-up based on attendance patterns.</p>
            <p><strong>Priority Order:</strong> 1) Missed last 3 meetings (urgent), 2) Has not attended group this year, 3) Did not attend group last year.</p>
            <div class="summary-stats">
                <div class="stat-card">
                    <div class="stat-number">{total_recent_missed}</div>
                    <div class="stat-label">URGENT: Missed Last 3 Meetings</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{total_this_year}</div>
                    <div class="stat-label">Has Not Attended Group in {current_year}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{total_last_year}</div>
                    <div class="stat-label">Did Not Attend Group in {last_year}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{len(follow_up_recent_missed) + len(follow_up_this_year) + len(follow_up_last_year)}</div>
                    <div class="stat-label">Total Groups with Issues</div>
                </div>
            </div>
        </div>
        """
    
    # *** URGENT SECTION - MISSED LAST 3 MEETINGS (FIRST!) ***
    html_content += f"""
        <div class="section-title urgent">
            🚨 URGENT: Missed Last 3 Meetings ({current_year})
        </div>
        """
    
    if follow_up_recent_missed:
        html_content += """
        <div class="urgent-note">
            <strong>⚠️ IMMEDIATE FOLLOW-UP REQUIRED:</strong> These members have missed their group's last 3 meetings this calendar year and need urgent pastoral contact.
        </div>
        """
        
        for group_name, members in follow_up_recent_missed.items():
            html_content += f"""
        <div class="group-section">
            <div class="group-title">
                🚨 {group_name} ({len(members)} people)
            </div>
            <div class="member-list">
            """
            
            for member in members:
                role_display = f" ({member['role']})" if member['role'] else ""
                html_content += f"""
                <div class="member-card">
                    <div class="member-name">🚨 {member['name']}</div>
                    <div class="member-role">Member ID: {member['id']}{role_display} - MISSED LAST 3 MEETINGS</div>
                </div>
                """
            
            html_content += """
            </div>
        </div>
        """
    else:
        html_content += f'<div class="no-followup">✅ Excellent! No members missed the last 3 meetings in {current_year}!</div>'
    
    # Add This Calendar Year section
    html_content += f"""
        <div class="section-title priority">
            PRIORITY: HAS NOT ATTENDED GROUP IN {current_year}
        </div>
        """
    
    if follow_up_this_year:
        for group_name, members in follow_up_this_year.items():
            html_content += f"""
        <div class="group-section">
            <div class="group-title">
                {group_name} ({len(members)} people)
            </div>
            <div class="member-list">
            """
            
            for member in members:
                role_display = f" ({member['role']})" if member['role'] else ""
                html_content += f"""
                <div class="member-card">
                    <div class="member-name">{member['name']}</div>
                    <div class="member-role">Member ID: {member['id']}{role_display}</div>
                </div>
                """
            
            html_content += """
            </div>
        </div>
        """
    else:
        html_content += f'<div class="no-followup">No members failed to attend group in {current_year}!</div>'
    
    # Add Last Calendar Year section
    html_content += f"""
        <div class="section-title last-year">
            SECONDARY: DID NOT ATTEND GROUP IN {last_year}
        </div>
        """
    
    if follow_up_last_year:
        for group_name, members in follow_up_last_year.items():
            html_content += f"""
        <div class="group-section">
            <div class="group-title">
                {group_name} ({len(members)} people)
            </div>
            <div class="member-list">
            """
            
            for member in members:
                role_display = f" ({member['role']})" if member['role'] else ""
                html_content += f"""
                <div class="member-card">
                    <div class="member-name">{member['name']}</div>
                    <div class="member-role">Member ID: {member['id']}{role_display}</div>
                </div>
                """
            
            html_content += """
            </div>
        </div>
        """
    else:
        html_content += f'<div class="no-followup">No members failed to attend group in {last_year}!</div>'
    
    html_content += """
    </div>
</body>
</html>
"""
    
    # Save the HTML file with detailed path information
    import os
    filename = f"bible_study_followup_members_calendar_years_{current_year}_{last_year}.html"
    current_dir = os.getcwd()
    full_path = os.path.join(current_dir, filename)
    
    # Remove existing file if it exists to ensure we can overwrite
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            print(f"   Removed existing file: {filename}")
        except Exception as e:
            print(f"   Warning: Could not remove existing file: {e}")
    
    try:
        with open(filename, 'w', encoding='utf-8', newline='') as f:
            f.write(html_content)
        print("Member list saved successfully!")
        print(f"   Directory: {current_dir}")
        print(f"   Filename: {filename}")
        print(f"   Full path: {full_path}")
        print(f"   File exists: {os.path.exists(full_path)}")
        if os.path.exists(full_path):
            file_size = os.path.getsize(full_path)
            print(f"   File size: {file_size} bytes")
            
        # Automatically open the member follow-up list in default browser
        try:
            webbrowser.open(f"file://{full_path}")
            print(f"   ✅ Opened member follow-up list in browser automatically!")
        except Exception as e:
            print(f"   ⚠️ Could not auto-open browser: {e}")
            
        return filename
    except Exception as e:
        print(f"Failed to create member list: {e}")
        import traceback
        traceback.print_exc()
        return None

def cleanup_old_files():
    """Remove any existing Bible study report files to ensure fresh generation"""
    import os
    import glob
    
    # Find all existing Bible study HTML files
    patterns = [
        "bible_study_followup_*.html",
        "bible_study_progressive_*.html"
    ]
    
    all_existing_files = []
    for pattern in patterns:
        all_existing_files.extend(glob.glob(pattern))
    
    if all_existing_files:
        print(f"Cleaning up {len(all_existing_files)} existing report files...")
        for file in all_existing_files:
            try:
                os.remove(file)
                print(f"   Removed: {file}")
            except Exception as e:
                print(f"   Warning: Could not remove {file}: {e}")
    else:
        print("No existing report files to clean up")

def create_bible_study_attendance_analysis():
    """Main function to create the attendance analysis"""
    current_year = datetime.now().year
    last_year = current_year - 1

    print("BIBLE STUDY GROUP ATTENDANCE ANALYSIS")
    print(f"This Calendar Year: {current_year}")
    print(f"Last Calendar Year: {last_year}")
    print("=" * 80)

    # Clean up any existing files first
    print("\nCleaning up existing files...")
    cleanup_old_files()

    # Test API connection first
    if not test_api_connection():
        return None

    # Step 1: Find This Calendar Year and Last Calendar Year attendance report groups
    print("\nStep 1: Finding calendar year attendance report groups...")
    current_year_group, last_year_group = find_attendance_report_groups()

    if not current_year_group:
        print("No current year 'Individual Group Attendance' report group found")
        return None

    print("Found current year report group")

    # Step 2: Download attendance data from both groups
    print("\nStep 2: Downloading attendance data from groups...")
    current_year_data = download_group_attendance_data(current_year_group)

    if not current_year_data:
        print("Failed to download current year attendance data")
        return None

    # Download last year data if available
    last_year_data = None
    if last_year_group:
        last_year_data = download_group_attendance_data(last_year_group)
        print("Downloaded last year attendance data")
    else:
        print("No last year report group found, continuing with current year only")

    # Step 3: Extract attendance data (incl. recent missed)
    current_attendance_data = extract_groups_from_attendance_data(current_year_data)
    current_monthly_data = extract_monthly_attendance_data(current_year_data, "Current Year")

    if not current_attendance_data:
        print("No attendance data extracted from current year report")
        return None

    if last_year_data:
        last_year_attendance_data = extract_groups_from_attendance_data(last_year_data)
        last_year_monthly_data = extract_monthly_attendance_data(last_year_data, "Last Year")
    else:
        last_year_attendance_data = {}
        last_year_monthly_data = {}

    # Debug: show recent missed counts
    print(f"\nDEBUG - Attendance data extracted:")
    for group_name, group_info in current_attendance_data.items():
        recent_missed = group_info.get('recent_missed', [])
        print(f"   {group_name}: {len(recent_missed)} recent missed members")

    # Step 4: Get all groups from API
    print("\nStep 4: Fetching all groups from API...")
    all_groups = fetch_all_groups_from_api()
    if not all_groups:
        print("No groups found from API")
        return None

    # Filter to only Bible Study Groups
    bible_study_groups = [g for g in all_groups if is_bible_study_group(g)]
    print(f"Found {len(bible_study_groups)} Bible Study groups")
    if not bible_study_groups:
        print("No Bible Study groups found with 'Bible Study Groups_' category")
        return None

    # Step 5: Extract all members from Bible Study Groups
    bible_study_members = {}
    for group in bible_study_groups:
        group_name = group.get('name', 'Unknown Group')
        if any(skip in group_name.lower() for skip in ['report', 'admin', 'staff']):
            continue
        if group.get('people') and group['people'].get('person'):
            group_people = group['people']['person']
            if not isinstance(group_people, list):
                group_people = [group_people] if group_people else []
            lst = []
            for person in group_people:
                person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
                if person_name:
                    lst.append({'name': person_name, 'id': person.get('id'), 'role': person.get('role', '')})
            if lst:
                bible_study_members[group_name] = lst
                print(f"   {group_name}: {len(lst)} members")

    if not bible_study_members:
        print("No Bible study group members found")
        return None

    # Step 6: Analyze attendance patterns
    print("\nStep 6: Analyzing attendance patterns...")
    all_attendance_data = {}
    for group_name, group_info in current_attendance_data.items():
        if group_name not in all_attendance_data:
            all_attendance_data[group_name] = {
                'this_year_attendees': [], 'this_year_all_people': [],
                'this_year_recent_missed': [], 'last_year_attendees': [], 'last_year_all_people': []
            }
        all_attendance_data[group_name]['this_year_attendees'] = group_info.get('attendees', [])
        all_attendance_data[group_name]['this_year_all_people'] = group_info.get('all_people', [])
        all_attendance_data[group_name]['this_year_recent_missed'] = group_info.get('recent_missed', [])

    for group_name, group_info in last_year_attendance_data.items():
        if group_name not in all_attendance_data:
            all_attendance_data[group_name] = {
                'this_year_attendees': [], 'this_year_all_people': [],
                'this_year_recent_missed': [], 'last_year_attendees': [], 'last_year_all_people': []
            }
        all_attendance_data[group_name]['last_year_attendees'] = group_info.get('attendees', [])
        all_attendance_data[group_name]['last_year_all_people'] = group_info.get('all_people', [])

    print("\nDEBUG: Available attendance group names:")
    for group_name in sorted(all_attendance_data.keys()):
        print(f"   '{group_name}'")

    follow_up_this_year = {}
    follow_up_last_year = {}
    follow_up_recent_missed = {}

    for api_group_name, members in bible_study_members.items():
        print(f"\n   Analyzing {api_group_name}:")
        print(f"      Total members from API: {len(members)}")
        matching_attendance = None

        # Exact match and partial fallbacks
        for attendance_group_name in all_attendance_data.keys():
            if normalize_name(api_group_name) == normalize_name(attendance_group_name):
                matching_attendance = all_attendance_data[attendance_group_name]
                break
        if not matching_attendance:
            for attendance_group_name in all_attendance_data.keys():
                api_n = normalize_name(api_group_name)
                att_n = normalize_name(attendance_group_name)
                if (api_n in att_n) or (att_n in api_n) or any(w in att_n for w in api_n.split() if len(w) > 3):
                    matching_attendance = all_attendance_data[attendance_group_name]
                    break

        if not matching_attendance:
            print(f"      No matching attendance data found for '{api_group_name}'")
            continue

        this_year_attendees = matching_attendance.get('this_year_attendees', [])
        this_year_all_people = matching_attendance.get('this_year_all_people', [])
        this_year_recent_missed = matching_attendance.get('this_year_recent_missed', [])
        last_year_attendees = matching_attendance.get('last_year_attendees', [])
        last_year_all_people = matching_attendance.get('last_year_all_people', [])

        print(f"      This calendar year: {len(this_year_attendees)} / {len(this_year_all_people)} attended")
        print(f"      Recent missed (last 3): {len(this_year_recent_missed)}")
        print(f"      Last calendar year: {len(last_year_attendees)} / {len(last_year_all_people)} attended")

        group_followup_this_year = []
        group_followup_last_year = []
        group_followup_recent_missed = []

        for member in members:
            member_name = member['name']
            member_norm = normalize_name(member_name)

            missed_recent = any(normalize_name(p) == member_norm for p in this_year_recent_missed)
            if missed_recent:
                group_followup_recent_missed.append(member)

            appears_this_year = any(normalize_name(p) == member_norm for p in this_year_all_people)
            if appears_this_year:
                attended_this_year = any(normalize_name(p) == member_norm for p in this_year_attendees)
                if not attended_this_year:
                    group_followup_this_year.append(member)

            appears_last_year = any(normalize_name(p) == member_norm for p in last_year_all_people)
            if appears_last_year:
                attended_last_year = any(normalize_name(p) == member_norm for p in last_year_attendees)
                if not attended_last_year:
                    group_followup_last_year.append(member)

        if group_followup_recent_missed:
            follow_up_recent_missed[api_group_name] = group_followup_recent_missed
        if group_followup_this_year:
            follow_up_this_year[api_group_name] = group_followup_this_year
        if group_followup_last_year:
            follow_up_last_year[api_group_name] = group_followup_last_year

    print("\nStep 7: Creating visualizations...")
    create_charts(follow_up_this_year, follow_up_last_year, current_year, last_year)

    # ---- Step 7b UPDATED: create the 3-across PNG grid for the per-group progressive charts ----
    print("\nStep 7b: Creating progressive monthly attendance grid (PNG)...")
    create_progressive_attendance_charts(current_monthly_data, last_year_monthly_data, current_year, last_year)

    # Step 8: Detailed follow-up HTML list (unchanged)
    print("\nStep 8: Generating follow-up member list...")
    create_followup_member_list(follow_up_recent_missed, follow_up_this_year, follow_up_last_year, current_year, last_year)

    return follow_up_recent_missed, follow_up_this_year, follow_up_last_year

if __name__ == "__main__":
    print("=" * 80)
    print("BIBLE STUDY GROUP ATTENDANCE ANALYSIS - CALENDAR YEAR VERSION")
    print("   Reports on 'Last Calendar Year' and 'This Calendar Year'")
    print("   URGENT: People who missed the last 3 meetings (immediate follow-up)")
    print("   Follow-up list: Zero attendance this year first, then last year")
    print("   Bar charts show number of people per group needing follow-up")
    print("   Progressive charts show monthly attendance trends by category")
    print("=" * 80)
    
    try:
        result = create_bible_study_attendance_analysis()
        
        if result is not None:
            follow_up_recent_missed, follow_up_this_year, follow_up_last_year = result
            total_recent_missed = sum(len(members) for members in follow_up_recent_missed.values()) if follow_up_recent_missed else 0
            total_this_year = sum(len(members) for members in follow_up_this_year.values()) if follow_up_this_year else 0
            total_last_year = sum(len(members) for members in follow_up_last_year.values()) if follow_up_last_year else 0
            
            print("\nDETAILED MEMBER FOLLOW-UP LIST:")
            print(f"   URGENT - Missed last 3 meetings: {total_recent_missed} members")
            print(f"   This calendar year: {total_this_year} members need follow-up")
            print(f"   Last calendar year: {total_last_year} members need follow-up")
            
            print("\nANALYSIS COMPLETE!")
            print("Using REAL attendance data from your Elvanto reports!")
            print("Calendar year analysis: Last year vs This year!")
            print("URGENT: Recent missed meetings analysis (last 3 meetings)!")
            print("Priority-ordered follow-up list generated!")
            print("Bar charts show people count per group!")
            print("Progressive monthly attendance charts created!")
            print("IMPORTANT: Only includes people who appeared in reports but didn't attend")
            print("(People not in church during that period are excluded)")
            
            # List all created files with full paths
            import os
            current_dir = os.getcwd()
            print("\nFILES CREATED:")
            print(f"   Working Directory: {current_dir}")
            
            # List all HTML files in directory
            html_files = [f for f in os.listdir('.') if f.endswith('.html') and 'bible_study' in f]
            if html_files:
                print("   Generated Files:")
                for file in sorted(html_files):
                    file_path = os.path.join(current_dir, file)
                    file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                    print(f"      • {file} ({file_size} bytes)")
                    print(f"        Location: {file_path}")
                
                print("\nTO OPEN FILES:")
                print(f"   1. Navigate to: {current_dir}")
                print("   2. Double-click any .html file to open in your browser")
                print("   3. Or copy the full path and paste into your browser address bar")
                print("\nFILE DESCRIPTIONS:")
                print("   • *followup_calendar_years* = Side-by-side bar charts of follow-up needs")
                print("   • *followup_members* = Detailed member lists for follow-up")
                print("   • *progressive_attendance* = Monthly attendance trends comparison")
            else:
                print("   Warning: No Bible study HTML files found in current directory")
                print("   All files in directory:")
                all_files = os.listdir('.')
                for file in all_files:
                    if file.endswith(('.html', '.py')):
                        print(f"      • {file}")
            
        else:
            print("Analysis failed - please check your API key and report availability")
            
    except KeyboardInterrupt:
        print("\nAnalysis cancelled by user")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\nPress Enter to exit...")
    input()
