#!/usr/bin/env python3
"""
Selective Bible Study Group Dashboard
Modified from Code GP Dashboard to include only specific charts:
- Regular Bible Study Groups
- International Food and Friends (IFF) 
- Ever Attended Taste and See
- Youth Group
- Kids Club
- Buzz Music & Play (new)
"""

import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'plotly', 'requests', 'kaleido']
    for package in packages:
        try:
            if package == 'beautifulsoup4':
                import bs4
            elif package == 'kaleido':
                import kaleido
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

# Get API key from config file - following Code GP pattern exactly
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

def make_request(endpoint, params=None):
    """Make authenticated request to Elvanto API - Code GP pattern exactly"""
    if not API_KEY:
        print("No API key provided")
        return None
    
    url = f"{BASE_URL}{endpoint}.json"
    auth = (API_KEY, '')
    
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

def find_attendance_reports():
    """Find generic attendance report groups (Code GP pattern) - following Code GP exactly"""
    print("📋 Searching for generic attendance report groups...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None, None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_year_group = None
    last_year_group = None
    two_years_ago_group = None
    current_year = datetime.now().year
    last_year = current_year - 1
    two_years_ago = current_year - 2

    print(f"   Scanning {len(groups)} groups for generic attendance reports...")
    print(f"   Looking for: {current_year}, {last_year}, and {two_years_ago} data...")
    
    for group in groups:
        group_name = group.get('name', '').lower()
        
        # Look for current year report (Code GP pattern)
        if ('report of group individual attendance' in group_name and 
            'last year' not in group_name and 'two years ago' not in group_name):
            current_year_group = group
            print(f"✅ Found current year report: {group.get('name')}")
            
        # Look for last year report (Code GP pattern)  
        elif 'report of last year group individual attendance' in group_name:
            last_year_group = group
            print(f"✅ Found last year report: {group.get('name')}")
            
        # Look for two years ago report
        elif 'report of two years ago group individual attendance' in group_name:
            two_years_ago_group = group
            print(f"✅ Found two years ago report: {group.get('name')}")
    
    if not current_year_group and not last_year_group and not two_years_ago_group:
        print("❌ No generic attendance reports found")
        print("   Available groups containing 'report':")
        for group in groups:
            if 'report' in group.get('name', '').lower():
                print(f"      • {group.get('name')}")

    return current_year_group, last_year_group, two_years_ago_group

def download_attendance_report_data(group):
    """Download attendance data from a report group (Code GP pattern) - following Code GP exactly"""
    if not group:
        return None
        
    group_name = group.get('name', 'Unknown')
    print(f"📥 Getting attendance data from report: {group_name}")
    
    # Extract URL from group location fields (Code GP pattern)
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field]).strip()
            print(f"   Found report URL in {field}: {report_url}")
            break
    
    if not report_url:
        print(f"   ❌ No report URL found in group fields")
        return None
    
    try:
        print(f"   📡 Downloading report data...")
        response = requests.get(report_url, timeout=30)
        
        if response.status_code == 200:
            print(f"   ✅ Downloaded {len(response.text)} characters")
            return response.text
        else:
            print(f"   ❌ HTTP Error {response.status_code}")
            return None
            
    except Exception as e:
        print(f"   ❌ Download failed: {e}")
        return None

def categorize_group(group_name):
    """Categorize groups for the selective dashboard - updated for our 6 specific categories"""
    name_lower = group_name.lower()
    
    if 'kids club' in name_lower:
        return 'kids_club'
    elif 'youth group' in name_lower:
        return 'youth_group'
    elif 'iff' in name_lower or 'international food' in name_lower:
        return 'iff'
    elif 'ever attended' in name_lower and 'taste' in name_lower:
        return 'taste_and_see'
    elif ('buzz' in name_lower and 'music' in name_lower) or 'buzz music' in name_lower:
        return 'buzz_music'
    else:
        return 'regular_bible_studies'

def extract_monthly_attendance_data(html_content, year_label):
    """Extract monthly attendance data from HTML report - following Code GP pattern exactly"""
    if not html_content:
        return {}
    
    print(f"📊 Parsing attendance data for {year_label}...")
    print(f"   HTML content length: {len(html_content)} characters")
    
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    
    if not table:
        print("   ❌ No table found in report")
        return {}
    
    rows = table.find_all('tr')
    if not rows:
        print("   ❌ No rows found in table")
        return {}
    
    print(f"   Found table with {len(rows)} rows")
    
    # Extract header to find month columns
    header_row = rows[0]
    header_cells = header_row.find_all(['th', 'td'])
    date_columns = {}
    
    print(f"   Header row has {len(header_cells)} columns")
    for i, cell in enumerate(header_cells):
        cell_text = cell.get_text().strip()
        if '/' in cell_text:  # Date format like "1/1"
            try:
                month_day = cell_text.split('/')
                month = int(month_day[0])
                date_columns[i] = month
                print(f"   Column {i}: {cell_text} -> Month {month}")
            except:
                continue
    
    print(f"   Found {len(date_columns)} date columns: {list(date_columns.values())}")
    
    # Extract attendance data - following Code GP pattern exactly
    monthly_data = {}
    current_group = None
    groups_found = 0
    people_processed = 0
    
    for row_idx, row in enumerate(rows[1:], 1):  # Skip header
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        row_data = [cell.get_text().strip() for cell in cells]
        
        # Check if this is a group header row - Code GP pattern
        first_cell = cells[0]
        is_group_header = False
        
        style = first_cell.get('style', '')
        class_attr = first_cell.get('class', [])
        
        if ('background' in style.lower() and 'black' in style.lower()) or \
           any('header' in str(cls).lower() or 'group' in str(cls).lower() for cls in class_attr):
            is_group_header = True
        
        first_cell_text = row_data[0] if row_data else ""
        if any(keyword in first_cell_text.lower() for keyword in 
               ['bible study', 'youth group', 'ever attended', 'kids club', 'small group', 
                'home group', 'iff', 'international', 'buzz music', 'buzz']):
            is_group_header = True
        
        if is_group_header:
            current_group = first_cell_text
            if current_group not in monthly_data:
                monthly_data[current_group] = {}
                groups_found += 1
                print(f"   Found group {groups_found}: '{current_group}'")
            continue
        
        # Process individual attendance row - Code GP pattern
        if current_group and row_data:
            first_cell_value = row_data[0].strip()
            
            if not first_cell_value or not ',' in first_cell_value:
                continue
            
            people_processed += 1
            attendances_this_person = 0
            
            # Count attendances by month for this person
            for col_idx, month in date_columns.items():
                if col_idx < len(row_data):
                    cell = str(row_data[col_idx]).strip().upper()
                    if cell == 'Y':  # Attended
                        if month not in monthly_data[current_group]:
                            monthly_data[current_group][month] = 0
                        monthly_data[current_group][month] += 1
                        attendances_this_person += 1
            
            if row_idx <= 5:  # Show first few people for verification
                print(f"   Person {people_processed}: '{first_cell_value}' in '{current_group}' - {attendances_this_person} attendances")
    
    print(f"   ✅ REAL DATA EXTRACTED:")
    print(f"   Groups found: {groups_found}")
    print(f"   People processed: {people_processed}")
    print(f"   Total monthly data points: {sum(len(group_data) for group_data in monthly_data.values())}")
    
    # Show summary of data extracted
    for group_name, group_monthly in monthly_data.items():
        total_attendances = sum(group_monthly.values())
        print(f"   '{group_name}': {total_attendances} total attendances across {len(group_monthly)} months")
    
    return monthly_data

def create_selective_attendance_charts(current_monthly_data, last_year_monthly_data, current_year, last_year):
    """Create selective attendance charts with only requested groups - following Code GP pattern"""
    
    print("Creating selective monthly attendance charts...")
    
    # Initialize data structures for specific categories only
    categories = {
        'regular_bible_studies': {},
        'iff': {},
        'taste_and_see': {},
        'youth_group': {},
        'kids_club': {},
        'buzz_music': {}
    }
    
    # Process current year data - CODE GP PATTERN EXACTLY
    for group_name, monthly_counts in current_monthly_data.items():
        category = categorize_group(group_name)
        if category in categories:  # Only process our target categories
            for month, count in monthly_counts.items():
                if month not in categories[category]:
                    categories[category][month] = 0
                categories[category][month] += count
    
    # Process last year data - Code GP pattern
    last_year_categories = {
        'regular_bible_studies': {},
        'iff': {},
        'taste_and_see': {},
        'youth_group': {},
        'kids_club': {},
        'buzz_music': {}
    }
    
    for group_name, monthly_counts in last_year_monthly_data.items():
        category = categorize_group(group_name)
        if category in last_year_categories:  # Only process our target categories
            for month, count in monthly_counts.items():
                if month not in last_year_categories[category]:
                    last_year_categories[category][month] = 0
                last_year_categories[category][month] += count
    
    # Create cumulative data for each category - BACK TO CODE GP PATTERN
    def create_cumulative_data(monthly_data, year):
        months = list(range(1, 13)) if year == last_year else list(range(1, datetime.now().month + 1))
        cumulative = []
        running_total = 0
        
        for month in months:
            running_total += monthly_data.get(month, 0)
            cumulative.append(running_total)
        
        return months, cumulative
    
    # Calculate grid size for our 6 specific charts
    total_charts = 6
    cols = 2
    rows = 3  # 3 rows, 2 columns for 6 charts
    
    # Create subplot titles for our specific charts
    subplot_titles = [
        'Regular Bible Study Groups',
        'International Food & Friends (IFF)',
        'Ever Attended Taste and See',
        'Youth Group',
        'Kids Club',
        'Buzz Music & Play'
    ]
    
    # Create subplots - Code GP pattern
    fig = make_subplots(
        rows=rows, cols=cols,
        subplot_titles=subplot_titles,
        vertical_spacing=0.1,
        horizontal_spacing=0.1
    )
    
    # Chart configuration with improved colors for better contrast
    chart_configs = [
        {'category': 'regular_bible_studies', 'row': 1, 'col': 1, 'color': '#1e40af'},  # Deep blue
        {'category': 'iff', 'row': 1, 'col': 2, 'color': '#059669'},                   # Deep emerald
        {'category': 'taste_and_see', 'row': 2, 'col': 1, 'color': '#b45309'},        # Deep amber (better than orange)
        {'category': 'youth_group', 'row': 2, 'col': 2, 'color': '#dc2626'},          # Deep red
        {'category': 'kids_club', 'row': 3, 'col': 1, 'color': '#7c3aed'},            # Deep violet
        {'category': 'buzz_music', 'row': 3, 'col': 2, 'color': '#be185d'}            # Deep rose
    ]
    
    # Month labels for x-axis
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # Add traces for each category - CODE GP PATTERN
    for config in chart_configs:
        category = config['category']
        row = config['row']
        col = config['col']
        color = config['color']
        
        # Create cumulative data for this category
        if categories[category]:
            this_months, this_cumulative = create_cumulative_data(categories[category], current_year)
            this_month_labels = [month_names[m-1] for m in this_months]
            
            # Add current year line - Enhanced styling
            fig.add_trace(
                go.Scatter(
                    x=this_month_labels,
                    y=this_cumulative,
                    mode='lines+markers',
                    line=dict(color=color, width=4, shape='spline'),
                    marker=dict(size=10, color=color, line=dict(width=2, color='white')),
                    name=f'{current_year}',
                    showlegend=(row == 1 and col == 1),
                    legendgroup='thisyear'
                ),
                row=row, col=col
            )
        
        # Add last year data if available
        if last_year_categories[category]:
            last_months, last_cumulative = create_cumulative_data(last_year_categories[category], last_year)
            last_month_labels = [month_names[m-1] for m in last_months]
            
            # Convert hex to rgba for transparency
            def hex_to_rgba(hex_color, alpha=0.5):
                hex_color = hex_color.lstrip('#')
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16) 
                b = int(hex_color[4:6], 16)
                return f'rgba({r},{g},{b},{alpha})'
            
            last_year_color = hex_to_rgba(color, 0.6)
            
            fig.add_trace(
                go.Scatter(
                    x=last_month_labels,
                    y=last_cumulative,
                    mode='lines+markers',
                    line=dict(color=last_year_color, width=3, dash='dash', shape='spline'),
                    marker=dict(size=8, color=last_year_color, line=dict(width=1, color='white')),
                    name=f'{last_year}',
                    showlegend=(row == 1 and col == 1),
                    legendgroup='lastyear'
                ),
                row=row, col=col
            )
    
    # Enhanced layout with professional styling
    fig.update_layout(
        title=dict(
            text=f'<b>Bible Study Group Attendance Dashboard</b><br><sub>Calendar Years {last_year} & {current_year}</sub>',
            x=0.5,
            font=dict(size=24, color='#1f2937')
        ),
        showlegend=True,
        height=420 * rows,
        width=1600,
        font=dict(size=12, family="Arial, sans-serif"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="center",
            x=0.5,
            font=dict(size=14),
            bgcolor="rgba(255,255,255,0.8)",
            bordercolor="rgba(0,0,0,0.2)",
            borderwidth=1
        ),
        plot_bgcolor='white',
        paper_bgcolor='#f8fafc',  # Light gray background
        margin=dict(t=120, b=60, l=60, r=60)
    )
    
    # Enhanced axes styling
    for i in range(1, rows + 1):
        for j in range(1, cols + 1):
            fig.update_yaxes(
                title_text="Cumulative Attendance", 
                row=i, col=j, 
                title_font_size=14,
                gridcolor='rgba(0,0,0,0.1)',
                zerolinecolor='rgba(0,0,0,0.2)',
                tickfont=dict(size=12)
            )
            fig.update_xaxes(
                title_text="Month", 
                row=i, col=j, 
                title_font_size=14,
                gridcolor='rgba(0,0,0,0.1)',
                tickfont=dict(size=12)
            )
    
    # Save chart as PNG with high quality
    filename = f"selective_bible_study_attendance_{current_year}_{last_year}.png"
    current_dir = os.getcwd()
    full_path = os.path.join(current_dir, filename)
    
    try:
        # Save as high-quality PNG
        fig.write_image(filename, format="png", width=1600, height=420*rows, scale=2)
        print(f"Selective charts saved successfully as PNG!")
        print(f"   Total charts created: 6 specific charts")
        print(f"   Directory: {current_dir}")
        print(f"   Filename: {filename}")
        print(f"   Full path: {full_path}")
        print(f"   File exists: {os.path.exists(full_path)}")
        
        if os.path.exists(full_path):
            file_size = os.path.getsize(full_path)
            print(f"   File size: {file_size:,} bytes ({file_size/1024/1024:.1f} MB)")
            
        # Try to open the PNG file with default image viewer
        try:
            if sys.platform.startswith('darwin'):  # macOS
                os.system(f'open "{full_path}"')
            elif sys.platform.startswith('win'):   # Windows
                os.system(f'start "" "{full_path}"')
            else:  # Linux
                os.system(f'xdg-open "{full_path}"')
            print(f"   ✅ Opened PNG chart with default image viewer!")
        except Exception as e:
            print(f"   ⚠️ Could not auto-open image: {e}")
            print(f"   💡 Manually open: {full_path}")
            
    except Exception as e:
        print(f"Error saving PNG charts: {e}")
        print("   💡 Note: PNG export requires 'kaleido' package")
        print("   💡 Fallback: Saving as HTML instead...")
        try:
            html_filename = filename.replace('.png', '.html')
            fig.write_html(html_filename)
            print(f"   ✅ Saved as HTML: {html_filename}")
        except Exception as html_error:
            print(f"   ❌ HTML fallback also failed: {html_error}")
        import traceback
        traceback.print_exc()

def create_selective_dashboard():
    """Main function to create the selective dashboard - following Code GP pattern exactly"""
    
    print("=" * 80)
    print("SELECTIVE BIBLE STUDY GROUP DASHBOARD")
    print("   Charts: Regular Bible Study Groups, IFF, Taste & See,")
    print("           Youth Group, Kids Club, Buzz Music & Play")
    print("   Following Code GP Dashboard pattern exactly")
    print("=" * 80)
    
    current_year = datetime.now().year
    last_year = current_year - 1
    
    # Step 1: Find attendance report groups - Code GP pattern
    print(f"\nStep 1: Finding attendance reports...")
    current_year_group, last_year_group, two_years_ago_group = find_attendance_reports()
    
    if not current_year_group:
        print("❌ No current year report found")
        return None
    
    # Step 2: Download report data - Code GP pattern
    print(f"\nStep 2: Downloading attendance data from groups...")
    current_year_data = download_attendance_report_data(current_year_group)
    
    if not current_year_data:
        print("Failed to download current year attendance data")
        return None
    
    # Download last year data if available - Code GP pattern
    last_year_data = None
    if last_year_group:
        last_year_data = download_attendance_report_data(last_year_group)
        print("Downloaded last year attendance data")
    else:
        print("No last year report group found, continuing with current year only")
        
    # Step 3: Extract attendance data from both reports - Code GP pattern
    current_monthly_data = extract_monthly_attendance_data(current_year_data, "Current Year")
    
    if not current_monthly_data:
        print("No attendance data extracted from current year report")
        return None
    
    if last_year_data:
        last_year_monthly_data = extract_monthly_attendance_data(last_year_data, "Last Year")
    else:
        last_year_monthly_data = {}
    
    # Step 4: Create selective visualizations - Code GP pattern
    print("\nStep 4: Creating selective visualizations...")
    create_selective_attendance_charts(current_monthly_data, last_year_monthly_data, current_year, last_year)
    
    print("\n✅ SELECTIVE DASHBOARD COMPLETE!")
    print(f"   Created charts for: Regular Bible Studies, IFF, Taste & See,")
    print(f"                      Youth Group, Kids Club, Buzz Music & Play")
    
    return current_monthly_data, last_year_monthly_data

if __name__ == "__main__":
    result = create_selective_dashboard()
    
    if result is not None:
        current_data, last_year_data = result
        print(f"\nData summary:")
        print(f"   Current year groups: {len(current_data)} groups")
        print(f"   Last year groups: {len(last_year_data)} groups")
        
        # Show which groups were found for each category
        categories_found = {}
        for group_name in current_data.keys():
            category = categorize_group(group_name)
            if category not in categories_found:
                categories_found[category] = []
            categories_found[category].append(group_name)
        
        print(f"\nGroups found by category:")
        category_names = {
            'regular_bible_studies': 'Regular Bible Study Groups',
            'iff': 'International Food & Friends',
            'taste_and_see': 'Ever Attended Taste and See',
            'youth_group': 'Youth Group',
            'kids_club': 'Kids Club',
            'buzz_music': 'Buzz Music & Play'
        }
        
        for category, display_name in category_names.items():
            groups = categories_found.get(category, [])
            print(f"   {display_name}: {len(groups)} groups")
            for group in groups:
                print(f"      - {group}")
