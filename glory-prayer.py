import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'pandas', 'plotly', 'requests', 'kaleido']
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
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from bs4 import BeautifulSoup
import re
import webbrowser
import os

print("✨ ENHANCED BRINGING GLORY TO JESUS CHRIST DASHBOARDS ✨")
print("="*60)
print("🙏 Prayer Meetings Analysis + 🍽️ Newcomers Lunch Progress")
print("Now with stunning visuals and 6-week moving averages!")

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

BASE_URL = "https://api.elvanto.com/v1"

def make_request(endpoint, params=None):
    """Make API request to Elvanto"""
    try:
        response = requests.post(f"{BASE_URL}/{endpoint}.json", 
                               auth=(API_KEY, ''), 
                               json=params or {}, 
                               timeout=30)
        if response.status_code == 200:
            data = response.json()
            if data.get('status') == 'ok':
                return data
            else:
                print(f"API Error: {data.get('error', {}).get('message', 'Unknown error')}")
                return None
        else:
            print(f"HTTP Error: {response.status_code}")
            return None
    except Exception as e:
        print(f"Network Error: {e}")
        return None

def find_service_attendance_reports():
    """Find both current year and last year service attendance reports"""
    print("\n📋 Searching for service attendance reports...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_year_group = None
    last_year_group = None

    for group in groups:
        group_name = group.get('name', '').lower()
        if 'report of last year service individual attendance' in group_name:
            last_year_group = group
            print(f"✅ Found last year service attendance report: {group.get('name')}")
        elif ('report of service individual attendance' in group_name and 
              'last year' not in group_name):
            current_year_group = group
            print(f"✅ Found current year service attendance report: {group.get('name')}")

    if not current_year_group:
        print("❌ Current year service attendance report not found")
    if not last_year_group:
        print("❌ Last year service attendance report not found")

    return current_year_group, last_year_group

def find_report_group(group_name_keywords):
    """Find a report group by keywords in the name"""
    print(f"\n📋 Searching for report group with keywords: {group_name_keywords}...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    # Look for group matching the keywords
    for group in groups:
        group_name = group.get('name', '').lower()
        if any(keyword.lower() in group_name for keyword in group_name_keywords):
            print(f"✅ Found report group: {group.get('name')}")
            return group

    print(f"❌ No group found with keywords: {group_name_keywords}")
    return None

def extract_service_attendance_data(group, year_label="Service Attendance"):
    """Extract service attendance data (including prayer meetings)"""
    print(f"\n🔍 Extracting {year_label} data...")

    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print("❌ No URL found in service attendance group")
        return None, None

    print(f"✅ Found service attendance report URL")

    # Fetch and parse HTML
    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find attendance table
        tables = soup.find_all('table')
        for table in tables:
            header_row = table.find('tr')
            if header_row:
                headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
                header_text = ' '.join(headers).lower()

                if 'first name' in header_text and ('attended' in header_text or any('/' in h for h in headers)):
                    print(f"✅ Found service attendance table")

                    attendance_records = []
                    for row in table.find_all('tr')[1:]:
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= len(headers):
                            row_data = {}
                            for i, header in enumerate(headers):
                                if i < len(cells):
                                    row_data[header] = cells[i].get_text(strip=True)
                            if any(row_data.values()):
                                attendance_records.append(row_data)

                    print(f"✅ Extracted {len(attendance_records)} service attendance records")
                    return pd.DataFrame(attendance_records), headers

        print("❌ Service attendance table not found")
        return None, None

    except Exception as e:
        print(f"❌ Error extracting service attendance data: {e}")
        return None, None

def parse_prayer_meeting_columns(headers, target_year, year_label):
    """Parse service columns and extract Prayer Meeting services, separating weekly from quarterly"""
    print(f"\n🙏 Parsing {year_label} service columns for Prayer Meetings...")

    today = datetime.now().date()
    weekly_prayer_services = []
    quarterly_prayer_services = []
    unparseable_count = 0
    future_count = 0
    wrong_year_count = 0
    non_saturday_count = 0
    non_prayer_count = 0

    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue

        # Look for prayer meetings in the header
        is_weekly_prayer = 'weekly prayer meeting' in header_lower
        is_quarterly_prayer = 'quarterly prayer meeting' in header_lower
        is_general_prayer = 'prayer meeting' in header_lower and not is_weekly_prayer and not is_quarterly_prayer

        if not (is_weekly_prayer or is_quarterly_prayer or is_general_prayer):
            non_prayer_count += 1
            continue

        # Parse service header (looking for date patterns)
        date_match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', header)
        if not date_match:
            date_match = re.search(r'(\d{1,2})/(\d{1,2})(?:\s|$)', header)
            if not date_match:
                unparseable_count += 1
                continue

        try:
            day = int(date_match.group(1))
            month = int(date_match.group(2))

            # Determine year
            if len(date_match.groups()) >= 3 and date_match.group(3):
                year = int(date_match.group(3))
            else:
                year = target_year

            # CRITICAL: Ensure we're only processing services from the target year
            if year != target_year:
                wrong_year_count += 1
                continue

            service_date = datetime(year, month, day)
            service_date_only = service_date.date()

            # For current year, only include past services
            if year == datetime.now().year and service_date_only > today:
                future_count += 1
                continue

            # Check if it's a Saturday (Prayer meetings are typically on Saturday)
            if service_date.weekday() != 5:  # Saturday = 5
                non_saturday_count += 1
                continue

            prayer_service = {
                'header': header,
                'date': service_date,
                'date_only': service_date_only,
                'type': 'quarterly' if is_quarterly_prayer else 'weekly'
            }

            if is_quarterly_prayer:
                quarterly_prayer_services.append(prayer_service)
            else:
                weekly_prayer_services.append(prayer_service)

        except ValueError:
            unparseable_count += 1
            continue

    # Sort by date
    weekly_prayer_services.sort(key=lambda x: x['date'])
    quarterly_prayer_services.sort(key=lambda x: x['date'])

    print(f"✅ Found {len(weekly_prayer_services)} Saturday Weekly Prayer services")
    print(f"✅ Found {len(quarterly_prayer_services)} Saturday Quarterly Prayer services")
    print(f"   Excluded {non_prayer_count} non-prayer columns")
    print(f"   Excluded {non_saturday_count} non-Saturday prayer meetings")
    print(f"   Excluded {wrong_year_count} wrong-year services")
    print(f"   Excluded {future_count} future services")
    print(f"   Couldn't parse {unparseable_count} headers")

    return weekly_prayer_services, quarterly_prayer_services

def calculate_prayer_meeting_attendance(df, prayer_services):
    """Calculate attendance for prayer meetings"""
    print(f"\n📊 Calculating Prayer Meeting attendance...")

    attendance_data = []

    for service in prayer_services:
        header = service['header']
        date = service['date']

        if header in df.columns:
            # Count people who attended ('Y' in the column)
            attendees = len(df[df[header] == 'Y'])
            
            attendance_data.append({
                'date': date,
                'attendance': attendees
            })

            print(f"   🙏 {date.strftime('%d %b %Y')}: {attendees} attendees")

    # Sort by date
    attendance_data.sort(key=lambda x: x['date'])
    
    print(f"✅ Calculated attendance for {len(attendance_data)} prayer meetings")
    return attendance_data

def extract_newcomers_lunch_data_from_report(group):
    """Extract newcomers lunch data from shared report"""
    print(f"\n🍽️ Extracting Newcomers Lunch data from shared report...")

    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print("❌ No URL found in Newcomers Lunch group")
        return []

    print(f"✅ Found Newcomers Lunch report URL")

    # Fetch and parse HTML
    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find the lunch attendance table
        tables = soup.find_all('table')
        for table in tables:
            header_row = table.find('tr')
            if header_row:
                headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
                header_text = ' '.join(headers).lower()

                # Look for table with Date and attendance columns
                if ('date' in header_text and 
                    ('attended' in header_text or 'members' in header_text)):
                    print(f"✅ Found Newcomers Lunch table")
                    print(f"📋 Headers: {headers}")

                    lunch_records = []
                    for row in table.find_all('tr')[1:]:  # Skip header
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= len(headers):
                            row_data = {}
                            for i, header in enumerate(headers):
                                if i < len(cells):
                                    cell_text = cells[i].get_text(strip=True)
                                    row_data[header] = cell_text

                            # Skip summary rows (Total, Average)
                            date_text = row_data.get('Date', '').strip()
                            if date_text and date_text.lower() not in ['total', 'average', '']:
                                lunch_records.append(row_data)

                    print(f"✅ Extracted {len(lunch_records)} lunch records")

                    # Convert to our expected format
                    lunch_data = []
                    for record in lunch_records:
                        date_str = record.get('Date', '')
                        attendance_str = (record.get('Members Attended', '') or 
                                        record.get('Attended', '') or 
                                        record.get('Attendance', ''))

                        # Parse date
                        try:
                            # Handle various date formats
                            date_formats = [
                                '%d %B, %Y',      # "8 September, 2024"
                                '%d %b, %Y',      # "8 Sep, 2024"
                                '%d/%m/%Y',       # "8/9/2024"
                                '%Y-%m-%d',       # "2024-09-08"
                                '%d %B %Y',       # "8 September 2024"
                                '%d %b %Y',       # "8 Sep 2024"
                                '%B %d, %Y',      # "September 8, 2024"
                                '%b %d, %Y'       # "Sep 8, 2024"
                            ]

                            lunch_date = None
                            for date_format in date_formats:
                                try:
                                    lunch_date = datetime.strptime(date_str, date_format)
                                    break
                                except:
                                    continue

                            if not lunch_date:
                                print(f"⚠️ Could not parse date: {date_str}")
                                continue

                            # Parse attendance
                            try:
                                attendance = int(attendance_str) if attendance_str.isdigit() else 0
                            except:
                                attendance = 0

                            lunch_data.append({
                                'date': lunch_date,
                                'attendance': attendance
                            })

                            print(f"   📅 {lunch_date.strftime('%d %b %Y')}: {attendance} attendees")

                        except Exception as e:
                            print(f"⚠️ Error processing record: {record} - {e}")
                            continue

                    # Sort by date
                    lunch_data.sort(key=lambda x: x['date'])
                    print(f"✅ Successfully processed {len(lunch_data)} lunch dates")
                    return lunch_data

        print("❌ Newcomers Lunch table not found in report")
        return []

    except Exception as e:
        print(f"❌ Error extracting Newcomers Lunch data: {e}")
        return []

def calculate_rolling_average(data, window=6):
    """Calculate rolling average with specified window size (NOW USING 6 WEEKS!)"""
    if len(data) < window:
        return data

    rolling_data = []
    for i in range(len(data)):
        if i < window - 1:
            # For early points, use expanding average
            avg = sum(data[:i+1]) / (i+1)
        else:
            # Use rolling window
            avg = sum(data[i-window+1:i+1]) / window
        rolling_data.append(avg)

    return rolling_data

def calculate_cumulative_lunch_attendance(lunch_data, current_year):
    """Calculate cumulative attendance for newcomers lunch throughout the year"""
    current_year_lunch = [record for record in lunch_data if record['date'].year == current_year]
    
    if not current_year_lunch:
        return []
    
    # Sort by date to ensure proper cumulative calculation
    current_year_lunch.sort(key=lambda x: x['date'])
    
    cumulative_data = []
    running_total = 0
    
    for record in current_year_lunch:
        running_total += record['attendance']
        cumulative_data.append({
            'date': record['date'],
            'attendance': record['attendance'],
            'cumulative': running_total
        })
        print(f"   📅 {record['date'].strftime('%d %b %Y')}: +{record['attendance']} attendees (total: {running_total})")
    
    return cumulative_data

def calculate_prorated_cumulative_target(annual_target, baseline):
    """Calculate prorated cumulative target for newcomers lunch"""
    now = datetime.now()
    start_of_year = datetime(now.year, 1, 1)
    end_of_year = datetime(now.year, 12, 31)
    
    # Calculate how much of the year has passed
    days_elapsed = (now - start_of_year).days
    total_days_in_year = (end_of_year - start_of_year).days + 1
    year_progress = days_elapsed / total_days_in_year
    
    # For cumulative target, we expect to reach the annual target by end of year
    # So prorated target is just year_progress * annual_target
    cumulative_target = annual_target * year_progress
    
    return cumulative_target, year_progress

def create_stunning_prayer_dashboard(current_prayer_data, last_year_prayer_data, current_weekly_services, current_quarterly_services, last_year_weekly_services, last_year_quarterly_services, current_year, last_year):
    """Create a stunning Prayer Meeting dashboard with separate zoom and quarterly charts"""
    print("\n✨ Creating STUNNING Prayer Meeting Dashboard with separate Zoom & Quarterly charts...")

    # Enhanced modern color palette
    colors = {
        'current_year': '#6366f1',     # Indigo
        'current_smooth': '#8b5cf6',   # Purple
        'last_year': '#ef4444',        # Red
        'last_smooth': '#f97316',      # Orange
        'target': '#10b981',           # Emerald
        'quarterly_current': '#06b6d4', # Cyan
        'quarterly_last': '#f59e0b',    # Amber
        'background': '#0f172a',       # Dark blue
        'grid': 'rgba(148, 163, 184, 0.1)',
        'text': '#f1f5f9'
    }

    # Create subplots: top for zoom meetings, bottom for quarterly
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=[
            "<b style='color:#f1f5f9'>🔗 Weekly Zoom Prayer Meetings</b>",
            "<b style='color:#f1f5f9'>📅 Quarterly Prayer Meetings</b>"
        ],
        vertical_spacing=0.12,
        row_heights=[0.65, 0.35]
    )

    # Step 1: Create filtered zoom prayer data (exclude quarterly meeting weeks)
    def filter_zoom_prayer_data(prayer_data, quarterly_services):
        """Filter out prayer data from weeks with quarterly meetings"""
        if not quarterly_services:
            return prayer_data
        
        quarterly_weeks = set()
        for q_service in quarterly_services:
            # Get the week start (Monday) for each quarterly service
            q_date = q_service['date']
            week_start = q_date - timedelta(days=q_date.weekday())
            quarterly_weeks.add(week_start)
        
        filtered_data = []
        for record in prayer_data:
            record_date = record['date']
            record_week_start = record_date - timedelta(days=record_date.weekday())
            if record_week_start not in quarterly_weeks:
                filtered_data.append(record)
        
        return filtered_data

    # Filter current year zoom data
    current_zoom_data = []
    if current_prayer_data and current_quarterly_services:
        current_zoom_data = filter_zoom_prayer_data(current_prayer_data, current_quarterly_services)
        print(f"   🔗 Current year zoom meetings: {len(current_zoom_data)} (filtered from {len(current_prayer_data)} total)")
    elif current_prayer_data:
        current_zoom_data = current_prayer_data
        print(f"   🔗 Current year zoom meetings: {len(current_zoom_data)} (no quarterly data to filter)")

    # Filter last year zoom data
    last_year_zoom_data = []
    if last_year_prayer_data and last_year_quarterly_services:
        last_year_zoom_data = filter_zoom_prayer_data(last_year_prayer_data, last_year_quarterly_services)
        print(f"   🔗 Last year zoom meetings: {len(last_year_zoom_data)} (filtered from {len(last_year_prayer_data)} total)")
    elif last_year_prayer_data:
        last_year_zoom_data = last_year_prayer_data
        print(f"   🔗 Last year zoom meetings: {len(last_year_zoom_data)} (no quarterly data to filter)")

    # Calculate 10% growth target from last year's zoom average
    zoom_target = None
    if last_year_zoom_data:
        last_year_values = [record['attendance'] for record in last_year_zoom_data]
        last_year_smooth = calculate_rolling_average(last_year_values, window=6)
        last_year_avg = sum(last_year_smooth) / len(last_year_smooth)
        zoom_target = last_year_avg * 1.10  # 10% growth target
        print(f"🎯 Zoom Prayer Target: {zoom_target:.1f} (10% growth from {last_year} 6-week avg: {last_year_avg:.1f})")

    # TOP CHART: Add last year zoom data with 6-week moving average
    if last_year_zoom_data:
        last_year_df = pd.DataFrame(last_year_zoom_data)
        last_year_df['date'] = pd.to_datetime(last_year_df['date'])
        # Normalize dates to current year for comparison
        last_year_df['normalized_date'] = last_year_df['date'].apply(
            lambda x: datetime(current_year, x.month, x.day)
        )
        
        # Calculate 6-week moving average
        last_year_rolling = calculate_rolling_average(last_year_df['attendance'].tolist(), window=6)
        
        # Raw data line (subtle)
        fig.add_trace(
            go.Scatter(
                x=last_year_df['normalized_date'],
                y=last_year_df['attendance'],
                mode='lines',
                name=f'{last_year} Raw',
                line=dict(color=colors['last_year'], width=1, dash='dot'),
                opacity=0.4,
                hovertemplate=f'<b>%{{x|%a %d %b}} {last_year}</b><br>Zoom Attendance: %{{y}}<br><extra></extra>',
                showlegend=True
            ),
            row=1, col=1
        )
        
        # Smoothed trend line (main focus)
        fig.add_trace(
            go.Scatter(
                x=last_year_df['normalized_date'],
                y=last_year_rolling,
                mode='lines',
                name=f'{last_year} Trend',
                line=dict(color=colors['last_smooth'], width=5, shape='spline'),
                hovertemplate=f'<b>%{{x|%a %d %b}} {last_year}</b><br>6-Week Average: %{{y:.1f}}<br><extra></extra>',
                showlegend=True
            ),
            row=1, col=1
        )

    # TOP CHART: Add current year zoom data with 6-week moving average
    if current_zoom_data:
        current_year_df = pd.DataFrame(current_zoom_data)
        current_year_df['date'] = pd.to_datetime(current_year_df['date'])
        
        # Calculate 6-week moving average
        current_year_rolling = calculate_rolling_average(current_year_df['attendance'].tolist(), window=6)
        
        # Raw data line (subtle)
        fig.add_trace(
            go.Scatter(
                x=current_year_df['date'],
                y=current_year_df['attendance'],
                mode='lines',
                name=f'{current_year} Raw',
                line=dict(color=colors['current_year'], width=1, dash='dot'),
                opacity=0.4,
                hovertemplate=f'<b>%{{x|%a %d %b}} {current_year}</b><br>Zoom Attendance: %{{y}}<br><extra></extra>',
                showlegend=True
            ),
            row=1, col=1
        )
        
        # Smoothed trend line (main focus)
        fig.add_trace(
            go.Scatter(
                x=current_year_df['date'],
                y=current_year_rolling,
                mode='lines',
                name=f'{current_year} Trend',
                line=dict(color=colors['current_smooth'], width=5, shape='spline'),
                hovertemplate=f'<b>%{{x|%a %d %b}} {current_year}</b><br>6-Week Average: %{{y:.1f}}<br><extra></extra>',
                showlegend=True
            ),
            row=1, col=1
        )
        
        # Add 10% growth target line for zoom meetings
        if zoom_target:
            fig.add_hline(
                y=zoom_target,
                line=dict(color=colors['target'], width=3, dash='dash'),
                annotation_text=f"🎯 Zoom Target: {zoom_target:.1f}",
                annotation_position="top right",
                annotation=dict(
                    font=dict(size=12, color=colors['target'], family="Inter"),
                    bgcolor="rgba(16, 185, 129, 0.1)",
                    bordercolor=colors['target'],
                    borderwidth=2,
                    borderpad=6
                ),
                row=1, col=1
            )

    # Step 2: Create quarterly comparison bar chart
    def create_quarterly_data_from_attendance(prayer_data, quarterly_services, year):
        """Convert quarterly services to quarterly summary, but ONLY for meetings with actual attendance"""
        quarterly_data = {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0}
        
        # Only populate quarters that have actual attendance records
        if quarterly_services and prayer_data:
            for service in quarterly_services:
                # Ensure the service is from the correct year
                if service['date'].year != year:
                    print(f"   ⚠️ Skipping {service['date'].strftime('%d %b %Y')} - wrong year (expected {year})")
                    continue
                
                # Find ALL attendance records for this service date (there might be duplicates)
                matching_records = []
                for record in prayer_data:
                    if (record['date'].date() == service['date'].date() and 
                        record['date'].year == year):
                        matching_records.append(record)
                
                if matching_records:
                    # If multiple records for same date, pick the one with highest attendance
                    # (quarterly meeting vs cancelled weekly meeting)
                    best_record = max(matching_records, key=lambda x: x['attendance'])
                    quarter = f"Q{(service['date'].month - 1) // 3 + 1}"
                    quarterly_data[quarter] = best_record['attendance']
                    
                    if len(matching_records) > 1:
                        attendances = [r['attendance'] for r in matching_records]
                        print(f"   📅 {year} {quarter} ({service['date'].strftime('%d %b')}): Found {len(matching_records)} records {attendances}, using highest: {best_record['attendance']}")
                    else:
                        print(f"   📅 {year} {quarter} ({service['date'].strftime('%d %b')}): {best_record['attendance']} attendees")
                else:
                    print(f"   ⚠️ No attendance record found for {year} {service['date'].strftime('%d %b %Y')}")
        
        return quarterly_data

    # Get quarterly data for both years (only with actual attendance)
    current_quarterly_data = create_quarterly_data_from_attendance(
        current_prayer_data, current_quarterly_services, current_year
    )
    last_year_quarterly_data = create_quarterly_data_from_attendance(
        last_year_prayer_data, last_year_quarterly_services, last_year
    )

    quarters = ['Q1', 'Q2', 'Q3', 'Q4']
    
    # BOTTOM CHART: Add quarterly bars (properly grouped by quarter)
    fig.add_trace(
        go.Bar(
            x=quarters,
            y=[last_year_quarterly_data[q] for q in quarters],
            name=f'{last_year} Quarterly',
            marker=dict(color=colors['quarterly_last'], line=dict(width=1, color='white')),
            hovertemplate=f'<b>%{{x}} {last_year}</b><br>Attendance: %{{y}}<br><extra></extra>',
            showlegend=True
        ),
        row=2, col=1
    )
    
    fig.add_trace(
        go.Bar(
            x=quarters,
            y=[current_quarterly_data[q] for q in quarters],
            name=f'{current_year} Quarterly',
            marker=dict(color=colors['quarterly_current'], line=dict(width=1, color='white')),
            hovertemplate=f'<b>%{{x}} {current_year}</b><br>Attendance: %{{y}}<br><extra></extra>',
            showlegend=True
        ),
        row=2, col=1
    )

    # Stunning layout with dark theme
    fig.update_layout(
        title=dict(
            text=f"<b style='font-size:26px; color:#f1f5f9'>🙏 Prayer Meeting Analysis Dashboard</b><br><span style='font-size:14px; color:#94a3b8'>Weekly Zoom Meetings vs Quarterly In-Person Gatherings • {last_year} vs {current_year}</span>",
            x=0.5,
            y=0.97,
            font=dict(family="Inter, system-ui, sans-serif")
        ),
        plot_bgcolor=colors['background'],
        paper_bgcolor=colors['background'],
        font=dict(family="Inter, system-ui, sans-serif", size=12, color=colors['text']),
        height=1000,
        width=1400,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.12,
            xanchor="center",
            x=0.5,
            font=dict(size=12, color=colors['text']),
            bgcolor="rgba(15, 23, 42, 0.8)",
            bordercolor="rgba(148, 163, 184, 0.3)",
            borderwidth=1
        ),
        margin=dict(l=80, r=80, t=120, b=100),
        hovermode='x unified',
        barmode='group'  # Enable grouped bars for the quarterly chart
    )

    # Beautiful axes for top chart (zoom meetings)
    fig.update_xaxes(
        range=[datetime(current_year, 1, 1), datetime(current_year, 12, 31)],
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=11, color=colors['text']),
        title=dict(text="<b>Calendar Year</b>", font=dict(size=12, color=colors['text'])),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True,
        tickformat='%b',
        dtick='M1',
        row=1, col=1
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=11, color=colors['text']),
        title=dict(text="<b>Zoom Attendees</b>", font=dict(size=12, color=colors['text'])),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True,
        row=1, col=1
    )

    # Beautiful axes for bottom chart (quarterly meetings)
    fig.update_xaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=11, color=colors['text']),
        title=dict(text="<b>Quarter</b>", font=dict(size=12, color=colors['text'])),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True,
        row=2, col=1
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=11, color=colors['text']),
        title=dict(text="<b>Quarterly Attendees</b>", font=dict(size=12, color=colors['text'])),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True,
        row=2, col=1
    )

    # Save the stunning prayer dashboard
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prayer_filename = f"stunning_prayer_dashboard_{timestamp}.html"

    try:
        fig.write_html(prayer_filename)
        print(f"\n✨ Saved STUNNING Prayer Dashboard: {prayer_filename}")
    except Exception as e:
        print(f"\n❌ Failed to save Prayer Dashboard: {e}")

    try:
        prayer_png = f"stunning_prayer_dashboard_{timestamp}.png"
        fig.write_image(prayer_png, width=1400, height=1000, scale=3)
        print(f"✨ Saved Prayer PNG: {prayer_png}")
    except Exception as e:
        print(f"⚠️ Prayer PNG save failed: {e}")

    return prayer_filename, fig

def create_stunning_newcomers_dashboard(lunch_data, current_year):
    """Create a stunning Newcomers Lunch dashboard"""
    print("\n✨ Creating STUNNING Newcomers Lunch Dashboard...")

    # Enhanced modern color palette for newcomers
    colors = {
        'cumulative': '#06b6d4',      # Cyan
        'individual': '#8b5cf6',      # Purple
        'target': '#f59e0b',          # Amber
        'background': '#0c1120',      # Dark navy
        'grid': 'rgba(148, 163, 184, 0.1)',
        'text': '#f1f5f9'
    }

    # Create the figure with dark theme
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=[
            "<b style='color:#f1f5f9'>📈 Cumulative Progress</b>",
            "<b style='color:#f1f5f9'>📊 Individual Events</b>"
        ],
        horizontal_spacing=0.12
    )

    if lunch_data:
        cumulative_lunch = calculate_cumulative_lunch_attendance(lunch_data, current_year)
        
        if cumulative_lunch:
            cumulative_df = pd.DataFrame(cumulative_lunch)
            cumulative_df['date'] = pd.to_datetime(cumulative_df['date'])
            
            # Left chart: Cumulative progress with gradient fill
            fig.add_trace(
                go.Scatter(
                    x=cumulative_df['date'],
                    y=cumulative_df['cumulative'],
                    mode='lines+markers',
                    name='Cumulative Attendees',
                    line=dict(color=colors['cumulative'], width=4),
                    marker=dict(size=10, symbol='circle', line=dict(width=2, color='white')),
                    fill='tonexty',
                    fillcolor='rgba(6, 182, 212, 0.3)',
                    hovertemplate='<b>%{x|%a %d %b}</b><br>This Event: %{customdata}<br>Cumulative: %{y}<br><extra></extra>',
                    customdata=cumulative_df['attendance']
                ),
                row=1, col=1
            )
            
            # Right chart: Individual event attendance with bars
            fig.add_trace(
                go.Bar(
                    x=cumulative_df['date'],
                    y=cumulative_df['attendance'],
                    name='Event Attendance',
                    marker=dict(
                        color=colors['individual'],
                        line=dict(width=1, color='white')
                    ),
                    hovertemplate='<b>%{x|%a %d %b}</b><br>Attendees: %{y}<br><extra></extra>'
                ),
                row=1, col=2
            )
            
            # Calculate cumulative target
            baseline_lunch = 16  # From strategic plan
            annual_target_lunch = baseline_lunch * 1.10  # 10% growth
            cumulative_target, year_progress = calculate_prorated_cumulative_target(annual_target_lunch, baseline_lunch)
            
            # Add target line to cumulative chart
            fig.add_hline(
                y=cumulative_target,
                line=dict(color=colors['target'], width=3, dash='dash'),
                annotation_text=f"🎯 Target: {cumulative_target:.1f}",
                annotation_position="top right",
                annotation=dict(
                    font=dict(size=14, color=colors['target'], family="Inter"),
                    bgcolor="rgba(245, 158, 11, 0.1)",
                    bordercolor=colors['target'],
                    borderwidth=2,
                    borderpad=8
                ),
                row=1, col=1
            )

    # Stunning layout with dark theme
    fig.update_layout(
        title=dict(
            text="<b style='font-size:28px; color:#f1f5f9'>🍽️ Newcomers Lunch Progress Dashboard</b><br><span style='font-size:16px; color:#94a3b8'>Building Community Through Fellowship</span>",
            x=0.5,
            y=0.95,
            font=dict(family="Inter, system-ui, sans-serif")
        ),
        plot_bgcolor=colors['background'],
        paper_bgcolor=colors['background'],
        font=dict(family="Inter, system-ui, sans-serif", size=12, color=colors['text']),
        height=800,
        width=1400,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            xanchor="center",
            x=0.5,
            font=dict(size=14, color=colors['text']),
            bgcolor="rgba(12, 17, 32, 0.8)",
            bordercolor="rgba(148, 163, 184, 0.3)",
            borderwidth=1
        ),
        margin=dict(l=80, r=80, t=140, b=120)
    )

    # Beautiful axes
    fig.update_xaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=12, color=colors['text']),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor=colors['grid'],
        tickfont=dict(size=12, color=colors['text']),
        linecolor="rgba(148, 163, 184, 0.3)",
        mirror=True
    )
    
    # Set y-axis titles
    fig.update_yaxes(title_text="<b>Cumulative Attendees</b>", row=1, col=1)
    fig.update_yaxes(title_text="<b>Event Attendance</b>", row=1, col=2)

    # Save the stunning newcomers dashboard
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    newcomers_filename = f"stunning_newcomers_dashboard_{timestamp}.html"

    try:
        fig.write_html(newcomers_filename)
        print(f"\n✨ Saved STUNNING Newcomers Dashboard: {newcomers_filename}")
    except Exception as e:
        print(f"\n❌ Failed to save Newcomers Dashboard: {e}")

    try:
        newcomers_png = f"stunning_newcomers_dashboard_{timestamp}.png"
        fig.write_image(newcomers_png, width=1400, height=800, scale=3)
        print(f"✨ Saved Newcomers PNG: {newcomers_png}")
    except Exception as e:
        print(f"⚠️ Newcomers PNG save failed: {e}")

    return newcomers_filename, fig

def main():
    """Main execution function"""
    try:
        print("🚀 Starting ENHANCED Bringing Glory to Jesus Christ analysis...")

        # Define current year variables at the start
        current_year = datetime.now().year
        last_year = current_year - 1

        # Step 1: Find both current year and last year service attendance reports
        current_service_group, last_year_service_group = find_service_attendance_reports()
        current_prayer_data = []
        last_year_prayer_data = []

        # Extract current year prayer meeting data
        current_weekly_services = []
        current_quarterly_services = []
        if current_service_group:
            current_service_df, current_service_headers = extract_service_attendance_data(current_service_group, "Current Year Service Attendance")
            if current_service_df is not None and current_service_headers:
                current_weekly_services, current_quarterly_services = parse_prayer_meeting_columns(
                    current_service_headers, current_year, "Current Year"
                )
                # Combine weekly and quarterly prayer meetings
                all_current_prayer_services = current_weekly_services + current_quarterly_services
                if all_current_prayer_services:
                    # Sort by date after combining
                    all_current_prayer_services.sort(key=lambda x: x['date'])
                    current_prayer_data = calculate_prayer_meeting_attendance(current_service_df, all_current_prayer_services)
                    print(f"✅ Current year: {len(current_weekly_services)} weekly + {len(current_quarterly_services)} quarterly = {len(all_current_prayer_services)} total prayer meetings")
                else:
                    print("⚠️ No current year prayer meetings found")
            else:
                print("⚠️ Could not extract current year service attendance data")
        else:
            print("⚠️ No current year service attendance report found")

        # Extract last year prayer meeting data
        last_year_weekly_services = []
        last_year_quarterly_services = []
        if last_year_service_group:
            last_year_service_df, last_year_service_headers = extract_service_attendance_data(last_year_service_group, "Last Year Service Attendance")
            if last_year_service_df is not None and last_year_service_headers:
                last_year_weekly_services, last_year_quarterly_services = parse_prayer_meeting_columns(
                    last_year_service_headers, last_year, "Last Year"
                )
                # Combine weekly and quarterly prayer meetings
                all_last_year_prayer_services = last_year_weekly_services + last_year_quarterly_services
                if all_last_year_prayer_services:
                    # Sort by date after combining
                    all_last_year_prayer_services.sort(key=lambda x: x['date'])
                    last_year_prayer_data = calculate_prayer_meeting_attendance(last_year_service_df, all_last_year_prayer_services)
                    print(f"✅ Last year: {len(last_year_weekly_services)} weekly + {len(last_year_quarterly_services)} quarterly = {len(all_last_year_prayer_services)} total prayer meetings")
                else:
                    print("⚠️ No last year prayer meetings found")
            else:
                print("⚠️ Could not extract last year service attendance data")
        else:
            print("⚠️ No last year service attendance report found")

        # Step 2: Find newcomers lunch report
        lunch_group = find_report_group(['newcomers lunch', 'lunch', 'report of newcomers'])
        lunch_data = []
        if lunch_group:
            lunch_data = extract_newcomers_lunch_data_from_report(lunch_group)
        else:
            print("⚠️ No newcomers lunch report found")

        # Step 3: Create TWO SEPARATE STUNNING dashboards
        print("\n✨ Creating TWO SEPARATE STUNNING dashboards...")
        
        # Create Prayer Dashboard
        prayer_file, prayer_fig = create_stunning_prayer_dashboard(
            current_prayer_data, last_year_prayer_data, 
            current_weekly_services, current_quarterly_services, 
            last_year_weekly_services, last_year_quarterly_services, 
            current_year, last_year
        )
        
        # Create Newcomers Dashboard  
        newcomers_file, newcomers_fig = create_stunning_newcomers_dashboard(lunch_data, current_year)

        # Step 4: Summary
        print(f"\n🎉 ENHANCED BRINGING GLORY TO JESUS CHRIST DASHBOARDS COMPLETE!")
        print(f"✨ Prayer Dashboard (Zoom + Quarterly): {prayer_file}")
        print(f"✨ Newcomers Dashboard: {newcomers_file}")
        print(f"🔗 Zoom prayer meetings: {len(current_weekly_services)} current + {len(last_year_weekly_services)} last year")
        print(f"📅 Quarterly meetings: {len(current_quarterly_services)} current + {len(last_year_quarterly_services)} last year")
        print(f"🍽️ Newcomers lunches: {len(lunch_data)} data points")

        # Step 5: Strategic summary with separate zoom and quarterly analysis
        if current_prayer_data and last_year_prayer_data:
            print(f"🔗 Zoom Prayer Analysis: {len(current_weekly_services)} current weekly + {len(last_year_weekly_services)} last year weekly meetings")
            print(f"📅 Quarterly Prayer Analysis: {len(current_quarterly_services)} current + {len(last_year_quarterly_services)} last year quarterly meetings")
        elif current_prayer_data:
            print(f"🔗 Zoom Prayer Analysis: {len(current_weekly_services)} {current_year} weekly meetings (no {last_year} data for comparison)")
            print(f"📅 Quarterly Prayer Analysis: {len(current_quarterly_services)} {current_year} quarterly meetings")
        elif last_year_prayer_data:
            print(f"🔗 Zoom Prayer Analysis: {len(last_year_weekly_services)} {last_year} weekly meetings (no {current_year} data)")
            print(f"📅 Quarterly Prayer Analysis: {len(last_year_quarterly_services)} {last_year} quarterly meetings")
        else:
            print("🙏 No prayer meeting data found")

        if lunch_data:
            current_year_lunch = [r for r in lunch_data if r['date'].year == current_year]
            if current_year_lunch:
                cumulative_total = sum(r['attendance'] for r in current_year_lunch)
                baseline = 16
                annual_target = baseline * 1.10
                cumulative_target, year_progress = calculate_prorated_cumulative_target(annual_target, baseline)
                lunch_progress = (cumulative_total / cumulative_target * 100) if cumulative_target > 0 else 0
                print(f"🍽️ Lunch: {lunch_progress:.0f}% of prorated cumulative target ({cumulative_total} vs {cumulative_target:.1f})")

        print(f"\n✨ VISUAL ENHANCEMENTS APPLIED:")
        print(f"   • SEPARATED prayer dashboard into Zoom vs Quarterly charts")
        print(f"   • TOP: Zoom prayer meetings with 6-week moving averages")
        print(f"   • BOTTOM: Quarterly meetings with side-by-side bar comparison")
        print(f"   • Filtered zoom data to exclude quarterly meeting weeks")
        print(f"   • Dark themes with modern color palettes")
        print(f"   • Enhanced typography and dual-chart layout")
        print(f"   • Professional annotations and hover effects")
        print(f"   • High-resolution output for presentations")

        # Open both dashboards
        try:
            prayer_path = os.path.abspath(prayer_file)
            newcomers_path = os.path.abspath(newcomers_file)
            webbrowser.open(f"file://{prayer_path}")
            # Small delay before opening second dashboard
            import time
            time.sleep(1)
            webbrowser.open(f"file://{newcomers_path}")
            print(f"\n🌐 Both stunning dashboards opened automatically!")
        except Exception as e:
            print(f"\n⚠️ Couldn't auto-open dashboards: {e}")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
