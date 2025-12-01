#!/usr/bin/env python3
"""
St George's Magill - Congregation Attendance Chart
Displays average attendance by congregation from 2014 onwards
- 2014-2024: Hardcoded historical data
- 2025+: Live data from Elvanto
"""

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
from datetime import datetime
import pandas as pd
import plotly.graph_objects as go
from bs4 import BeautifulSoup
import re
import webbrowser
import os

print("🏛️ ST GEORGE'S MAGILL - CONGREGATION ATTENDANCE CHART")
print("="*60)

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

# Hardcoded historical data (2014-2024)
# These are calculated with 9:30 AM combined services pro-rated
HISTORICAL_DATA = {
    2014: {'8:30 AM': 61.744124, '10:15/10:30 AM': 77.093086, 'Evening': None},
    2015: {'8:30 AM': 60.066957, '10:15/10:30 AM': 85.581192, 'Evening': None},
    2016: {'8:30 AM': 60.064083, '10:15/10:30 AM': 81.747238, 'Evening': None},
    2017: {'8:30 AM': 56.791974, '10:15/10:30 AM': 73.135299, 'Evening': None},
    2018: {'8:30 AM': 46.470666, '10:15/10:30 AM': 77.307111, 'Evening': None},
    2019: {'8:30 AM': 42.050988, '10:15/10:30 AM': 74.245308, 'Evening': None},
    2020: {'8:30 AM': 37.037184, '10:15/10:30 AM': 69.578200, 'Evening': None},
    2021: {'8:30 AM': 35.791424, '10:15/10:30 AM': 71.504872, 'Evening': 30.00},
    2022: {'8:30 AM': 36.353323, '10:15/10:30 AM': 72.627446, 'Evening': 30.00},
    2023: {'8:30 AM': 31.303324, '10:15/10:30 AM': 64.889524, 'Evening': 25.44},
    2024: {'8:30 AM': 26.952966, '10:15/10:30 AM': 67.948995, 'Evening': 23.54},
}

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

def find_attendance_report_groups():
    """Find current year, last year, and two years ago attendance report groups"""
    print("\n📋 Searching for attendance report groups...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return []
    
    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []
    
    # Find all service attendance report groups
    report_groups = []
    
    for group in groups:
        group_name = group.get('name', '').lower()
        
        # Match current year, last year, two years ago patterns
        if 'service individual attendance' in group_name:
            if 'two years ago' in group_name:
                report_groups.append(('two_years_ago', group))
                print(f"✅ Found two years ago group: {group.get('name')}")
            elif 'last year' in group_name:
                report_groups.append(('last_year', group))
                print(f"✅ Found last year group: {group.get('name')}")
            elif 'report of service individual attendance' in group_name:
                report_groups.append(('current_year', group))
                print(f"✅ Found current year group: {group.get('name')}")
    
    return report_groups

def parse_column_header(header):
    """Parse column headers like '9:30 AMMorning Prayer 14/01' or 'Communion 2nd Order 02/06/2024 8:30 AM'"""
    
    # Extract time (look for patterns like 8:30 AM, 10:30 AM, 6:30 PM, etc.)
    time_pattern = r'(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm))'
    time_match = re.search(time_pattern, header)
    time = time_match.group(1).strip() if time_match else None
    
    # Normalize time format
    if time:
        time = re.sub(r'\s+', ' ', time).upper()
    
    # Extract date (look for DD/MM or DD/MM/YYYY patterns)
    date_pattern = r'(\d{1,2}/\d{1,2}(?:/\d{4})?)'
    date_match = re.search(date_pattern, header)
    date_str = date_match.group(1) if date_match else None
    
    return time, date_str

def download_group_attendance_data(group):
    """Download attendance data from a group"""
    if not group:
        return None
    
    group_id = group.get('id')
    group_name = group.get('name', 'Unknown')
    
    print(f"\n📊 Downloading data from: {group_name}")
    
    # Get the report (it's stored as a group with custom fields)
    response = make_request('groups/getInfo', {
        'id': group_id,
        'fields': ['people']
    })
    
    if not response or 'group' not in response:
        print(f"   ❌ Failed to get group data")
        return None
    
    group_data = response['group']
    people = group_data.get('people', {}).get('person', [])
    
    if not isinstance(people, list):
        people = [people] if people else []
    
    print(f"   Found {len(people)} people records")
    
    # Parse attendance from custom fields or HTML table
    attendance_data = []
    
    for person in people:
        person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
        
        # The attendance is usually in custom fields or a table
        # We need to look at the structure to find service dates and attendance
        custom_fields = person.get('custom_fields', {})
        
        # This is where we'd parse the attendance records
        # The exact structure depends on how Elvanto stores the report
        # For now, we'll try to extract from any HTML content
        
    return attendance_data

def get_elvanto_year_data(year_offset):
    """Get attendance data for a specific year from Elvanto
    year_offset: 0 = current year, 1 = last year, 2 = two years ago
    Simple approach: ignore 9:30, use only final January Sunday + rest of year
    """
    current_year = datetime.now().year
    target_year = current_year - year_offset
    
    print(f"\n📅 Getting {target_year} data from Elvanto...")
    
    # Find the appropriate report group
    report_groups = find_attendance_report_groups()
    
    target_group = None
    if year_offset == 0:
        target_group = next((g for t, g in report_groups if t == 'current_year'), None)
    elif year_offset == 1:
        target_group = next((g for t, g in report_groups if t == 'last_year'), None)
    elif year_offset == 2:
        target_group = next((g for t, g in report_groups if t == 'two_years_ago'), None)
    
    if not target_group:
        print(f"   ⚠️ No report group found for {target_year}")
        return None
    
    # Extract URL from group location fields
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if target_group.get(field) and 'http' in str(target_group[field]):
            report_url = str(target_group[field])
            break
    
    if not report_url:
        print(f"   ❌ No URL found in group")
        return None
    
    print(f"   ✅ Found report URL")
    
    # Fetch and parse HTML
    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Find attendance table
        tables = soup.find_all('table')
        table = None
        headers = []
        
        for t in tables:
            header_row = t.find('tr')
            if header_row:
                test_headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
                header_text = ' '.join(test_headers).lower()
                
                if 'first name' in header_text and any('/' in h for h in test_headers):
                    table = t
                    headers = test_headers
                    print(f"   ✅ Found attendance table with {len(headers)} columns")
                    break
        
        if not table:
            print(f"   ❌ No attendance table found")
            return None
        
        # Extract all attendance records into a DataFrame
        attendance_records = []
        for row in table.find_all('tr')[1:]:  # Skip header
            cells = row.find_all(['td', 'th'])
            if len(cells) >= len(headers):
                row_data = {}
                for i, header in enumerate(headers):
                    if i < len(cells):
                        row_data[header] = cells[i].get_text(strip=True)
                if any(row_data.values()):  # Skip empty rows
                    attendance_records.append(row_data)
        
        print(f"   ✅ Extracted {len(attendance_records)} attendance records")
        
        if not attendance_records:
            return None
        
        # Create DataFrame
        df = pd.DataFrame(attendance_records)
        
        # Parse service columns - IGNORE 9:30 ENTIRELY
        print(f"\n   📅 Parsing service columns (ignoring 9:30 AM and future services)...")
        
        services_by_time = {
            '8:30 AM': [],
            '10:30 AM': [],
            '6:30 PM': []
        }
        
        january_services_by_time = {
            '8:30 AM': [],
            '10:30 AM': [],
            '6:30 PM': []
        }
        
        future_services_count = 0
        ignored_930_count = 0
        today = datetime.now().date()
        
        for header in headers:
            # Skip name columns
            if any(skip in header.lower() for skip in ['first name', 'last name', 'category', 'email', 'phone']):
                continue
            
            # Parse time and date from header
            time_str, date_str = parse_column_header(header)
            
            if not time_str or not date_str:
                continue
            
            # Parse full date from date string
            service_date = None
            month = None
            try:
                # Try DD/MM/YYYY format
                date_match = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
                if date_match:
                    day = int(date_match.group(1))
                    month = int(date_match.group(2))
                    year = int(date_match.group(3))
                    service_date = datetime(year, month, day).date()
                else:
                    # Try DD/MM format (assume target year)
                    date_match = re.match(r'(\d{1,2})/(\d{1,2})', date_str)
                    if date_match:
                        day = int(date_match.group(1))
                        month = int(date_match.group(2))
                        service_date = datetime(target_year, month, day).date()
            except ValueError:
                # Invalid date
                continue
            
            # Skip future services (haven't happened yet)
            if service_date and service_date > today:
                future_services_count += 1
                continue
            
            if header not in df.columns:
                continue
            
            # Normalize time and SKIP 9:30
            time_normalized = time_str.upper().replace(' ', '')
            if '9:30' in time_normalized or '930' in time_normalized:
                # IGNORE 9:30 AM services completely
                ignored_930_count += 1
                continue
            elif '8:30' in time_normalized or '830' in time_normalized:
                service_time = '8:30 AM'
            elif '10:30' in time_normalized or '1030' in time_normalized or '10:15' in time_normalized or '1015' in time_normalized:
                service_time = '10:30 AM'
            elif '6:30' in time_normalized or '630' in time_normalized:
                service_time = '6:30 PM'
            else:
                continue
            
            # Count attendees
            attendance_count = len(df[df[header] == 'Y'])
            
            # Categorize by month
            if month == 1:
                january_services_by_time[service_time].append(attendance_count)
            else:
                services_by_time[service_time].append(attendance_count)
        
        # Apply "final January only" filter: keep only last January service
        print(f"\n   🔍 Filtering results:")
        print(f"      Excluded {future_services_count} future services (not yet happened)")
        print(f"      Excluded {ignored_930_count} 9:30 AM combined services")
        print(f"   🔍 Applying 'final January only' filter...")
        
        for service_time in ['8:30 AM', '10:30 AM', '6:30 PM']:
            jan_services = january_services_by_time[service_time]
            if jan_services:
                # Keep only the last (final) January service
                final_january = jan_services[-1]
                services_by_time[service_time].append(final_january)
                print(f"      {service_time} January: Using final service only ({final_january}), excluded {len(jan_services)-1}")
            else:
                print(f"      {service_time} January: No services found")
        
        # Extract final lists
        am_830_attendances = services_by_time['8:30 AM']
        am_1030_attendances = services_by_time['10:30 AM']
        pm_630_attendances = services_by_time['6:30 PM']
        
        print(f"\n   📊 Service counts:")
        print(f"      8:30 AM: {len(am_830_attendances)} services")
        print(f"      10:30 AM: {len(am_1030_attendances)} services")
        print(f"      6:30 PM: {len(pm_630_attendances)} services")
        
        # Calculate final averages (NO PRO-RATING)
        final_avg_830 = sum(am_830_attendances) / len(am_830_attendances) if am_830_attendances else None
        final_avg_1030 = sum(am_1030_attendances) / len(am_1030_attendances) if am_1030_attendances else None
        final_avg_630 = sum(pm_630_attendances) / len(pm_630_attendances) if pm_630_attendances else None
        
        print(f"\n   ✅ Final averages (9:30 AM services ignored):")
        print(f"      8:30 AM: {final_avg_830:.1f}" if final_avg_830 else "      8:30 AM: None")
        print(f"      10:30 AM: {final_avg_1030:.1f}" if final_avg_1030 else "      10:30 AM: None")
        print(f"      6:30 PM: {final_avg_630:.1f}" if final_avg_630 else "      6:30 PM: None")
        
        return {
            '8:30 AM': final_avg_830,
            '10:15/10:30 AM': final_avg_1030,
            'Evening': final_avg_630
        }
        
    except Exception as e:
        print(f"   ❌ Error fetching/parsing data: {e}")
        import traceback
        traceback.print_exc()
        return None

def create_chart(data_by_year):
    """Create the congregation attendance chart"""
    print("\n📊 Creating chart...")
    
    # Sort years
    years = sorted(data_by_year.keys())
    
    # Prepare data for each congregation
    congregation_830 = []
    congregation_1030 = []
    congregation_evening = []
    
    for year in years:
        congregation_830.append(data_by_year[year].get('8:30 AM'))
        congregation_1030.append(data_by_year[year].get('10:15/10:30 AM'))
        congregation_evening.append(data_by_year[year].get('Evening'))
    
    # Calculate percentage changes for footnote
    first_830 = next((x for x in congregation_830 if x is not None), None)
    last_830 = next((x for x in reversed(congregation_830) if x is not None), None)
    pct_change_830 = ((last_830 - first_830) / first_830 * 100) if (first_830 and last_830) else 0
    
    first_1030 = next((x for x in congregation_1030 if x is not None), None)
    last_1030 = next((x for x in reversed(congregation_1030) if x is not None), None)
    pct_change_1030 = ((last_1030 - first_1030) / first_1030 * 100) if (first_1030 and last_1030) else 0
    
    print(f"\n📈 Percentage changes:")
    print(f"   8:30 AM: {pct_change_830:.1f}% (from {first_830:.1f} to {last_830:.1f})")
    print(f"   10:30 AM: {pct_change_1030:.1f}% (from {first_1030:.1f} to {last_1030:.1f})")
    
    # Create figure
    fig = go.Figure()
    
    # Add bars for each congregation (no text labels)
    # Colors matched to the model screenshot
    fig.add_trace(go.Bar(
        name='8:30 AM Congregation',
        x=years,
        y=congregation_830,
        marker_color='rgb(105, 108, 200)',  # Blue/purple color from model
    ))
    
    fig.add_trace(go.Bar(
        name='10:15/10:30 AM Congregation',
        x=years,
        y=congregation_1030,
        marker_color='rgb(118, 192, 128)',  # Green color from model
    ))
    
    fig.add_trace(go.Bar(
        name='Evening Congregation',
        x=years,
        y=congregation_evening,
        marker_color='rgb(255, 128, 96)',  # Orange color from model
    ))
    
    # Update layout
    max_year = max(years)
    num_years = max_year - 2014 + 1
    
    fig.update_layout(
        # title={
        #     'text': f'Average Attendance by Congregation (2014-{max_year})',
        #     'x': 0.5,
        #     'xanchor': 'center',
        #     'font': {'size': 18, 'color': 'black', 'family': 'Helvetica, Arial', 'weight': 'bold'}
        # },
        xaxis={
            'title': '',
            'tickmode': 'linear',
            'tick0': 2014,
            'dtick': 1,
            'showgrid': False,
            'showline': True,
            'linecolor': 'lightgray',
        },
        yaxis={
            'title': '',
            'range': [0, 105],
            'showgrid': True,
            'gridcolor': 'lightgray',
            'showline': True,
            'linecolor': 'lightgray',
        },
        barmode='group',
        bargap=0.15,  # Smaller gap between years (default is 0.2)
        bargroupgap=0.05,  # Smaller gap between bars in a group (default is 0.1)
        plot_bgcolor='white',
        paper_bgcolor='white',
        font={'family': 'Arial', 'size': 12, 'color': 'black'},
        legend={
            'orientation': 'h',
            'yanchor': 'top',
            'y': -0.08,  # Moved closer to chart
            'xanchor': 'center',
            'x': 0.5,
            'font': {'size': 15}  # Increased from 11 to 15 (only slightly smaller than title at 18)
        },
        margin={'l': 60, 'r': 40, 't': 80, 'b': 180},  # Reduced bottom margin
        height=850,
        width=1000,
    )
    
    # Add annotation with note (dynamically calculated percentages, much larger font)
    note_text = (f"Note: The occasional 9:30 AM combined congregation has been pro-rated between the 8:30 AM and 10:30 AM congregations based on<br>"
                 f"the ratio of regular attendances. The 10:15 AM congregation became the 10:30 AM congregation over time. The 6:30 PM congregation<br>"
                 f"was added in 2021 and is included in \"Evening Congregation.\" Over the {num_years}-year period, the 8:30 AM congregation decreased by<br>"
                 f"{abs(pct_change_830):.0f}%, while the 10:15/10:30 AM congregation decreased by {abs(pct_change_1030):.0f}%.")
    
    fig.add_annotation(
        text=note_text,
        xref='paper',
        yref='paper',
        x=0.5,
        y=-0.15,  # Moved closer to legend
        xanchor='center',
        yanchor='top',
        showarrow=False,
        font={'size': 15, 'color': 'gray'},  # Increased from 12 to 15 (only slightly smaller than title at 18)
        align='center',
    )
    
    return fig

def main():
    """Main function"""
    # Start with historical data
    data_by_year = HISTORICAL_DATA.copy()
    
    current_year = datetime.now().year
    
    # Get data for years after 2024 from Elvanto
    for year_offset in range(3):  # Check current year and previous 2 years
        target_year = current_year - year_offset
        
        if target_year > 2024:  # Only get Elvanto data for years after 2024
            print(f"\n{'='*60}")
            print(f"Getting {target_year} data from Elvanto...")
            print(f"{'='*60}")
            
            try:
                elvanto_data = get_elvanto_year_data(year_offset)
                
                if elvanto_data and any(v is not None for v in elvanto_data.values()):
                    data_by_year[target_year] = elvanto_data
                    print(f"\n✅ {target_year} data added to chart:")
                    print(f"   8:30 AM: {elvanto_data.get('8:30 AM')}")
                    print(f"   10:30 AM: {elvanto_data.get('10:15/10:30 AM')}")
                    print(f"   Evening: {elvanto_data.get('Evening')}")
                else:
                    print(f"\n⚠️  {target_year} data not available - chart will only show 2014-2024")
            except Exception as e:
                print(f"\n❌ Error getting {target_year} data: {e}")
                import traceback
                traceback.print_exc()
    
    print(f"\n📊 Chart will display years: {sorted(data_by_year.keys())}")
    
    # Create chart
    fig = create_chart(data_by_year)
    
    # Create outputs directory if it doesn't exist
    os.makedirs('outputs', exist_ok=True)
    
    # Save HTML
    html_path = 'outputs/congregation_attendance.html'
    fig.write_html(html_path)
    print(f"\n✅ HTML saved: {html_path}")
    
    # Save PNG
    png_path = 'outputs/congregation_attendance.png'
    fig.write_image(png_path, width=1000, height=850)
    print(f"✅ PNG saved: {png_path}")
    
    # Open HTML in browser
    abs_html_path = os.path.abspath(html_path)
    webbrowser.open('file://' + abs_html_path)
    print(f"\n🌐 Opening chart in browser...")
    
    print("\n" + "="*60)
    print("✅ Chart generation complete!")

if __name__ == "__main__":
    main()
