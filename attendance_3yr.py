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

# Now import everything
import requests
import json
from datetime import datetime, timedelta
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from bs4 import BeautifulSoup
import re
import numpy as np

print("🏛️ ST GEORGE'S MAGILL - ENHANCED ATTENDANCE ANALYSIS")
print("="*55)

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

def find_attendance_report_groups():
    """Find current year, last year, and two years ago attendance report groups"""
    print("\n📋 Searching for attendance report groups...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None, None
    
    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []
    
    current_year_group = None
    last_year_group = None
    two_years_ago_group = None
    
    for group in groups:
        group_name = group.get('name', '').lower()
        if 'report of two years ago service individual attendance' in group_name:
            two_years_ago_group = group
            print(f"✅ Found two years ago group: {group.get('name')}")
        elif 'report of last year service individual attendance' in group_name:
            last_year_group = group
            print(f"✅ Found last year group: {group.get('name')}")
        elif 'report of service individual attendance' in group_name and 'last year' not in group_name and 'two years ago' not in group_name:
            current_year_group = group
            print(f"✅ Found current year group: {group.get('name')}")
    
    if not current_year_group:
        print("❌ Current year attendance report group not found")
    if not last_year_group:
        print("❌ Last year attendance report group not found")
    if not two_years_ago_group:
        print("❌ Two years ago attendance report group not found")
    
    return current_year_group, last_year_group, two_years_ago_group

def parse_column_header(header):
    """Parse column headers like '9:30 AMMorning Prayer 14/01' or 'Communion 2nd Order 02/06/2024 8:30 AM'"""
    
    # Extract time (look for patterns like 8:30 AM, 10:30 AM, etc.)
    time_match = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', header, re.IGNORECASE)
    if not time_match:
        return None
    
    time_str = f"{time_match.group(1)} {time_match.group(2).upper()}"
    
    # Extract date - handle both DD/MM and DD/MM/YYYY formats
    # Try DD/MM/YYYY format first (current year data)
    date_match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', header)
    if date_match:
        day = int(date_match.group(1))
        month = int(date_match.group(2))
        year = int(date_match.group(3))
        date_format = 'full'
        
        # Extract service name (everything before the date)
        service_name = header
        # Remove the full date and time parts
        service_name = re.sub(r'\d{1,2}/\d{1,2}/\d{4}\s*\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
        service_name = service_name.strip()
        
        return {
            'time': time_str,
            'day': day,
            'month': month,
            'year': year,
            'service_name': service_name,
            'original_header': header,
            'date_format': date_format
        }
    
    # Try DD/MM format (last year data)
    date_match = re.search(r'(\d{1,2})/(\d{1,2})(?:\s|$)', header)
    if date_match:
        day = int(date_match.group(1))
        month = int(date_match.group(2))
        date_format = 'short'
        
        # Extract service name (everything between time and date)
        service_name = header
        # Remove time part
        service_name = re.sub(r'\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
        # Remove date part
        service_name = re.sub(r'\d{1,2}/\d{1,2}(?:\s|$)', '', service_name)
        service_name = service_name.strip()
        
        return {
            'time': time_str,
            'day': day,
            'month': month,
            'year': None,  # Year will be assigned later
            'service_name': service_name,
            'original_header': header,
            'date_format': date_format
        }
    
    return None

def should_exclude_service(service_name, day_of_week):
    """Check if service should be excluded based on name or day"""
    
    # Exclude non-Sunday services
    if day_of_week != 6:  # Sunday = 6
        return True
    
    service_name_lower = service_name.lower()
    
    # Exclude specific service types
    exclude_patterns = [
        'wednesday', 'saturday', 'prayer meeting',
        'good friday', 'christmas', 'christmas day',
        'easter vigil', 'maundy thursday'
    ]
    
    for pattern in exclude_patterns:
        if pattern in service_name_lower:
            return True
    
    return False

def extract_attendance_data_from_group(group, year_label):
    """Extract attendance data from a specific group"""
    print(f"\n🔍 Extracting {year_label} attendance data...")
    
    # Determine target year based on group name
    if 'two years ago' in year_label.lower():
        target_year = datetime.now().year - 2
    elif 'last year' in year_label.lower():
        target_year = datetime.now().year - 1
    else:
        target_year = datetime.now().year
    
    print(f"   Target year: {target_year}")
    
    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break
    
    if not report_url:
        print(f"❌ No URL found for {year_label} group")
        return None, None
    
    print(f"✅ Found {year_label} report URL")
    
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
                    print(f"✅ Found {year_label} attendance table")
                    
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
                    
                    print(f"✅ Extracted {len(attendance_records)} {year_label} records")
                    return pd.DataFrame(attendance_records), headers, target_year
        
        print(f"❌ {year_label} attendance table not found")
        return None, None, None
        
    except Exception as e:
        print(f"❌ Error extracting {year_label} data: {e}")
        return None, None, None

def parse_service_columns_for_year_flexible(headers, target_year, year_label):
    """Parse service columns more flexibly - allow recent data even if not from target year"""
    print(f"\n📅 Analyzing {year_label} service columns (flexible mode)...")
    
    service_columns = []
    excluded_count = 0
    unparseable_count = 0
    invalid_date_count = 0
    
    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue
            
        # Parse the column header
        parsed = parse_column_header(header)
        if not parsed:
            unparseable_count += 1
            continue
        
        # Determine which year to use
        if parsed.get('date_format') == 'full':
            # Use year from the header itself
            service_year = parsed['year']
        else:
            # For flexible mode, accept recent data (2024 or later)
            service_year = max(target_year - 1, parsed.get('year', target_year))
        
        try:
            # Create date using the determined year
            service_date = datetime(service_year, parsed['month'], parsed['day'])
            day_of_week = service_date.weekday()  # Monday=0, Sunday=6
            
            # Check if it's actually a Sunday and not excluded
            if should_exclude_service(parsed['service_name'], day_of_week):
                excluded_count += 1
                continue
            
            # Normalize service time for consistency
            time_str = parsed['time']
            if '9:30' in time_str:
                normalized_time = '9:30 AM'
            elif '8:30' in time_str:
                normalized_time = '8:30 AM'
            elif '10:30' in time_str:
                normalized_time = '10:30 AM'
            elif '6:30' in time_str:
                normalized_time = '6:30 PM'
            elif '6:00' in time_str:
                normalized_time = '6:30 PM'  # Treat 6:00 PM as 6:30 PM
            else:
                normalized_time = 'Other'
            
            service_columns.append({
                'header': header,
                'date': service_date,
                'time': normalized_time,
                'service_name': parsed['service_name'],
                'year': service_year
            })
            
        except ValueError as e:
            # Invalid date (e.g., Feb 30)
            invalid_date_count += 1
            continue
    
    # Sort by date
    service_columns.sort(key=lambda x: x['date'])
    
    print(f"\n✅ Found {len(service_columns)} valid Sunday services (flexible)")
    print(f"   Excluded {excluded_count} non-Sunday or special services")
    print(f"   Couldn't parse {unparseable_count} headers")
    print(f"   {invalid_date_count} invalid dates")
    
    return service_columns

def parse_service_columns_for_year(headers, target_year, year_label):
    """Parse service columns and validate they're Sundays in the target year"""
    print(f"\n📅 Analyzing {year_label} service columns for {target_year}...")
    
    service_columns = []
    excluded_count = 0
    unparseable_count = 0
    invalid_date_count = 0
    
    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue
            
        # Parse the column header
        parsed = parse_column_header(header)
        if not parsed:
            unparseable_count += 1
            continue
        
        # Determine which year to use
        if parsed.get('date_format') == 'full':
            # Use year from the header itself (current year data)
            service_year = parsed['year']
        else:
            # Use target year (last year data)
            service_year = target_year
        
        try:
            # Create date using the determined year
            service_date = datetime(service_year, parsed['month'], parsed['day'])
            day_of_week = service_date.weekday()  # Monday=0, Sunday=6
            
            # For current year data, filter to only include data from current year (2025) onwards
            if 'current' in year_label.lower() and service_year < target_year:
                excluded_count += 1
                continue
            
            # Check if it's actually a Sunday and not excluded
            if should_exclude_service(parsed['service_name'], day_of_week):
                excluded_count += 1
                continue
            
            # Normalize service time for consistency
            time_str = parsed['time']
            if '9:30' in time_str:
                normalized_time = '9:30 AM'
            elif '8:30' in time_str:
                normalized_time = '8:30 AM'
            elif '10:30' in time_str:
                normalized_time = '10:30 AM'
            elif '6:30' in time_str:
                normalized_time = '6:30 PM'
            elif '6:00' in time_str:
                normalized_time = '6:30 PM'  # Treat 6:00 PM as 6:30 PM
            else:
                normalized_time = 'Other'
            
            service_columns.append({
                'header': header,
                'date': service_date,
                'time': normalized_time,
                'service_name': parsed['service_name'],
                'year': service_year
            })
            
        except ValueError as e:
            # Invalid date (e.g., Feb 30)
            invalid_date_count += 1
            continue
    
    # Sort by date
    service_columns.sort(key=lambda x: x['date'])
    
    print(f"\n✅ Found {len(service_columns)} valid Sunday services")
    print(f"   Excluded {excluded_count} non-Sunday or special services")
    print(f"   Couldn't parse {unparseable_count} headers")
    print(f"   {invalid_date_count} invalid dates")
    
    return service_columns

def calculate_service_attendance_by_year(df, service_columns, year_label):
    """Calculate attendance by service time for a given year"""
    print(f"\n📊 Calculating {year_label} attendance by service...")
    
    # Group services by date
    sundays_data = {}
    
    for svc in service_columns:
        date = svc['date']
        time = svc['time']
        header = svc['header']
        
        if date not in sundays_data:
            sundays_data[date] = {
                'overall': set(),
                '8:30 AM': set(),
                '9:30 AM': set(),
                '10:30 AM': set(),
                '6:30 PM': set(),
                'services': []
            }
        
        # Track which services happened on this date
        sundays_data[date]['services'].append({'time': time, 'header': header})
        
        # Count attendance for this specific service
        if header in df.columns:
            service_attendees = set(df[df[header] == 'Y'].index)
            sundays_data[date][time].update(service_attendees)
            sundays_data[date]['overall'].update(service_attendees)
    
    # Convert to list format and handle combined services
    attendance_data = []
    
    for date, data in sundays_data.items():
        sunday_record = {
            'date': date,
            'overall': len(data['overall']),
            '8:30 AM': len(data['8:30 AM']),
            '9:30 AM': len(data['9:30 AM']),
            '10:30 AM': len(data['10:30 AM']),
            '6:30 PM': len(data['6:30 PM']),
            'services_held': [svc['time'] for svc in data['services']]
        }
        
        attendance_data.append(sunday_record)
    
    return sorted(attendance_data, key=lambda x: x['date'])

def apply_pro_rata_logic_to_dataset(attendance_data, dataset_name):
    """Apply pro-rata logic to split 9:30 AM combined services within the same dataset"""
    print(f"\n🔄 Applying pro-rata logic for {dataset_name}...")
    
    # Calculate 8:30 to 10:30 ratio from services where both exist
    ratios_830_to_1030 = []
    
    for record in attendance_data:
        if record['8:30 AM'] > 0 and record['10:30 AM'] > 0:
            ratio = record['8:30 AM'] / record['10:30 AM']
            ratios_830_to_1030.append(ratio)
    
    if ratios_830_to_1030:
        avg_ratio = np.mean(ratios_830_to_1030)
        print(f"✅ Calculated 8:30 AM to 10:30 AM ratio from {len(ratios_830_to_1030)} services: {avg_ratio:.2f}")
    else:
        avg_ratio = 0.4  # Default assumption
        print(f"⚠️ No ratio data found in {dataset_name}, using default: {avg_ratio:.2f}")
    
    # Apply pro-rata logic to 9:30 AM services
    combined_services_count = 0
    
    for record in attendance_data:
        if record['9:30 AM'] > 0:
            total_930 = record['9:30 AM']
            combined_services_count += 1
            
            # Split based on ratio
            proportion_830 = avg_ratio / (1 + avg_ratio)
            proportion_1030 = 1 / (1 + avg_ratio)
            
            estimated_830 = int(total_930 * proportion_830)
            estimated_1030 = int(total_930 * proportion_1030)
            
            # Add to existing counts (in case there were separate services too)
            record['8:30 AM'] += estimated_830
            record['10:30 AM'] += estimated_1030
            
            print(f"  {record['date'].strftime('%d %b %Y')}: Split {total_930} attendees -> {estimated_830} (8:30) + {estimated_1030} (10:30)")
    
    if combined_services_count > 0:
        print(f"✅ Applied pro-rata logic to {combined_services_count} combined services")
    else:
        print("ℹ️ No combined 9:30 AM services found")
    
    return attendance_data

def calculate_rolling_average(data, window=4):
    """Calculate rolling average with specified window size"""
    if len(data) < window:
        return data  # Not enough data for rolling average
    
    rolling_data = []
    for i in range(len(data)):
        if i < window - 1:
            # For early points, use expanding average
            avg = np.mean(data[:i+1])
        else:
            # Use rolling window
            avg = np.mean(data[i-window+1:i+1])
        rolling_data.append(avg)
    
    return rolling_data

def filter_for_average_calculation(df, col_name):
    """Filter dataframe for average calculation - keep only final January data point"""
    print(f"📊 Filtering data for average calculation (final January point only)...")
    
    # Get all non-zero values for this column
    df_nonzero = df[df[col_name] > 0].copy()
    
    if len(df_nonzero) == 0:
        return df_nonzero
    
    # Separate January from other months
    january_data = df_nonzero[df_nonzero['date'].dt.month == 1]
    non_january_data = df_nonzero[df_nonzero['date'].dt.month != 1]
    
    # If we have January data, keep only the final (latest) January point
    if len(january_data) > 0:
        final_january = january_data.loc[january_data['date'].idxmax()]  # Latest date
        print(f"   January: Using final point only - {final_january['date'].strftime('%d %b %Y')} ({final_january[col_name]:.0f})")
        print(f"   January: Excluded {len(january_data) - 1} earlier January services")
        
        # Combine final January point with all non-January data
        filtered_df = pd.concat([pd.DataFrame([final_january]), non_january_data], ignore_index=True)
    else:
        print(f"   January: No January data found")
        filtered_df = non_january_data
    
    print(f"   Total services for average: {len(filtered_df)} (was {len(df_nonzero)})")
    
    return filtered_df

def create_enhanced_combined_dashboard(current_data, last_data, two_years_data, years_choice):
    """Create a single landscape dashboard with all four service charts - HTML and PNG"""
    print(f"\n📊 Creating combined attendance dashboard with {years_choice} year(s) of data...")
    
    current_df = pd.DataFrame(current_data) if current_data else pd.DataFrame()
    last_df = pd.DataFrame(last_data) if last_data and years_choice >= 2 else pd.DataFrame()
    two_years_df = pd.DataFrame(two_years_data) if two_years_data and years_choice == 3 else pd.DataFrame()
    
    current_year = datetime.now().year
    last_year = current_year - 1
    two_years_ago = current_year - 2
    today = datetime.now().date()
    
    print(f"📅 Filtering data to exclude future dates (today: {today})")
    print(f"📊 Note: Average calculations use only the final data point from January")
    
    # Filter out future dates from current year data
    if len(current_df) > 0:
        print(f"   Current data before filtering: {len(current_df)} records")
        current_df = current_df[current_df['date'].dt.date <= today]
        print(f"   Current data after filtering: {len(current_df)} records")
    
    # Determine actual years in the data
    current_years = set()
    last_years = set()
    two_years_years = set()
    
    if len(current_df) > 0:
        current_years = set(current_df['date'].dt.year)
        print(f"Current data contains years: {sorted(current_years)}")
    
    if len(last_df) > 0 and years_choice >= 2:
        last_years = set(last_df['date'].dt.year)
        print(f"Last year data contains years: {sorted(last_years)}")
        
    if len(two_years_df) > 0 and years_choice == 3:
        two_years_years = set(two_years_df['date'].dt.year)
        print(f"Two years ago data contains years: {sorted(two_years_years)}")
    
    # Normalize dates to the same calendar year for comparison
    base_year = 2024
    
    if len(current_df) > 0:
        current_df = current_df.copy()
        current_df['normalized_date'] = current_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        current_df['actual_year'] = current_df['date'].dt.year
    
    if len(last_df) > 0 and years_choice >= 2:
        last_df = last_df.copy()
        last_df['normalized_date'] = last_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        last_df['actual_year'] = last_df['date'].dt.year
        
    if len(two_years_df) > 0 and years_choice == 3:
        two_years_df = two_years_df.copy()
        two_years_df['normalized_date'] = two_years_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        two_years_df['actual_year'] = two_years_df['date'].dt.year
    
    # Adjust labels based on actual data
    if current_years and max(current_years) >= current_year:
        current_label = f"{current_year}"
    elif current_years:
        current_label = f"{max(current_years)}"
    else:
        current_label = "Current Data"
    
    if years_choice >= 2 and last_years:
        last_label = f"{max(last_years) if last_years else last_year}"
    else:
        last_label = "Previous Year"
        
    if years_choice == 3 and two_years_years:
        two_years_label = f"{max(two_years_years) if two_years_years else two_years_ago}"
    else:
        two_years_label = "Two Years Ago"
    
    # Enhanced color scheme
    colors = {
        'current': '#1e40af',
        'current_smooth': '#3b82f6',
        'last': '#dc2626',
        'last_smooth': '#ef4444',
        'two_years': '#059669',
        'two_years_smooth': '#10b981'
    }
    
    # Chart configurations for 2x2 grid
    charts_config = [
        {'col': '8:30 AM', 'title': '8:30 AM Congregation Attendance', 'row': 1, 'col_pos': 1},
        {'col': '10:30 AM', 'title': '10:30 AM Congregation Attendance', 'row': 1, 'col_pos': 2},
        {'col': '6:30 PM', 'title': '6:30 PM Congregation Attendance', 'row': 2, 'col_pos': 1},
        {'col': 'overall', 'title': 'Combined Congregation Attendance', 'row': 2, 'col_pos': 2}
    ]
    
    # PRE-CALCULATE ALL STATS TO EMBED IN SUBPLOT TITLES
    print(f"\n🔄 Pre-calculating statistics for subplot titles...")
    
    subplot_titles_with_stats = []
    all_stats = []
    
    for chart_index, chart_config in enumerate(charts_config, 1):
        col_name = chart_config['col']
        chart_title = chart_config['title']
        
        print(f"\n📈 Pre-calculating stats for {chart_title}...")
        
        # Calculate statistics for this specific chart
        chart_stats = []
        last_avg = None
        current_avg = None
        two_years_avg = None
        
        # Calculate two years ago data if available and requested
        if years_choice == 3 and len(two_years_df) > 0 and col_name in two_years_df.columns:
            two_years_df_filtered = filter_for_average_calculation(two_years_df, col_name)
            if len(two_years_df_filtered) > 0:
                two_years_avg = two_years_df_filtered[col_name].mean()
        
        # Calculate last year data if available and requested
        if years_choice >= 2 and len(last_df) > 0 and col_name in last_df.columns:
            last_df_filtered = filter_for_average_calculation(last_df, col_name)
            if len(last_df_filtered) > 0:
                last_avg = last_df_filtered[col_name].mean()
        
        # Calculate current year data if available
        if len(current_df) > 0 and col_name in current_df.columns:
            current_df_filtered = filter_for_average_calculation(current_df, col_name)
            if len(current_df_filtered) > 0:
                current_avg = current_df_filtered[col_name].mean()
        
        # Build compact stats string based on years_choice
        stats_parts = []
        if years_choice == 3 and two_years_avg is not None:
            stats_parts.append(f"{two_years_label}: {two_years_avg:.1f}")
        if years_choice >= 2 and last_avg is not None:
            stats_parts.append(f"{last_label}: {last_avg:.1f}")
        if current_avg is not None:
            stats_parts.append(f"{current_label}: {current_avg:.1f}")
        
        # Add change calculation (only if we have comparison data)
        if years_choice >= 2 and current_avg is not None and last_avg is not None:
            change = current_avg - last_avg
            change_pct = (change / last_avg * 100) if last_avg > 0 else 0
            direction = "↗" if change > 0 else "↘" if change < 0 else "→"
            stats_parts.append(f"Δ {direction}{change:+.1f} ({change_pct:+.1f}%)")
        
        # Add strategic target for 10:30 AM (only if we have last year data)
        if col_name == '10:30 AM' and years_choice >= 2 and last_avg is not None:
            target_2025 = last_avg * 1.10
            stats_parts.append(f"Target: {target_2025:.1f}")
            if current_avg is not None:
                total_increase_needed = target_2025 - last_avg
                increase_achieved = current_avg - last_avg
                target_progress = (increase_achieved / total_increase_needed) * 100 if total_increase_needed > 0 else 0
                stats_parts.append(f"Progress: {target_progress:.0f}%")
        
        # Create simple title (no stats) for subplot
        subplot_titles_with_stats.append(chart_title)
        
        # Store stats for summary
        all_stats.append({
            'title': chart_title,
            'stats': stats_parts
        })
    
    # Determine dashboard subtitle based on years_choice
    if years_choice == 1:
        dashboard_subtitle = "Current Year Analysis"
    elif years_choice == 2:
        dashboard_subtitle = "Two-Year Comparison Dashboard"
    else:
        dashboard_subtitle = "Three-Year Comparison Dashboard"
    
    # Create subplot figure with 2x2 layout
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=None,
        vertical_spacing=0.25,
        horizontal_spacing=0.08,
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    print(f"\n🔄 Processing {len(charts_config)} charts for combined dashboard...")
    
    for chart_index, chart_config in enumerate(charts_config, 1):
        col_name = chart_config['col']
        chart_title = chart_config['title']
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        
        print(f"\n📈 Processing chart {chart_index}/4: {chart_title}...")
        
        # Add two years ago data if available and requested
        if years_choice == 3 and len(two_years_df) > 0 and col_name in two_years_df.columns:
            print(f"  📊 Processing {two_years_label} data...")
            
            two_years_smooth = calculate_rolling_average(two_years_df[col_name].values, window=4)
            
            hover_text_raw = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {row_data[col_name]}"
                for _, row_data in two_years_df.iterrows()
            ]
            
            hover_text_smooth = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {smooth_val:.1f} (4-week avg)"
                for (_, row_data), smooth_val in zip(two_years_df.iterrows(), two_years_smooth)
            ]
            
            fig.add_trace(
                go.Scatter(
                    x=two_years_df['normalized_date'],
                    y=two_years_df[col_name],
                    mode='lines',
                    name=f'{two_years_label} (Raw)' if chart_index == 1 else None,
                    line=dict(color=colors['two_years'], width=1, dash='dot'),
                    opacity=0.5,
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_raw,
                    showlegend=(chart_index == 1),
                    legendgroup='two_years_raw'
                ),
                row=row, col=col_pos
            )
            
            fig.add_trace(
                go.Scatter(
                    x=two_years_df['normalized_date'],
                    y=two_years_smooth,
                    mode='lines+markers',
                    name=f'{two_years_label}' if chart_index == 1 else None,
                    line=dict(color=colors['two_years_smooth'], width=3),
                    marker=dict(size=4, symbol='square'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_smooth,
                    showlegend=(chart_index == 1),
                    legendgroup='two_years_smooth'
                ),
                row=row, col=col_pos
            )
        
        # Add last year data if available and requested
        if years_choice >= 2 and len(last_df) > 0 and col_name in last_df.columns:
            print(f"  📊 Processing {last_label} data...")
            
            last_smooth = calculate_rolling_average(last_df[col_name].values, window=4)
            
            hover_text_raw = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {row_data[col_name]}"
                for _, row_data in last_df.iterrows()
            ]
            
            hover_text_smooth = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {smooth_val:.1f} (4-week avg)"
                for (_, row_data), smooth_val in zip(last_df.iterrows(), last_smooth)
            ]
            
            fig.add_trace(
                go.Scatter(
                    x=last_df['normalized_date'],
                    y=last_df[col_name],
                    mode='lines',
                    name=f'{last_label} (Raw)' if chart_index == 1 else None,
                    line=dict(color=colors['last'], width=1, dash='dot'),
                    opacity=0.5,
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_raw,
                    showlegend=(chart_index == 1),
                    legendgroup='last_raw'
                ),
                row=row, col=col_pos
            )
            
            fig.add_trace(
                go.Scatter(
                    x=last_df['normalized_date'],
                    y=last_smooth,
                    mode='lines+markers',
                    name=f'{last_label}' if chart_index == 1 else None,
                    line=dict(color=colors['last_smooth'], width=3),
                    marker=dict(size=4, symbol='circle'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_smooth,
                    showlegend=(chart_index == 1),
                    legendgroup='last_smooth'
                ),
                row=row, col=col_pos
            )
        
        # Add current year data if available
        if len(current_df) > 0 and col_name in current_df.columns:
            print(f"  📊 Processing {current_label} data...")
            
            current_smooth = calculate_rolling_average(current_df[col_name].values, window=4)
            
            hover_text_raw = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {row_data[col_name]}"
                for _, row_data in current_df.iterrows()
            ]
            
            hover_text_smooth = [
                f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {smooth_val:.1f} (4-week avg)"
                for (_, row_data), smooth_val in zip(current_df.iterrows(), current_smooth)
            ]
            
            fig.add_trace(
                go.Scatter(
                    x=current_df['normalized_date'],
                    y=current_df[col_name],
                    mode='lines',
                    name=f'{current_label} (Raw)' if chart_index == 1 else None,
                    line=dict(color=colors['current'], width=1, dash='dot'),
                    opacity=0.5,
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_raw,
                    showlegend=(chart_index == 1),
                    legendgroup='current_raw'
                ),
                row=row, col=col_pos
            )
            
            fig.add_trace(
                go.Scatter(
                    x=current_df['normalized_date'],
                    y=current_smooth,
                    mode='lines+markers',
                    name=f'{current_label}' if chart_index == 1 else None,
                    line=dict(color=colors['current_smooth'], width=3),
                    marker=dict(size=4, symbol='diamond'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text_smooth,
                    showlegend=(chart_index == 1),
                    legendgroup='current_smooth'
                ),
                row=row, col=col_pos
            )
        
        # Add invisible traces for statistics in the legend
        stats = all_stats[chart_index - 1]
        if stats['stats']:
            for stat_text in stats['stats']:
                fig.add_trace(
                    go.Scatter(
                        x=[],
                        y=[],
                        mode='markers',
                        name=f"{chart_title[:5]}... | {stat_text}",
                        showlegend=True,
                        visible='legendonly',
                        marker=dict(size=8, symbol='square', color='gray'),
                        legendgroup=f'stats_{chart_index}'
                    ),
                    row=row, col=col_pos
                )
        
        if col_name == '10:30 AM' and years_choice >= 2:
            stats = all_stats[chart_index - 1]
            
            last_avg = None
            current_avg = None
            for stat in stats['stats']:
                if f"{last_label}:" in stat:
                    last_avg = float(stat.split(': ')[1])
                elif f"{current_label}:" in stat:
                    current_avg = float(stat.split(': ')[1])
            
            if last_avg is not None:
                target_2025 = last_avg * 1.10
                
                fig.add_hline(
                    y=target_2025,
                    line=dict(color='#059669', width=2, dash='dash'),
                    annotation_text=f"Target",
                    annotation_position="top right",
                    annotation=dict(
                        font=dict(size=10, color='#059669'),
                        bgcolor="rgba(255,255,255,0.9)",
                        bordercolor='#059669',
                        borderwidth=1
                    ),
                    row=row, col=col_pos
                )
    
    # Set consistent x-axis range for all charts
    x_min = datetime(base_year, 1, 1)
    x_max = datetime(base_year, 12, 31)
    
    print(f"\n🔍 Calculating individual y-axis ranges for optimal visibility...")
    
    # Calculate individual y-axis ranges for each chart
    chart_ranges = {}
    
    for chart_config in charts_config:
        col_name = chart_config['col']
        chart_values = []
        
        if years_choice == 3 and len(two_years_df) > 0 and col_name in two_years_df.columns:
            chart_values.extend(two_years_df[col_name].tolist())
        if years_choice >= 2 and len(last_df) > 0 and col_name in last_df.columns:
            chart_values.extend(last_df[col_name].tolist())
        if len(current_df) > 0 and col_name in current_df.columns:
            chart_values.extend(current_df[col_name].tolist())
        
        if chart_values:
            chart_min = max(0, min(chart_values) - 5)
            chart_max = max(chart_values) + 10
            padding = (chart_max - chart_min) * 0.1
            chart_max += padding
        else:
            chart_min, chart_max = 0, 50
        
        chart_ranges[col_name] = (chart_min, chart_max)
        chart_title = chart_config['title']
        print(f"    📊 {chart_title}: Y-axis range {chart_min:.0f} to {chart_max:.0f}")
    
    # Update layout for each subplot with individual ranges
    for i, chart_config in enumerate(charts_config):
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        col_name = chart_config['col']
        y_min, y_max = chart_ranges[col_name]
        
        fig.update_xaxes(
            range=[x_min, x_max],
            tickformat='%b',
            dtick='M2',
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(0,0,0,0.1)',
            tickfont=dict(size=10),
            row=row, col=col_pos
        )
        fig.update_yaxes(
            range=[y_min, y_max],
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(0,0,0,0.1)',
            tickfont=dict(size=10),
            row=row, col=col_pos
        )
    
    # Build title text
    title_main = "<b style='font-size:24px'>St George's Magill - Attendance Analysis</b>"
    title_sub = f"<br><span style='font-size:14px; color:#64748b'>{dashboard_subtitle}</span>"
    title_note = f"<br><span style='font-size:11px; color:#94a3b8'>Generated {datetime.now().strftime('%B %d, %Y')} • *Averages use only final January data point</span>"
    full_title = title_main + title_sub + title_note
    
    # Build annotation texts
    stats_text_0 = f"<b>8:30 AM Attendance</b><br>{' | '.join(all_stats[0]['stats'])}" if all_stats[0]['stats'] else "<b>8:30 AM Attendance</b>"
    stats_text_1 = f"<b>10:30 AM Attendance</b><br>{' | '.join(all_stats[1]['stats'])}" if all_stats[1]['stats'] else "<b>10:30 AM Attendance</b>"
    stats_text_2 = f"<b>6:30 PM Attendance</b><br>{' | '.join(all_stats[2]['stats'])}" if all_stats[2]['stats'] else "<b>6:30 PM Attendance</b>"
    stats_text_3 = f"<b>Combined Attendance</b><br>{' | '.join(all_stats[3]['stats'])}" if all_stats[3]['stats'] else "<b>Combined Attendance</b>"
    
    # Overall layout
    fig.update_layout(
        title=dict(
            text=full_title,
            x=0.5,
            y=0.985,
            font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=20, color='#1e293b')
        ),
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=12),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=1200,
        width=1600,
        hovermode='closest',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.08,
            xanchor="center",
            x=0.5,
            font=dict(family="Inter, system-ui, sans-serif", size=12, color='#374151'),
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor="rgba(0,0,0,0.1)",
            borderwidth=1
        ),
        margin=dict(l=80, r=80, t=140, b=120),
        annotations=[
            dict(
                text="",
                x=0.01, y=0.01,
                xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=1, color='white'),
                bgcolor="rgba(255,255,255,0)",
                borderwidth=0
            ),
            dict(
                text=stats_text_0,
                x=0.44, y=0.94,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=10, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_1,
                x=0.90, y=0.94,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=10, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_2,
                x=0.44, y=0.45,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=10, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_3,
                x=0.90, y=0.45,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=10, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            )
        ]
    )
    
    # Create outputs directory
    import os
    outputs_dir = "outputs"
    os.makedirs(outputs_dir, exist_ok=True)
    
    # Use fixed filenames (without timestamps) - this will replace old files
    html_filename = os.path.join(outputs_dir, "attendance_dashboard_combined.html")
    png_filename = os.path.join(outputs_dir, "attendance_dashboard_combined.png")
    
    # Delete old files if they exist
    if os.path.exists(html_filename):
        os.remove(html_filename)
        print(f"\n🗑️ Deleted old HTML file")
    if os.path.exists(png_filename):
        os.remove(png_filename)
        print(f"🗑️ Deleted old PNG file")
    
    # Save both HTML and PNG
    try:
        fig.write_html(html_filename)
        print(f"\n✅ Saved HTML: {html_filename}")
    except Exception as e:
        print(f"\n❌ Failed to save HTML: {e}")
        return None
    
    try:
        fig.write_image(png_filename, width=1600, height=1200, scale=3)
        print(f"✅ Saved PNG: {png_filename}")
    except Exception as e:
        print(f"⚠️ PNG save failed: {e}")
    
    print(f"\n📊 Dashboard files ready in /outputs")
    
    # Print summary statistics
    print(f"\n📈 DASHBOARD SUMMARY:")
    for stat in all_stats:
        print(f"  {stat['title']}:")
        for line in stat['stats']:
            print(f"    {line}")
    
    return html_filename

def list_all_files_in_directory():
    """List all files in current directory for debugging"""
    import os
    print("\n📁 ALL FILES IN CURRENT DIRECTORY:")
    files = os.listdir('.')
    for file in sorted(files):
        if file.endswith(('.html', '.png')):
            print(f"   📄 {file}")

def main():
    """Main execution function (always generates 3-year charts, saves both HTML and PNG to outputs)"""
    try:
        print("🚀 Starting enhanced year-over-year attendance analysis...")
        print("📊 Note: Updated to use only the final January data point for averages")
        print("✅ Will generate charts with 3 years of data (auto-selected)")
        print("📁 Output: HTML (interactive) and PNG files saved to /outputs directory")
        print("🔄 Old dashboard files will be replaced with new versions")

        # Always use 3 years
        years_choice = 3

        # Find all three attendance report groups
        current_group, last_year_group, two_years_ago_group = find_attendance_report_groups()

        if not current_group:
            print("❌ Cannot proceed without current year attendance report group")
            return
        
        if not last_year_group:
            print("❌ Last year attendance report group required but not found")
            return
            
        if not two_years_ago_group:
            print("❌ Two years ago attendance report group required but not found")
            return

        # Extract attendance data for 3 years
        current_df, current_headers, current_year = extract_attendance_data_from_group(current_group, "Current Year")
        last_df, last_headers, last_year = extract_attendance_data_from_group(last_year_group, "Last Year")
        two_years_df, two_years_headers, two_years_year = extract_attendance_data_from_group(two_years_ago_group, "Two Years Ago")

        if current_df is None or last_df is None or two_years_df is None:
            print("❌ One or more datasets missing - cannot proceed")
            return

        # Parse service columns
        current_services = parse_service_columns_for_year(current_headers, current_year, "Current Year")
        last_services = parse_service_columns_for_year(last_headers, last_year, "Last Year")
        two_years_services = parse_service_columns_for_year(two_years_headers, two_years_year, "Two Years Ago")

        print(f"\n📊 Data Summary:")
        print(f"   Current year services: {len(current_services)}")
        print(f"   Last year services: {len(last_services)}")
        print(f"   Two years ago services: {len(two_years_services)}")

        if not current_services:
            print("⚠️ No current year data found. Trying flexible parsing...")
            current_services = parse_service_columns_for_year_flexible(current_headers, current_year, "Current Year (Flexible)")
        
        if not current_services:
            print("❌ Cannot proceed without current year service data")
            return

        # Calculate attendance
        current_attendance = apply_pro_rata_logic_to_dataset(
            calculate_service_attendance_by_year(current_df, current_services, "Current Year"), 
            "Current Year"
        )
        last_attendance = apply_pro_rata_logic_to_dataset(
            calculate_service_attendance_by_year(last_df, last_services, "Last Year"), 
            "Last Year"
        )
        two_years_attendance = apply_pro_rata_logic_to_dataset(
            calculate_service_attendance_by_year(two_years_df, two_years_services, "Two Years Ago"), 
            "Two Years Ago"
        )

        # Create combined dashboard (3 years) - saves both HTML and PNG to /outputs
        html_file = create_enhanced_combined_dashboard(
            current_attendance, last_attendance, two_years_attendance, years_choice
        )
        
        if not html_file:
            print("❌ Failed to create dashboard")
            return

        # Open the HTML file automatically (for interactive hover features)
        try:
            import webbrowser
            import os
            abs_path = os.path.abspath(html_file)
            webbrowser.open(f"file://{abs_path}")
            print(f"\n🌐 Interactive dashboard opened automatically!")
        except Exception as e:
            print(f"\n⚠️ Couldn't auto-open dashboard: {e}")

        print(f"\n🎉 UPDATED ATTENDANCE DASHBOARD WITH FINAL JANUARY METHOD!")
        print(f"📊 THREE-YEAR COMPARISON: Current, Last Year, and Two Years Ago")
        print(f"📄 HTML file (interactive): {html_file}")
        print(f"🖼️ PNG file (static): {html_file.replace('.html', '.png')}")
        print(f"\n✨ FEATURES:")
        print(f"   • Uses final January data point for yearly averages")
        print(f"   • Year-over-year comparison with color-coded lines (Blue, Red, Green)")
        print(f"   • Statistics embedded in chart titles")
        print(f"   • Avoids early January irregularities")
        print(f"   • Interactive HTML with hover-over data details")
        print(f"   • Static PNG for reports and presentations")
        print(f"   • Both files saved to /outputs directory")
        print(f"   • Old dashboard files automatically replaced")
        print(f"   • HTML auto-opens for immediate viewing\n")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
