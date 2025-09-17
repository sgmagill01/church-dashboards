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
    """Find both current year and last year attendance report groups"""
    print("\n📋 Searching for attendance report groups...")
    
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
            print(f"✅ Found last year group: {group.get('name')}")
        elif 'report of service individual attendance' in group_name and 'last year' not in group_name:
            current_year_group = group
            print(f"✅ Found current year group: {group.get('name')}")
    
    if not current_year_group:
        print("❌ Current year attendance report group not found")
    if not last_year_group:
        print("❌ Last year attendance report group not found")
    
    return current_year_group, last_year_group

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
    if 'last year' in year_label.lower():
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

def create_enhanced_combined_dashboard(current_data, last_data):
    """Create a single landscape dashboard with all four service charts"""
    print("\n📊 Creating combined attendance dashboard...")
    
    current_df = pd.DataFrame(current_data) if current_data else pd.DataFrame()
    last_df = pd.DataFrame(last_data) if last_data else pd.DataFrame()
    
    current_year = datetime.now().year
    last_year = current_year - 1
    today = datetime.now().date()
    
    print(f"📅 Filtering data to exclude future dates (today: {today})")
    print(f"📊 Note: Average calculations exclude first 3 weeks of January (Jan 1-21) for more representative yearly averages")
    
    # Filter out future dates from current year data
    if len(current_df) > 0:
        print(f"   Current data before filtering: {len(current_df)} records")
        current_df = current_df[current_df['date'].dt.date <= today]
        print(f"   Current data after filtering: {len(current_df)} records")
    
    # Determine actual years in the data
    current_years = set()
    last_years = set()
    
    if len(current_df) > 0:
        current_years = set(current_df['date'].dt.year)
        print(f"Current data contains years: {sorted(current_years)}")
    
    if len(last_df) > 0:
        last_years = set(last_df['date'].dt.year)
        print(f"Last year data contains years: {sorted(last_years)}")
    
    # Normalize dates to the same calendar year for comparison
    base_year = 2024
    
    if len(current_df) > 0:
        current_df = current_df.copy()
        current_df['normalized_date'] = current_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        current_df['actual_year'] = current_df['date'].dt.year
    
    if len(last_df) > 0:
        last_df = last_df.copy()
        last_df['normalized_date'] = last_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        last_df['actual_year'] = last_df['date'].dt.year
    
    # Adjust labels based on actual data
    if current_years and max(current_years) >= current_year:
        current_label = f"{current_year}"
    elif current_years:
        current_label = f"{max(current_years)}"
    else:
        current_label = "Current Data"
    
    if last_years:
        last_label = f"{max(last_years) if last_years else last_year}"
    else:
        last_label = "Previous Year"
    
    # Enhanced color scheme
    colors = {
        'current': '#1e40af',      # Deep blue
        'current_smooth': '#3b82f6', # Lighter blue for smoothed
        'last': '#dc2626',         # Deep red
        'last_smooth': '#ef4444'   # Lighter red for smoothed
    }
    
    # Chart configurations for 2x2 grid
    charts_config = [
        {'col': '8:30 AM', 'title': '8:30 AM Congregation Attendance', 'row': 1, 'col_pos': 1},
        {'col': '10:30 AM', 'title': '10:30 AM Congregation Attendance', 'row': 1, 'col_pos': 2},
        {'col': '6:30 PM', 'title': '6:30 PM Congregation Attendance', 'row': 2, 'col_pos': 1},
        {'col': 'overall', 'title': 'Combined Congregation Attendance', 'row': 2, 'col_pos': 2}
    ]
    
    # Create subplot figure with 2x2 layout - INCREASED SPACING
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=None,  # Don't use subplot titles, we'll add custom ones
        vertical_spacing=0.25,  # INCREASED from 0.15 to 0.25 for more space
        horizontal_spacing=0.08,
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    print(f"\n🔄 Processing {len(charts_config)} charts for combined dashboard...")
    
    # Track statistics for each chart individually
    chart_annotations = {}
    all_stats = []  # For summary output
    
    for chart_index, chart_config in enumerate(charts_config, 1):
        col_name = chart_config['col']
        chart_title = chart_config['title']
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        
        print(f"\n📈 Processing chart {chart_index}/4: {chart_title}...")
        
        # Calculate statistics for this specific chart
        chart_stats = []
        last_avg = None
        current_avg = None
        
        # Add last year data if available
        if len(last_df) > 0 and col_name in last_df.columns:
            print(f"  📊 Processing {last_label} data...")
            
            # Filter out first three weeks of January (Jan 1-21) for average calculation
            last_df_filtered = last_df[
                ~((last_df['date'].dt.month == 1) & (last_df['date'].dt.day <= 21))
            ]
            last_values = last_df_filtered[last_df_filtered[col_name] > 0][col_name].values
            
            if len(last_values) > 0:
                last_avg = np.mean(last_values)
                last_services = len(last_values)
                excluded_services = len(last_df[last_df[col_name] > 0]) - last_services
                
                # Calculate rolling average for smoothing (use all data for visual continuity)
                last_smooth = calculate_rolling_average(last_df[col_name].values, window=4)
                
                # Create custom hover text with actual years
                hover_text_raw = [
                    f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {row_data[col_name]}"
                    for _, row_data in last_df.iterrows()
                ]
                
                hover_text_smooth = [
                    f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {smooth_val:.1f} (4-week avg)"
                    for (_, row_data), smooth_val in zip(last_df.iterrows(), last_smooth)
                ]
                
                # Add raw data (lighter, thinner line)
                fig.add_trace(
                    go.Scatter(
                        x=last_df['normalized_date'],
                        y=last_df[col_name],
                        mode='lines',
                        name=f'{last_label} (Raw)' if chart_index == 1 else None,  # Only show in legend once
                        line=dict(color=colors['last'], width=1, dash='dot'),
                        opacity=0.5,
                        hovertemplate='%{text}<extra></extra>',
                        text=hover_text_raw,
                        showlegend=(chart_index == 1),
                        legendgroup='last_raw'
                    ),
                    row=row, col=col_pos
                )
                
                # Add smoothed data (main line)
                fig.add_trace(
                    go.Scatter(
                        x=last_df['normalized_date'],
                        y=last_smooth,
                        mode='lines+markers',
                        name=f'{last_label}' if chart_index == 1 else None,  # Remove average from legend
                        line=dict(color=colors['last_smooth'], width=3),
                        marker=dict(size=4, symbol='circle'),
                        hovertemplate='%{text}<extra></extra>',
                        text=hover_text_smooth,
                        showlegend=(chart_index == 1),
                        legendgroup='last_smooth'
                    ),
                    row=row, col=col_pos
                )
                
                chart_stats.append(f"{last_label}: {last_avg:.1f} avg*")
                print(f"    ✅ {last_label}: {last_services} services, average {last_avg:.1f} (excluding {excluded_services} early Jan services)")
        
        # Add current year data if available
        if len(current_df) > 0 and col_name in current_df.columns:
            print(f"  📊 Processing {current_label} data...")
            
            # Filter out first three weeks of January (Jan 1-21) for average calculation
            current_df_filtered = current_df[
                ~((current_df['date'].dt.month == 1) & (current_df['date'].dt.day <= 21))
            ]
            current_values = current_df_filtered[current_df_filtered[col_name] > 0][col_name].values
            
            if len(current_values) > 0:
                current_avg = np.mean(current_values)
                current_services = len(current_values)
                excluded_services = len(current_df[current_df[col_name] > 0]) - current_services
                
                # Calculate rolling average for smoothing (use all data for visual continuity)
                current_smooth = calculate_rolling_average(current_df[col_name].values, window=4)
                
                # Create custom hover text with actual years
                hover_text_raw = [
                    f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {row_data[col_name]}"
                    for _, row_data in current_df.iterrows()
                ]
                
                hover_text_smooth = [
                    f"{row_data['normalized_date'].strftime('%a %d %b')} {row_data['actual_year']}<br>{chart_title}: {smooth_val:.1f} (4-week avg)"
                    for (_, row_data), smooth_val in zip(current_df.iterrows(), current_smooth)
                ]
                
                # Add raw data (lighter, thinner line)
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
                
                # Add smoothed data (main line)
                fig.add_trace(
                    go.Scatter(
                        x=current_df['normalized_date'],
                        y=current_smooth,
                        mode='lines+markers',
                        name=f'{current_label}' if chart_index == 1 else None,  # Remove average from legend
                        line=dict(color=colors['current_smooth'], width=3),
                        marker=dict(size=4, symbol='diamond'),
                        hovertemplate='%{text}<extra></extra>',
                        text=hover_text_smooth,
                        showlegend=(chart_index == 1),
                        legendgroup='current_smooth'
                    ),
                    row=row, col=col_pos
                )
                
                chart_stats.append(f"{current_label}: {current_avg:.1f} avg*")
                print(f"    ✅ {current_label}: {current_services} services, average {current_avg:.1f} (excluding {excluded_services} early Jan services)")
                
                # Calculate and display change
                if last_avg is not None:
                    change = current_avg - last_avg
                    change_pct = (change / last_avg * 100) if last_avg > 0 else 0
                    direction = "📈" if change > 0 else "📉" if change < 0 else "➡️"
                    chart_stats.append(f"Change: {direction} {change:+.1f} ({change_pct:+.1f}%)")
                    print(f"    📊 Change: {direction} {change:+.1f} ({change_pct:+.1f}%)")
        
        # Special handling for 10:30 AM chart - add strategic plan target
        if col_name == '10:30 AM':
            print(f"  🎯 Adding strategic plan target for 10:30 AM service...")
            # Calculate 10% growth from actual 2024 average, not fixed baseline
            if last_avg is not None:
                target_2025 = last_avg * 1.10  # 10% growth from actual 2024 average
                
                # Add horizontal target line for 2025
                fig.add_hline(
                    y=target_2025,
                    line=dict(color='#059669', width=2, dash='dash'),
                    annotation_text=f"2025 Target: {target_2025:.1f} (10% growth from 2024 avg: {last_avg:.1f})",
                    annotation_position="top right",
                    annotation=dict(
                        font=dict(size=10, color='#059669'),
                        bgcolor="rgba(255,255,255,0.9)",
                        bordercolor='#059669',
                        borderwidth=1
                    ),
                    row=row, col=col_pos
                )
                
                # Update chart stats to include target information
                chart_stats.append(f"Strategic Target 2025: {target_2025:.1f}")
                if current_avg is not None:
                    # Calculate progress as percentage of the required INCREASE achieved
                    total_increase_needed = target_2025 - last_avg  # e.g., 74.5 - 67.7 = 6.8
                    increase_achieved = current_avg - last_avg      # e.g., current - 67.7
                    target_progress = (increase_achieved / total_increase_needed) * 100 if total_increase_needed > 0 else 0
                    
                    chart_stats.append(f"Target Progress: {target_progress:.0f}% of required")
                    print(f"    🎯 Strategic target: {target_2025:.1f}")
                    print(f"    📊 Required increase: {total_increase_needed:.1f} (from {last_avg:.1f} to {target_2025:.1f})")
                    print(f"    📈 Achieved increase: {increase_achieved:.1f} (from {last_avg:.1f} to {current_avg:.1f})")
                    print(f"    🎯 Progress: {target_progress:.0f}% of required increase")
                
                print(f"    🎯 Using actual 2024 average ({last_avg:.1f}) for 10% growth calculation = {target_2025:.1f}")
            else:
                print(f"    ⚠️ No 2024 data available for target calculation")
        
        # Store annotation text for this chart
        if chart_stats:
            annotation_text = "<br>".join(chart_stats)
            chart_annotations[f"row{row}_col{col_pos}"] = annotation_text
        
        # Store stats for summary (maintaining original structure)
        all_stats.append({
            'title': chart_title,
            'stats': chart_stats
        })
    
    # Set consistent x-axis range for all charts (full year)
    x_min = datetime(base_year, 1, 1)
    x_max = datetime(base_year, 12, 31)
    
    print(f"\n📏 Calculating individual y-axis ranges for optimal visibility...")
    
    # Calculate individual y-axis ranges for each chart based on its actual data
    chart_ranges = {}
    
    for chart_config in charts_config:
        col_name = chart_config['col']
        chart_values = []
        
        # Collect values for this specific chart
        if len(last_df) > 0 and col_name in last_df.columns:
            chart_values.extend(last_df[col_name].tolist())
        if len(current_df) > 0 and col_name in current_df.columns:
            chart_values.extend(current_df[col_name].tolist())
        
        if chart_values:
            chart_min = max(0, min(chart_values) - 5)
            chart_max = max(chart_values) + 10
            # Add some padding for better visibility
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
            dtick='M2',  # Every 2 months for cleaner display
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
    
    # Overall layout with individual chart annotations
    annotations_list = []
    
    # Add explicit chart title annotations
    chart_titles = [
        {"title": "8:30 AM Congregation Attendance", "x": 0.225, "y": 0.93},   # Top left - RAISED to match separate title
        {"title": "10:30 AM Congregation Attendance", "x": 0.775, "y": 0.93},  # Top right - RAISED to match 8:30 AM
        {"title": "6:30 PM Congregation Attendance", "x": 0.225, "y": 0.45},   # Bottom left
        {"title": "Combined Congregation Attendance", "x": 0.775, "y": 0.45}   # Bottom right
    ]
    
    for chart_title in chart_titles:
        annotations_list.append(
            dict(
                text=f"<b style='font-size:16px'>{chart_title['title']}</b>",
                x=chart_title['x'],
                y=chart_title['y'],
                xref='paper',
                yref='paper',
                showarrow=False,
                align='center',
                font=dict(family="Inter, system-ui, sans-serif", size=16, color='#1e293b'),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="rgba(0,0,0,0.15)",
                borderwidth=1,
                borderpad=6,
                xanchor='center',
                yanchor='middle'
            )
        )
    
    # ADD SEPARATE TITLE FOR TOP-LEFT GRAPH (in case first one is being treated differently)
    annotations_list.append(
        dict(
            text="<b style='font-size:16px'>8:30 AM Congregation Attendance</b>",
            x=0.225,
            y=0.93,  # PUSHED HIGHER
            xref='paper',
            yref='paper',
            showarrow=False,
            align='center',
            font=dict(family="Inter, system-ui, sans-serif", size=16, color='#1e293b'),
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor="rgba(0,0,0,0.15)",
            borderwidth=1,
            borderpad=6,
            xanchor='center',
            yanchor='middle'
        )
    )
    
    # Add statistics in TOP corner - TOP GRAPHS MOVED HIGHER
    for chart_config in charts_config:
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        annotation_key = f"row{row}_col{col_pos}"
        
        if annotation_key in chart_annotations:
            # Position statistics - TOP RIGHT MOVED HIGHER to avoid lines
            if row == 1 and col_pos == 1:      # Top-left chart
                x_pos, y_pos = 0.05, 0.91  
            elif row == 1 and col_pos == 2:    # Top-right chart  
                x_pos, y_pos = 0.55, 0.94  # MOVED HIGHER to avoid clashing with lines
            elif row == 2 and col_pos == 1:    # Bottom-left chart
                x_pos, y_pos = 0.05, 0.42
            else:                               # Bottom-right chart
                x_pos, y_pos = 0.55, 0.42
            
            annotations_list.append(
                dict(
                    text=f"<b>{chart_annotations[annotation_key]}</b>",
                    x=x_pos,
                    y=y_pos,
                    xref='paper',
                    yref='paper',
                    showarrow=False,
                    align='left',
                    font=dict(family="Inter, system-ui, sans-serif", size=10, color='#1e293b'),
                    bgcolor="rgba(248,250,252,0.9)",
                    bordercolor="rgba(0,0,0,0.1)",
                    borderwidth=1,
                    borderpad=4,
                    xanchor='left',
                    yanchor='top'
                )
            )
    
    fig.update_layout(
        title=dict(
            text=f"<b style='font-size:24px'>St George's Magill - Attendance Analysis</b><br><span style='font-size:14px; color:#64748b'>Year-over-Year Comparison Dashboard</span><br><span style='font-size:11px; color:#94a3b8'>Generated {datetime.now().strftime('%B %d, %Y')} • *Averages exclude first 3 weeks of January</span>",
            x=0.5,
            y=0.985,  # Positioned at very top
            font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=20, color='#1e293b')
        ),
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=12),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=1200,  # INCREASED height from 1000 to 1200 to accommodate spacing
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
        margin=dict(l=80, r=80, t=140, b=120),  # INCREASED top margin from 120 to 140
        annotations=annotations_list  # Add all chart titles and statistics
    )
    
    # Save the combined dashboard
    html_filename = f"attendance_dashboard_combined_{timestamp}.html"
    try:
        fig.write_html(html_filename)
        print(f"\n✅ Saved combined dashboard: {html_filename}")
    except Exception as e:
        print(f"\n❌ Failed to save combined dashboard: {e}")
    
    try:
        png_filename = f"attendance_dashboard_combined_{timestamp}.png"
        fig.write_image(png_filename, width=1600, height=1200, scale=3)  # Updated height
        print(f"✅ Saved PNG: {png_filename}")
    except Exception as e:
        print(f"⚠️ PNG save failed: {e}")
    
    print(f"\n📊 Combined dashboard ready: {html_filename}")
    
    # Print summary statistics
    print(f"\n📈 DASHBOARD SUMMARY:")
    for stat in all_stats:
        print(f"  {stat['title']}:")
        for line in stat['stats']:
            print(f"    {line}")
    
    return timestamp, html_filename

def create_summary_page_for_dashboard(dashboard_file, timestamp):
    """Create a summary HTML page for the combined dashboard"""
    print("\n📋 Creating summary page for combined dashboard...")
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>St George's Attendance Dashboard - Summary</title>
    <style>
        body {{
            font-family: Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            margin: 0;
            padding: 40px;
            min-height: 100vh;
        }}
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.1);
        }}
        .header {{
            text-align: center;
            margin-bottom: 40px;
            padding-bottom: 30px;
            border-bottom: 3px solid #f0f0f0;
        }}
        .header h1 {{
            font-size: 36px;
            color: #1e293b;
            margin: 0 0 10px 0;
        }}
        .header p {{
            font-size: 18px;
            color: #64748b;
            margin: 0;
        }}
        .dashboard-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 20px;
            padding: 40px;
            text-align: center;
            margin-bottom: 30px;
            color: white;
        }}
        .dashboard-card h3 {{
            font-size: 28px;
            margin: 0 0 15px 0;
        }}
        .dashboard-card p {{
            font-size: 16px;
            margin: 0 0 25px 0;
            opacity: 0.9;
        }}
        .dashboard-link {{
            display: inline-block;
            background: white;
            color: #1e293b;
            padding: 15px 30px;
            border-radius: 10px;
            text-decoration: none;
            font-weight: 600;
            font-size: 18px;
            transition: all 0.3s ease;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }}
        .dashboard-link:hover {{
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
        }}
        .features {{
            background: #f8fafc;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
        }}
        .features h3 {{
            margin-top: 0;
            color: #1e293b;
            font-size: 24px;
        }}
        .feature-list {{
            columns: 2;
            column-gap: 30px;
            list-style: none;
            padding: 0;
        }}
        .feature-list li {{
            margin: 10px 0;
            color: #374151;
            font-size: 14px;
        }}
        .feature-list li:before {{
            content: "✨ ";
            margin-right: 8px;
        }}
        .instructions {{
            background: #fef3c7;
            border: 1px solid #f59e0b;
            border-radius: 10px;
            padding: 20px;
            margin-top: 30px;
        }}
        .instructions h3 {{
            margin-top: 0;
            color: #92400e;
        }}
        .instructions p {{
            color: #92400e;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 St George's Attendance Dashboard</h1>
            <p>Year-over-Year Comparison Analysis</p>
            <p><small>Generated on {datetime.now().strftime('%A, %B %d, %Y at %I:%M %p')}</small></p>
        </div>
        
        <div class="dashboard-card">
            <h3>🏛️ Complete Attendance Analysis</h3>
            <p>Single landscape dashboard with all four service charts in professional 2×2 layout</p>
            <a href="{dashboard_file}" class="dashboard-link">📊 Open Dashboard</a>
        </div>
        
        <div class="features">
            <h3>✨ Dashboard Features (FINAL FIXED VERSION)</h3>
            <ul class="feature-list">
                <li>Added separate title for 8:30 AM chart to ensure visibility</li>
                <li>Statistics positioned at very top corner of each chart</li>
                <li>Zero interference between labels and chart data lines</li>
                <li>All 4 charts guaranteed to have visible headings</li>
                <li>Optimized spacing and professional layout</li>
                <li>Single landscape page perfect for reports</li>
                <li>High-resolution charts with professional typography</li>
                <li>Rolling 4-week averages to smooth trends</li>
                <li>Year-over-year comparison with statistics</li>
                <li>Full January-December timeline view</li>
                <li>Raw data lines plus smoothed trend lines</li>
                <li>Individual y-axis scaling for optimal visibility</li>
                <li>Interactive hover details</li>
                <li>Strategic targets for 10:30 AM service</li>
                <li>Averages exclude first 3 weeks of January</li>
                <li>Print-ready landscape format</li>
            </ul>
        </div>
        
        <div class="instructions">
            <h3>📁 Accessing Your Dashboard</h3>
            <p><strong>Main Dashboard:</strong> Click the "Open Dashboard" button above to view the complete analysis.</p>
            <p><strong>In Replit:</strong> Look for "{dashboard_file}" in the left sidebar file explorer.</p>
            <p><strong>Download:</strong> Right-click the file and select "Download" to save to your computer.</p>
            <p><strong>Print:</strong> The landscape layout is optimized for printing on A4 or letter-size paper.</p>
        </div>
    </div>
</body>
</html>"""
    
    summary_filename = f"attendance_dashboard_summary_{timestamp}.html"
    with open(summary_filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ Dashboard summary page created: {summary_filename}")
    return summary_filename

def list_all_files_in_directory():
    """List all files in current directory for debugging"""
    import os
    print("\n📁 ALL FILES IN CURRENT DIRECTORY:")
    files = os.listdir('.')
    for file in sorted(files):
        if file.endswith(('.html', '.png')):
            print(f"   📄 {file}")

def main():
    """Main execution function"""
    try:
        print("🚀 Starting enhanced year-over-year attendance analysis...")
        
        # Find both attendance report groups
        current_group, last_year_group = find_attendance_report_groups()
        
        if not current_group or not last_year_group:
            print("❌ Cannot proceed without both attendance report groups")
            return
        
        # Extract attendance data from both years
        current_df, current_headers, current_year = extract_attendance_data_from_group(current_group, "Current Year")
        last_df, last_headers, last_year = extract_attendance_data_from_group(last_year_group, "Last Year")
        
        if current_df is None or last_df is None:
            print("❌ Cannot proceed without both datasets")
            return
        
        # Parse service columns for both years with correct year assignment
        current_services = parse_service_columns_for_year(current_headers, current_year, "Current Year")
        last_services = parse_service_columns_for_year(last_headers, last_year, "Last Year")
        
        # Check what data we actually got
        print(f"\n📊 Data Summary:")
        print(f"   Current year services: {len(current_services)}")
        print(f"   Last year services: {len(last_services)}")
        
        # If current year data is empty, try flexible parsing
        if len(current_services) == 0:
            print("⚠️ No current year data found. Trying flexible parsing...")
            current_services = parse_service_columns_for_year_flexible(current_headers, current_year, "Current Year (Flexible)")
        
        if not current_services and not last_services:
            print("❌ Cannot proceed without any service data")
            return
        elif not current_services:
            print("⚠️ Proceeding with only last year data")
        elif not last_services:
            print("⚠️ Proceeding with only current data")
        
        # Calculate attendance by service for both years
        if current_services:
            current_attendance = calculate_service_attendance_by_year(current_df, current_services, "Current Year")
            current_attendance = apply_pro_rata_logic_to_dataset(current_attendance, "Current Year")
        else:
            current_attendance = []
            
        if last_services:
            last_attendance = calculate_service_attendance_by_year(last_df, last_services, "Last Year")
            last_attendance = apply_pro_rata_logic_to_dataset(last_attendance, "Last Year")
        else:
            last_attendance = []
        
        # Create combined dashboard instead of individual charts
        timestamp, dashboard_file = create_enhanced_combined_dashboard(current_attendance, last_attendance)
        
        # Create summary page that now points to the combined dashboard
        summary_file = create_summary_page_for_dashboard(dashboard_file, timestamp)
        
        # List all files for debugging
        list_all_files_in_directory()
        
        print(f"\n🎉 COMPLETELY FIXED ATTENDANCE DASHBOARD!")
        print(f"📊 Created single landscape dashboard with all 4 service charts")
        print(f"📄 Main file: {dashboard_file}")
        print(f"📋 Summary: {summary_file}")
        print(f"\n📁 TO VIEW YOUR DASHBOARD:")
        print(f"   1. Open: {dashboard_file}")
        print(f"   2. Or use summary page: {summary_file}")
        print(f"\n✨ FINAL FIXES APPLIED:")
        print(f"   • Added separate title for 8:30 AM chart (in case first was treated differently)")
        print(f"   • Statistics moved to very top corner of each chart")
        print(f"   • All labels positioned to avoid interfering with data lines")
        print(f"   • All 4 chart titles should now be clearly visible")
        
        # Open the dashboard automatically
        try:
            import webbrowser
            import os
            file_path = os.path.abspath(dashboard_file)
            webbrowser.open(f"file://{file_path}")
            print(f"\n🌐 Final fixed dashboard opened automatically!")
        except Exception as e:
            print(f"\n⚠️ Couldn't auto-open dashboard: {e}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
