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

# Word-based ministry positions (from using_gifts_dashboard.py)
WORD_BASED_POSITIONS = {
    'Acoustic Guitar', 'Bible Reader', 'BSG Leader', 'Band Leader', 'Bass', 'Cajon', 'Comm. Celebrant',
    'Cornet', 'Drums', 'Electric Guitar', 'Flute', 'Gospel Story Teller', 'Jnr Kids Assist',
    'Jnr Kids Leader', 'Keyboard', 'Kids Church Assist', 'Kids Church Leader',
    'Kids Club Assistant', 'Kids Talk', 'Oboe', 'Organ', 'Prayer Leader', 'Preacher',
    'Service Leader', 'Vocals', 'Youth Assist', 'Youth Group Talk', 'Youth Leader'
}

print("🙌 ST GEORGE'S MAGILL - USING GIFTS ANALYSIS")
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

def find_serving_report_groups():
    """Find current year, last year, and two years ago serving report groups"""
    print("\n📋 Searching for serving report groups...")

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
        if 'report of two years ago serving' in group_name:
            two_years_ago_group = group
            print(f"✅ Found two years ago group: {group.get('name')}")
        elif 'report of last year serving' in group_name:
            last_year_group = group
            print(f"✅ Found last year group: {group.get('name')}")
        elif 'report of serving' in group_name and 'last year' not in group_name and 'two years ago' not in group_name:
            current_year_group = group
            print(f"✅ Found current year group: {group.get('name')}")

    if not current_year_group:
        print("❌ Current year serving report group not found")
    if not last_year_group:
        print("❌ Last year serving report group not found")
    if not two_years_ago_group:
        print("❌ Two years ago serving report group not found")

    return current_year_group, last_year_group, two_years_ago_group

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
            print(f"✅ Found two years ago attendance group: {group.get('name')}")
        elif 'report of last year service individual attendance' in group_name:
            last_year_group = group
            print(f"✅ Found last year attendance group: {group.get('name')}")
        elif 'report of service individual attendance' in group_name and 'last year' not in group_name and 'two years ago' not in group_name:
            current_year_group = group
            print(f"✅ Found current year attendance group: {group.get('name')}")

    return current_year_group, last_year_group, two_years_ago_group

def fetch_category_lookup():
    """Build category_id -> category_name mapping (from applying_word.py)"""
    resp = make_request('people/categories/getAll', {})
    if not resp or not resp.get('categories'):
        return {}

    cats = resp['categories'].get('category', [])
    if not isinstance(cats, list):
        cats = [cats] if cats else []

    lookup = {}
    for c in cats:
        cid = c.get('id')
        cname = c.get('name', '').strip()
        if cid:
            lookup[cid] = cname

    return lookup

def get_regulars_at_year_start(target_year):
    """
    Get count of 'regulars' (Congregation_ + RosteredMember_) at Jan 1 of target_year.
    Uses the same method as applying_word.py - scans all people and counts those in
    Congregation_ or RosteredMember_ categories.

    Returns the count of people in Congregation_ or RosteredMember_ categories.
    """
    print(f"\n🔍 Calculating regulars at Jan 1, {target_year}...")

    def _truthy(v):
        # Handles 1, "1", True, "true", "yes", "y", "on" (case-insensitive)
        return str(v).strip().lower() in {"1", "true", "yes", "y", "on"}

    try:
        # Build category_id -> name lookup (like applying_word.py)
        cat_lookup = fetch_category_lookup()
        if not cat_lookup:
            print(f"   ⚠️ No categories returned; cannot determine regulars for {target_year}")
            return None

        page = 1
        page_size = 1000
        regular_ids = set()
        counts_by_cat = {"Congregation_": 0, "RosteredMember_": 0}
        total_scanned = 0

        print(f"   🧑‍🤝‍🧑 Scanning people for regulars (Congregation_ + RosteredMember_)...")

        while True:
            resp = make_request('people/getAll', {
                'page': page,
                'page_size': page_size
            })
            if not resp or not resp.get('people'):
                break

            people = resp['people'].get('person', [])
            if not isinstance(people, list):
                people = [people] if people else []

            batch_count = len(people)
            total_scanned += batch_count

            for p in people:
                # Exclude deceased and archived (like applying_word.py)
                if _truthy(p.get('deceased', 0)):
                    continue
                status = (p.get('status') or "").strip().lower()
                if status == 'deceased':
                    continue
                if _truthy(p.get('archived', 0)):
                    continue

                cid = p.get('category_id')
                if not cid:
                    continue

                cname = cat_lookup.get(cid, '')
                if cname == 'Congregation_' or cname == 'RosteredMember_':
                    pid = p.get('id')
                    if pid:
                        if pid not in regular_ids and cname in counts_by_cat:
                            counts_by_cat[cname] += 1
                        regular_ids.add(pid)

            # Pagination
            paging = resp.get('paging') or {}
            if 'pages' in paging:
                try:
                    pages_reported_by_api = int(paging.get('pages'))
                    if page >= pages_reported_by_api:
                        break
                except Exception:
                    pass

            if batch_count < page_size:
                break

            page += 1

        regular_count = len(regular_ids)
        print(f"   ✅ Found {regular_count} regulars (Congregation_: {counts_by_cat['Congregation_']}, RosteredMember_: {counts_by_cat['RosteredMember_']})")
        print(f"   📊 Scanned {total_scanned} total people")

        return regular_count

    except Exception as e:
        print(f"   ❌ Error calculating regulars for {target_year}: {e}")
        import traceback
        traceback.print_exc()
        return None

def find_new_serving_members_reports():
    """Find all three people category change reports (current, last year, two years ago)"""
    print("\n📋 Searching for people category change report groups...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None, None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_report = None
    last_report = None
    two_years_report = None

    for group in groups:
        group_name = group.get('name', '').lower()
        if group_name == 'report of people category change':
            current_report = group
            print(f"✅ Found current year report: {group.get('name')}")
        elif group_name == 'report of last year people category change':
            last_report = group
            print(f"✅ Found last year report: {group.get('name')}")
        elif group_name == 'report of two years ago people category change':
            two_years_report = group
            print(f"✅ Found two years ago report: {group.get('name')}")

    if not current_report:
        print("⚠️ 'Report of People Category Change' group not found")
    if not last_report:
        print("⚠️ 'Report of Last Year People Category Change' group not found")
    if not two_years_report:
        print("⚠️ 'Report of Two Years Ago People Category Change' group not found")

    return current_report, last_report, two_years_report

def extract_new_serving_members_data(group):
    """Extract new serving members data from the report group"""
    print("\n🔍 Extracting new serving members data from report...")

    if not group:
        print("❌ No report group provided")
        return []

    # Extract URL from group location field
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print("❌ No URL found in report group location")
        return []

    print(f"✅ Found report URL")

    # Fetch and parse HTML
    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find the report table
        tables = soup.find_all('table')
        for table in tables:
            header_row = table.find('tr')
            if header_row:
                headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
                header_text = ' '.join(headers).lower()

                # Look for the people category history table
                if 'person' in header_text and 'change to' in header_text:
                    print(f"✅ Found new serving members table")

                    new_serving_members = []
                    for row in table.find_all('tr')[1:]:  # Skip header
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= len(headers):
                            row_data = {}
                            for i, header in enumerate(headers):
                                if i < len(cells):
                                    row_data[header] = cells[i].get_text(strip=True)

                            # Only include people who changed TO RosteredMember_
                            change_to = row_data.get('Change To', '')
                            if change_to == 'RosteredMember_':
                                person_name = row_data.get('Person', '')
                                member_id = row_data.get('Member ID', '')
                                if person_name:
                                    new_serving_members.append({
                                        'name': person_name,
                                        'member_id': member_id,
                                        'date_changed': row_data.get('Date', ''),
                                        'change_from': row_data.get('Change From', '')
                                    })

                    print(f"✅ Extracted {len(new_serving_members)} new serving members")
                    return new_serving_members

        print("❌ New serving members table not found in report")
        return []

    except Exception as e:
        print(f"❌ Error extracting new serving members data: {e}")
        return []

def get_detailed_person_info(person_id):
    """Get detailed person information including full volunteer positions"""
    try:
        # Try without 'volunteer' field first - it may be included by default
        response = make_request('people/getInfo', {
            'id': person_id,
            'fields': ['demographics', 'departments', 'locations', 'family']
        })
        if response and 'person' in response:
            person_data = response['person']
            # API returns person as a list with one dict
            if isinstance(person_data, list) and person_data:
                return person_data[0]  # Return the first (and only) person dict
            return person_data
        return None
    except Exception as e:
        print(f"    ⚠️ Error getting detailed info for person {person_id}: {e}")
    return None

def analyze_new_serving_members_word_based(new_serving_members, all_people, groups):
    """Analyze which new serving members are in word-based ministry using Member ID matching"""
    print(f"\n📖 Analyzing {len(new_serving_members)} new serving members for word-based ministry...")

    # Create lookup of all people by Member ID
    people_by_id = {}
    for person in all_people:
        pid = person.get('id')
        if pid:
            people_by_id[pid] = person

    new_word_based_count = 0
    new_word_based_details = []
    not_found_count = 0

    for new_member in new_serving_members:
        member_name = new_member['name']
        member_id = new_member.get('member_id', '')

        if not member_id:
            not_found_count += 1
            continue

        # Get detailed person info from API
        person_data = get_detailed_person_info(member_id)

        if not person_data:
            not_found_count += 1
            continue

        # Check volunteer positions
        # The 'volunteer' field is just a flag (0/1)
        # The actual departments data is in the 'departments' field
        is_word_based = False
        word_based_positions_found = []

        departments_data = person_data.get('departments')

        if departments_data:
            # departments_data is a dict with 'department' key
            if isinstance(departments_data, dict):
                dept_list = departments_data.get('department', [])
                if not isinstance(dept_list, list):
                    dept_list = [dept_list] if dept_list else []

                for dept in dept_list:
                    # Positions can be directly under department or under sub_departments
                    # Check sub_departments first (most common structure)
                    sub_depts_data = dept.get('sub_departments', {})
                    if sub_depts_data:
                        sub_dept_list = sub_depts_data.get('sub_department', [])
                        if not isinstance(sub_dept_list, list):
                            sub_dept_list = [sub_dept_list] if sub_dept_list else []

                        for sub_dept in sub_dept_list:
                            positions = sub_dept.get('positions', {})
                            if positions:
                                pos_list = positions.get('position', [])
                                if not isinstance(pos_list, list):
                                    pos_list = [pos_list] if pos_list else []

                                for pos in pos_list:
                                    pos_name = pos.get('name', '').strip()
                                    if pos_name in WORD_BASED_POSITIONS:
                                        is_word_based = True
                                        word_based_positions_found.append(pos_name)

                    # Also check for positions directly under department
                    positions = dept.get('positions', {})
                    if positions:
                        pos_list = positions.get('position', [])
                        if not isinstance(pos_list, list):
                            pos_list = [pos_list] if pos_list else []

                        for pos in pos_list:
                            pos_name = pos.get('name', '').strip()
                            if pos_name in WORD_BASED_POSITIONS:
                                is_word_based = True
                                word_based_positions_found.append(pos_name)

        if is_word_based:
            new_word_based_count += 1
            new_word_based_details.append({
                'name': member_name,
                'member_id': member_id,
                'positions': word_based_positions_found,
                'date_changed': new_member.get('date_changed', ''),
                'change_from': new_member.get('change_from', '')
            })

    print(f"   Total new serving members: {len(new_serving_members)}")
    print(f"   In word-based ministry: {new_word_based_count}")

    return new_word_based_count, new_word_based_details

def parse_column_header(header):
    """Parse column headers like '15/01 8:30 AM' or '08/01 10:30 AM'"""

    # Extract time (look for patterns like 8:30 AM, 10:30 AM, etc.)
    time_match = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', header, re.IGNORECASE)
    if not time_match:
        return None

    time_str = f"{time_match.group(1)} {time_match.group(2).upper()}"

    # Extract date - handle DD/MM format
    date_match = re.search(r'(\d{1,2})/(\d{1,2})', header)
    if date_match:
        day = int(date_match.group(1))
        month = int(date_match.group(2))

        # Normalize service time for consistency
        if '8:30' in time_str:
            normalized_time = '8:30'
        elif '9:30' in time_str:
            normalized_time = '9:30'  # Combined service
        elif '10:30' in time_str:
            normalized_time = '10:30'
        elif '6:30' in time_str or '6:00' in time_str:
            normalized_time = '6:30'
        else:
            normalized_time = 'Other'

        return {
            'time': time_str,
            'normalized_time': normalized_time,
            'day': day,
            'month': month,
            'original_header': header
        }

    return None

def extract_serving_data_from_group(group, year_label):
    """Extract serving data from a specific group"""
    print(f"\n🔍 Extracting {year_label} serving data...")

    # Determine target year based on group name
    if 'two years ago' in year_label.lower():
        target_year = datetime.now().year - 2
    elif 'last year' in year_label.lower():
        target_year = datetime.now().year - 1
    else:
        target_year = datetime.now().year

    print(f"   Target year: {target_year}")

    # Extract URL from group location field
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print(f"❌ No URL found for {year_label} group")
        return None, None, None

    print(f"✅ Found {year_label} report URL")

    # Fetch and parse HTML
    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find serving table
        tables = soup.find_all('table')
        for table in tables:
            header_row = table.find('tr')
            if header_row:
                headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
                header_text = ' '.join(headers).lower()

                # Look for volunteer positions table
                if 'volunteers' in header_text or any('/' in h for h in headers):
                    print(f"✅ Found {year_label} serving table")

                    serving_records = []
                    for row in table.find_all('tr')[1:]:  # Skip header
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= len(headers):
                            row_data = {}
                            for i, header in enumerate(headers):
                                if i < len(cells):
                                    row_data[header] = cells[i].get_text(strip=True)
                            if any(row_data.values()):
                                serving_records.append(row_data)

                    print(f"✅ Extracted {len(serving_records)} {year_label} serving records")
                    return pd.DataFrame(serving_records), headers, target_year

        print(f"❌ {year_label} serving table not found")
        return None, None, None

    except Exception as e:
        print(f"❌ Error extracting {year_label} data: {e}")
        return None, None, None

def parse_service_columns_for_year(headers, target_year, year_label):
    """Parse service columns for a specific year"""
    print(f"\n📅 Analyzing {year_label} service columns for {target_year}...")

    service_columns = []
    unparseable_count = 0

    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['volunteers', 'position', 'category', 'email', 'phone', 'member id']):
            continue

        # Parse the column header
        parsed = parse_column_header(header)
        if not parsed:
            unparseable_count += 1
            continue

        try:
            # Create date using the target year
            service_date = datetime(target_year, parsed['month'], parsed['day'])

            service_columns.append({
                'header': header,
                'date': service_date,
                'time': parsed['normalized_time'],
                'year': target_year
            })

        except ValueError as e:
            # Invalid date (e.g., Feb 30)
            continue

    # Sort by date
    service_columns.sort(key=lambda x: x['date'])

    print(f"✅ Found {len(service_columns)} valid service columns")
    print(f"   Couldn't parse {unparseable_count} headers")

    return service_columns

def is_word_based_position(position_text):
    """Check if a position (or combined positions) contains any word-based ministry role"""
    if not position_text or pd.isna(position_text) or position_text == '':
        return False

    # Split by common delimiters (/, comma, etc.)
    positions = re.split(r'[/,]', position_text)

    # Check if any position matches word-based ministry
    for pos in positions:
        pos_clean = pos.strip()
        if pos_clean in WORD_BASED_POSITIONS:
            return True

    return False

def calculate_cumulative_servers(df, service_columns, year_label):
    """
    Calculate cumulative unique servers by week for each congregation.
    Uses an efficient algorithm with sets to avoid recounting.
    Also tracks word-based ministry servers separately.
    """
    print(f"\n📊 Calculating {year_label} cumulative serving participation...")

    # Group services by date and time
    services_by_date = {}

    for svc in service_columns:
        date = svc['date']
        time = svc['time']
        header = svc['header']

        if date not in services_by_date:
            services_by_date[date] = []

        services_by_date[date].append({
            'time': time,
            'header': header
        })

    # Sort dates
    sorted_dates = sorted(services_by_date.keys())

    # Track cumulative servers using sets (efficient!)
    cumulative_servers = {
        'overall': set(),
        '8:30': set(),
        '10:30': set(),
        '6:30': set(),
        'word_based': set()  # Track word-based ministry separately
    }

    # Build cumulative data
    cumulative_data = []

    for date in sorted_dates:
        services = services_by_date[date]

        # Process each service on this date
        for service in services:
            time = service['time']
            header = service['header']

            # Skip if time is not one of our main congregations
            if time not in ['8:30', '10:30', '6:30']:
                continue

            # Find who served at this service
            if header in df.columns:
                # Get people who served (non-empty cells)
                servers = df[df[header].notna() & (df[header] != '')].index.tolist()

                # Add to cumulative sets
                for server_idx in servers:
                    # Use the person's name as identifier
                    person_name = f"{df.loc[server_idx, 'Volunteers'] if 'Volunteers' in df.columns else server_idx}"

                    cumulative_servers['overall'].add(person_name)
                    cumulative_servers[time].add(person_name)

                    # Check if this person served in a word-based ministry position
                    position_text = df.loc[server_idx, header]
                    if is_word_based_position(position_text):
                        cumulative_servers['word_based'].add(person_name)

        # Record cumulative counts for this date
        cumulative_data.append({
            'date': date,
            'overall': len(cumulative_servers['overall']),
            '8:30': len(cumulative_servers['8:30']),
            '10:30': len(cumulative_servers['10:30']),
            '6:30': len(cumulative_servers['6:30']),
            'word_based': len(cumulative_servers['word_based'])
        })

    print(f"✅ Processed {len(sorted_dates)} service dates")
    print(f"   Total unique servers by end of year: {len(cumulative_servers['overall'])}")
    print(f"     8:30 AM: {len(cumulative_servers['8:30'])}")
    print(f"     10:30 AM: {len(cumulative_servers['10:30'])}")
    print(f"     6:30 PM: {len(cumulative_servers['6:30'])}")
    print(f"     Word-based ministry: {len(cumulative_servers['word_based'])}")

    return cumulative_data

def create_serving_dashboard(current_data, last_data, two_years_data,
                             current_regulars=None, last_regulars=None, two_years_regulars=None,
                             current_recruitment=0, last_recruitment=0, two_years_recruitment=0):
    """Create a 2x3 dashboard showing cumulative serving participation"""
    print(f"\n📊 Creating using gifts dashboard...")
    print(f"   Using regulars at year start as denominators (like applying_word.py)")

    current_df = pd.DataFrame(current_data) if current_data else pd.DataFrame()
    last_df = pd.DataFrame(last_data) if last_data else pd.DataFrame()
    two_years_df = pd.DataFrame(two_years_data) if two_years_data else pd.DataFrame()

    current_year = datetime.now().year
    last_year = current_year - 1
    two_years_ago = current_year - 2
    today = datetime.now().date()

    print(f"📅 Filtering data to exclude future dates (today: {today})")

    # Load strategic plan targets from config
    try:
        from config import USING_GIFTS_TARGETS
        pct_serving_target = USING_GIFTS_TARGETS['pct_congregation_serving']
        print(f"✅ Loaded strategic plan targets from config.py")
    except (ImportError, KeyError) as e:
        print(f"⚠️ Could not load strategic plan targets: {e}")
        pct_serving_target = None

    # Filter out future dates from current year data
    if len(current_df) > 0:
        print(f"   Current data before filtering: {len(current_df)} records")
        current_df = current_df[current_df['date'].dt.date <= today]
        print(f"   Current data after filtering: {len(current_df)} records")

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

    if len(two_years_df) > 0:
        two_years_df = two_years_df.copy()
        two_years_df['normalized_date'] = two_years_df['date'].apply(
            lambda x: datetime(base_year, x.month, x.day)
        )
        two_years_df['actual_year'] = two_years_df['date'].dt.year

    # Enhanced color scheme
    colors = {
        'current': '#1e40af',
        'last': '#dc2626',
        'two_years': '#059669'
    }

    # Chart configurations for 2x3 grid
    charts_config = [
        {'col': '8:30', 'title': '8:30 AM Congregation - Cumulative Servers', 'row': 1, 'col_pos': 1, 'chart_type': 'count'},
        {'col': '10:30', 'title': '10:30 AM Congregation - Cumulative Servers', 'row': 1, 'col_pos': 2, 'chart_type': 'count'},
        {'col': 'overall', 'title': 'Combined - Cumulative Servers', 'row': 2, 'col_pos': 1, 'chart_type': 'count'},
        {'col': 'word_based', 'title': 'Combined - Word-Based', 'row': 2, 'col_pos': 2, 'chart_type': 'word_based'},
        {'col': 'overall', 'title': 'Combined - % of Congregation Serving', 'row': 3, 'col_pos': 1, 'chart_type': 'percentage'},
        {'col': 'recruitment', 'title': 'Word-Based Ministry Recruitment', 'row': 3, 'col_pos': 2, 'chart_type': 'recruitment'}
    ]

    # Pre-calculate statistics for annotations
    print(f"\n🔄 Pre-calculating statistics for charts...")

    all_stats = []

    for chart_config in charts_config:
        col_name = chart_config['col']
        chart_title = chart_config['title']
        chart_type = chart_config.get('chart_type', 'count')

        print(f"\n📈 Pre-calculating stats for {chart_title}...")

        stats_parts = []

        # Get final values for each year
        if chart_type == 'percentage':
            # Calculate percentages
            if len(two_years_df) > 0 and col_name in two_years_df.columns and two_years_regulars:
                two_years_final = two_years_df[col_name].iloc[-1] if len(two_years_df) > 0 else 0
                two_years_pct = (two_years_final / two_years_regulars) * 100
                stats_parts.append(f"{two_years_ago}: {two_years_pct:.1f}%")

            if len(last_df) > 0 and col_name in last_df.columns and last_regulars:
                last_final = last_df[col_name].iloc[-1] if len(last_df) > 0 else 0
                last_pct = (last_final / last_regulars) * 100
                stats_parts.append(f"{last_year}: {last_pct:.1f}%")

            if len(current_df) > 0 and col_name in current_df.columns and current_regulars:
                current_final = current_df[col_name].iloc[-1] if len(current_df) > 0 else 0
                current_pct = (current_final / current_regulars) * 100
                stats_parts.append(f"{current_year} YTD: {current_pct:.1f}%")
        elif chart_type == 'recruitment':
            # Recruitment stats
            stats_parts.append(f"{two_years_ago}: {two_years_recruitment}")
            stats_parts.append(f"{last_year}: {last_recruitment}")
            stats_parts.append(f"{current_year}: {current_recruitment}")
        else:
            # Regular counts
            if len(two_years_df) > 0 and col_name in two_years_df.columns:
                two_years_final = two_years_df[col_name].iloc[-1] if len(two_years_df) > 0 else 0
                stats_parts.append(f"{two_years_ago}: {two_years_final:.0f}")

            if len(last_df) > 0 and col_name in last_df.columns:
                last_final = last_df[col_name].iloc[-1] if len(last_df) > 0 else 0
                stats_parts.append(f"{last_year}: {last_final:.0f}")

            if len(current_df) > 0 and col_name in current_df.columns:
                current_final = current_df[col_name].iloc[-1] if len(current_df) > 0 else 0
                stats_parts.append(f"{current_year} YTD: {current_final:.0f}")

        all_stats.append({
            'title': chart_title,
            'stats': stats_parts
        })

    # Create subplot figure with 2x3 layout (3 rows, 2 columns)
    fig = make_subplots(
        rows=3, cols=2,
        subplot_titles=None,
        vertical_spacing=0.12,
        horizontal_spacing=0.08,
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )

    print(f"\n🔄 Processing {len(charts_config)} charts for combined dashboard...")

    for chart_index, chart_config in enumerate(charts_config, 1):
        col_name = chart_config['col']
        chart_title = chart_config['title']
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        chart_type = chart_config.get('chart_type', 'count')

        print(f"\n📈 Processing chart {chart_index}/{len(charts_config)}: {chart_title}...")

        # Skip line chart rendering for recruitment chart (it will be rendered as bar chart later)
        if chart_type == 'recruitment':
            print(f"  ⏭️ Skipping line chart rendering (will render as bar chart)")

        # Add two years ago data if available
        if len(two_years_df) > 0 and col_name in two_years_df.columns and chart_type != 'recruitment':
            print(f"  📊 Processing {two_years_ago} data...")

            # Calculate y values based on chart type
            if chart_type == 'percentage' and two_years_regulars:
                y_values = (two_years_df[col_name] / two_years_regulars) * 100
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>📊 {val:.1f}%"
                             for date, val in zip(two_years_df['normalized_date'], y_values)]
            else:
                y_values = two_years_df[col_name]
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>👥 {val:.0f} unique servers"
                             for date, val in zip(two_years_df['normalized_date'], y_values)]

            fig.add_trace(
                go.Scatter(
                    x=two_years_df['normalized_date'],
                    y=y_values,
                    mode='lines+markers',
                    name=f'{two_years_ago}' if chart_index == 1 else None,
                    line=dict(color=colors['two_years'], width=3),
                    marker=dict(size=4, symbol='diamond'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text,
                    showlegend=(chart_index == 1),
                    legendgroup='two_years'
                ),
                row=row, col=col_pos
            )

        # Add last year data if available
        if len(last_df) > 0 and col_name in last_df.columns and chart_type != 'recruitment':
            print(f"  📊 Processing {last_year} data...")

            # Calculate y values based on chart type
            if chart_type == 'percentage' and last_regulars:
                y_values = (last_df[col_name] / last_regulars) * 100
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>📊 {val:.1f}%"
                             for date, val in zip(last_df['normalized_date'], y_values)]
            else:
                y_values = last_df[col_name]
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>👥 {val:.0f} unique servers"
                             for date, val in zip(last_df['normalized_date'], y_values)]

            fig.add_trace(
                go.Scatter(
                    x=last_df['normalized_date'],
                    y=y_values,
                    mode='lines+markers',
                    name=f'{last_year}' if chart_index == 1 else None,
                    line=dict(color=colors['last'], width=3),
                    marker=dict(size=4, symbol='diamond'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text,
                    showlegend=(chart_index == 1),
                    legendgroup='last'
                ),
                row=row, col=col_pos
            )

        # Add current year data if available
        if len(current_df) > 0 and col_name in current_df.columns and chart_type != 'recruitment':
            print(f"  📊 Processing {current_year} data...")

            # Calculate y values based on chart type
            if chart_type == 'percentage' and current_regulars:
                y_values = (current_df[col_name] / current_regulars) * 100
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>📊 {val:.1f}%"
                             for date, val in zip(current_df['normalized_date'], y_values)]
            else:
                y_values = current_df[col_name]
                hover_text = [f"📅 {date.strftime('%d %b %Y')}<br>👥 {val:.0f} unique servers"
                             for date, val in zip(current_df['normalized_date'], y_values)]

            fig.add_trace(
                go.Scatter(
                    x=current_df['normalized_date'],
                    y=y_values,
                    mode='lines+markers',
                    name=f'{current_year}' if chart_index == 1 else None,
                    line=dict(color=colors['current'], width=3),
                    marker=dict(size=4, symbol='diamond'),
                    hovertemplate='%{text}<extra></extra>',
                    text=hover_text,
                    showlegend=(chart_index == 1),
                    legendgroup='current'
                ),
                row=row, col=col_pos
            )

        # Special handling for recruitment chart - use bar chart instead of line chart
        if chart_type == 'recruitment':
            print(f"  📊 Creating recruitment bar chart...")
            print(f"  🎯 Adding strategic plan target lines...")

            # Create bar chart data with string labels for x-axis
            years = [str(two_years_ago), str(last_year), str(current_year)]
            values = [two_years_recruitment, last_recruitment, current_recruitment]

            fig.add_trace(
                go.Bar(
                    x=years,
                    y=values,
                    marker=dict(color=['#059669', '#dc2626', '#1e40af']),
                    showlegend=False,
                    text=values,
                    textposition='outside'
                ),
                row=row, col=col_pos
            )
            print(f"    ✅ Bar chart trace added to subplot")

            # Add target lines from config
            try:
                from config import USING_GIFTS_TARGETS
                target_data = USING_GIFTS_TARGETS['word_based_ministry_recruitment']

                # Get baseline
                baseline = target_data['baseline']['value']
                baseline_year = target_data['baseline']['year']

                # Add baseline line
                fig.add_hline(
                    y=baseline,
                    line=dict(color='#6b7280', width=2, dash='dot'),
                    annotation_text=f"{baseline_year} Baseline: {baseline}",
                    annotation_position="bottom right",
                    annotation=dict(font=dict(size=10, color='#6b7280')),
                    row=row, col=col_pos
                )
                print(f"    ✅ Added baseline: {baseline}")

                # Add target lines
                target_colors = ['#10b981', '#14b8a6', '#06b6d4']  # Emerald, Teal, Cyan
                target_positions = ['top right', 'top left', 'top left']
                target_years = sorted(target_data['targets'].keys())

                for idx, target_year in enumerate(target_years[:3]):
                    target_count = target_data['targets'][target_year]
                    color = target_colors[idx]
                    position = target_positions[idx]

                    fig.add_hline(
                        y=target_count,
                        line=dict(color=color, width=2, dash='dash'),
                        annotation_text=f"{target_year} Target: {target_count}",
                        annotation_position=position,
                        annotation=dict(font=dict(size=10, color=color)),
                        row=row, col=col_pos
                    )
                    print(f"    ✅ Added {target_year} target: {target_count}")

            except (ImportError, KeyError) as e:
                print(f"    ⚠️ Could not load recruitment targets: {e}")

        # Add target lines for percentage chart (Chart 5)
        if chart_type == 'percentage':
            print(f"  🎯 Adding strategic plan target lines...")

            # Load targets from config
            try:
                from config import USING_GIFTS_TARGETS
                target_data = USING_GIFTS_TARGETS['pct_congregation_serving']

                # Get baseline and targets
                baseline_pct = target_data['baseline']['value'] * 100
                baseline_year = target_data['baseline']['year']

                # Add baseline line
                fig.add_hline(
                    y=baseline_pct,
                    line=dict(color='#6b7280', width=2, dash='dot'),
                    annotation_text=f"{baseline_year} Baseline: {baseline_pct:.0f}%",
                    annotation_position="bottom right",
                    annotation=dict(font=dict(size=10, color='#6b7280')),
                    row=row, col=col_pos
                )
                print(f"    ✅ Added baseline: {baseline_pct:.0f}%")

                # Define colors for the three target lines
                target_colors = ['#10b981', '#14b8a6', '#06b6d4']  # Emerald, Teal, Cyan
                target_positions = ['top right', 'top left', 'top left']

                # Get available target years
                target_years = sorted(target_data['targets'].keys())

                for idx, target_year in enumerate(target_years[:3]):
                    target_pct = target_data['targets'][target_year] * 100
                    color = target_colors[idx]
                    position = target_positions[idx]

                    fig.add_hline(
                        y=target_pct,
                        line=dict(color=color, width=2, dash='dash'),
                        annotation_text=f"{target_year} Target: {target_pct:.0f}%",
                        annotation_position=position,
                        annotation=dict(
                            font=dict(size=10, color=color)
                        ),
                        row=row, col=col_pos
                    )
                    print(f"    ✅ Added {target_year} target: {target_pct:.0f}%")

            except (ImportError, KeyError) as e:
                print(f"    ⚠️ Could not load strategic plan targets: {e}")

        # Add target lines for word-based ministry chart
        if chart_type == 'word_based':
            print(f"  🎯 Adding word-based ministry target lines...")

            try:
                from config import USING_GIFTS_TARGETS
                target_data = USING_GIFTS_TARGETS['number_in_word_based_ministry']

                # Get baseline and targets
                baseline_count = target_data['baseline']['value']

                # Add baseline
                fig.add_hline(
                    y=baseline_count,
                    line=dict(color='#6b7280', width=2, dash='dot'),
                    annotation_text=f"Baseline: {baseline_count}",
                    annotation_position="bottom right",
                    annotation=dict(font=dict(size=10, color='#6b7280')),
                    row=row, col=col_pos
                )
                print(f"    ✅ Added baseline: {baseline_count}")

                # Add target lines
                target_colors = ['#10b981', '#14b8a6', '#06b6d4']
                target_positions = ['top right', 'top left', 'top left']
                target_years = sorted(target_data['targets'].keys())

                for idx, target_year in enumerate(target_years[:3]):
                    target_count = target_data['targets'][target_year]
                    color = target_colors[idx]
                    position = target_positions[idx]

                    fig.add_hline(
                        y=target_count,
                        line=dict(color=color, width=2, dash='dash'),
                        annotation_text=f"{target_year} Target: {target_count}",
                        annotation_position=position,
                        annotation=dict(font=dict(size=10, color=color)),
                        row=row, col=col_pos
                    )
                    print(f"    ✅ Added {target_year} target: {target_count}")

            except (ImportError, KeyError) as e:
                print(f"    ⚠️ Could not load word-based targets: {e}")

    # Set consistent x-axis range for all charts
    x_min = datetime(base_year, 1, 1)
    x_max = datetime(base_year, 12, 31)

    # Update layout for each subplot
    for i, chart_config in enumerate(charts_config):
        row = chart_config['row']
        col_pos = chart_config['col_pos']
        chart_type = chart_config.get('chart_type', 'count')

        # Skip date range for recruitment chart (it's a bar chart with year labels)
        if chart_type != 'recruitment':
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
        else:
            # For recruitment chart, use categorical x-axis with year labels
            fig.update_xaxes(
                type='category',
                categoryorder='array',
                categoryarray=[str(two_years_ago), str(last_year), str(current_year)],
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(0,0,0,0.1)',
                tickfont=dict(size=10),
                row=row, col=col_pos
            )
        fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(0,0,0,0.1)',
            tickfont=dict(size=10),
            row=row, col=col_pos
        )

    # Build title text with strategic plan context
    title_main = "<b style='font-size:24px'>St George's Magill - Using Gifts Report</b>"
    title_sub = f"<br><span style='font-size:14px; color:#64748b'>Three-Year Cumulative Serving Trends</span>"

    # Add strategic plan context if available
    if pct_serving_target and current_year in pct_serving_target['targets']:
        target_pct = pct_serving_target['targets'][current_year] * 100
        baseline_pct = pct_serving_target['baseline']['value'] * 100
        title_note = f"<br><span style='font-size:11px; color:#94a3b8'>Generated {datetime.now().strftime('%B %d, %Y')} • Shows cumulative unique servers throughout the year • Strategic Plan Target: {target_pct:.0f}% of congregation serving (from {baseline_pct:.0f}% baseline)</span>"
    else:
        title_note = f"<br><span style='font-size:11px; color:#94a3b8'>Generated {datetime.now().strftime('%B %d, %Y')} • Shows cumulative unique servers throughout the year</span>"

    full_title = title_main + title_sub + title_note

    # Build annotation texts for all 6 charts
    stats_text_0 = f"<b>{all_stats[0]['title'].split(' - ')[0]}</b><br>{' | '.join(all_stats[0]['stats'])}" if all_stats[0]['stats'] else f"<b>{all_stats[0]['title'].split(' - ')[0]}</b>"
    stats_text_1 = f"<b>{all_stats[1]['title'].split(' - ')[0]}</b><br>{' | '.join(all_stats[1]['stats'])}" if all_stats[1]['stats'] else f"<b>{all_stats[1]['title'].split(' - ')[0]}</b>"
    stats_text_2 = f"<b>{all_stats[2]['title'].split(' - ')[0]}</b><br>{' | '.join(all_stats[2]['stats'])}" if all_stats[2]['stats'] else f"<b>{all_stats[2]['title'].split(' - ')[0]}</b>"
    stats_text_3 = f"<b>{all_stats[3]['title'].split(' - ')[0]}</b><br>{' | '.join(all_stats[3]['stats'])}" if all_stats[3]['stats'] else f"<b>{all_stats[3]['title'].split(' - ')[0]}</b>"
    stats_text_4 = f"<b>{all_stats[4]['title'].split(' - ')[0]}</b><br>{' | '.join(all_stats[4]['stats'])}" if all_stats[4]['stats'] else f"<b>{all_stats[4]['title'].split(' - ')[0]}</b>"
    # Chart 5 is recruitment chart, use full title
    stats_text_5 = f"<b>{all_stats[5]['title']}</b><br>{' | '.join(all_stats[5]['stats'])}" if all_stats[5]['stats'] else f"<b>{all_stats[5]['title']}</b>"

    # Overall layout (increased height for 3 rows)
    fig.update_layout(
        title=dict(
            text=full_title,
            x=0.5,
            y=0.99,
            font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=20, color='#1e293b')
        ),
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=12),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=1800,  # Increased from 1200 for 3 rows
        width=1600,
        hovermode='closest',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.05,
            xanchor="center",
            x=0.5,
            font=dict(family="Inter, system-ui, sans-serif", size=12, color='#374151'),
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor="rgba(0,0,0,0.1)",
            borderwidth=1
        ),
        margin=dict(l=80, r=80, t=120, b=100),
        annotations=[
            # Dummy annotations for plotly (needed for proper spacing - one per actual annotation)
            dict(
                text='',
                x=0.5, y=0.5,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            dict(
                text='',
                x=0.1, y=0.9,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            dict(
                text='',
                x=0.9, y=0.9,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            # Row 1 dummies
            dict(
                text='',
                x=0.22, y=1.02,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            dict(
                text='',
                x=0.78, y=1.02,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            # Row 2 dummies
            dict(
                text='',
                x=0.22, y=0.66,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            dict(
                text='',
                x=0.78, y=0.66,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            # Row 3 dummies
            dict(
                text='',
                x=0.22, y=0.30,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            dict(
                text='',
                x=0.78, y=0.30,
                xref='paper', yref='paper',
                showarrow=False, visible=False
            ),
            # Row 1
            dict(
                text=stats_text_0,
                x=0.44, y=1.02,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_1,
                x=0.90, y=1.02,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            # Row 2
            dict(
                text=stats_text_2,
                x=0.44, y=0.66,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_3,
                x=0.90, y=0.66,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            # Row 3
            dict(
                text=stats_text_4,
                x=0.44, y=0.30,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
                bgcolor="rgba(248,250,252,0.9)",
                bordercolor="rgba(0,0,0,0.1)",
                borderwidth=1, borderpad=5,
                xanchor='right', yanchor='top'
            ),
            dict(
                text=stats_text_5,
                x=0.90, y=0.30,
                xref='paper', yref='paper',
                showarrow=False, align='right',
                font=dict(family="Inter, system-ui, sans-serif", size=11, color='#1e293b'),
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

    # Use fixed filenames
    html_filename = os.path.join(outputs_dir, "gifts_dashboard.html")
    png_filename = os.path.join(outputs_dir, "gifts_dashboard.png")

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
        fig.write_image(png_filename, width=1600, height=1800, scale=3)
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

def main():
    """Main execution function"""
    try:
        print("🚀 Starting serving participation analysis...")
        print("📊 Will generate charts with 3 years of data")
        print("📁 Output: HTML (interactive) and PNG files saved to /outputs directory")
        print("🔄 Old dashboard files will be replaced with new versions")

        # Find all three serving report groups
        current_group, last_year_group, two_years_ago_group = find_serving_report_groups()

        if not current_group:
            print("❌ Cannot proceed without current year serving report group")
            return

        if not last_year_group:
            print("❌ Last year serving report group required but not found")
            return

        if not two_years_ago_group:
            print("❌ Two years ago serving report group required but not found")
            return

        # Extract serving data for 3 years
        current_df, current_headers, current_year = extract_serving_data_from_group(current_group, "Current Year")
        last_df, last_headers, last_year = extract_serving_data_from_group(last_year_group, "Last Year")
        two_years_df, two_years_headers, two_years_year = extract_serving_data_from_group(two_years_ago_group, "Two Years Ago")

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
            print("❌ Cannot proceed without current year service data")
            return

        # Calculate cumulative servers
        current_cumulative = calculate_cumulative_servers(current_df, current_services, "Current Year")
        last_cumulative = calculate_cumulative_servers(last_df, last_services, "Last Year")
        two_years_cumulative = calculate_cumulative_servers(two_years_df, two_years_services, "Two Years Ago")

        # Get regulars count at start of each year for percentage calculation
        # (Like applying_word.py - uses Congregation_ + RosteredMember_ at Jan 1)
        print("\n📊 Calculating regulars at year start for percentage denominators...")
        print("   (Using same method as applying_word.py - Congregation_ + RosteredMember_)")

        current_regulars = get_regulars_at_year_start(current_year)
        last_regulars = get_regulars_at_year_start(last_year)
        two_years_regulars = get_regulars_at_year_start(two_years_year)

        # Get word-based ministry recruitment numbers for all three years
        print("\n📊 Calculating word-based ministry recruitment for all years...")
        print("   (Counting new RosteredMember_ who serve in word-based positions)")

        # Get all people for matching
        people_resp = make_request('people/getAll', {'page_size': 1000})
        all_people = []
        if people_resp and people_resp.get('people'):
            people = people_resp['people'].get('person', [])
            if not isinstance(people, list):
                people = [people] if people else []
            all_people = people

        # Get all groups for group leader checking
        groups_resp = make_request('groups/getAll', {'page_size': 1000})
        all_groups = []
        if groups_resp and groups_resp.get('groups'):
            groups = groups_resp['groups'].get('group', [])
            if not isinstance(groups, list):
                groups = [groups] if groups else []
            all_groups = groups

        # Get recruitment reports and analyze
        current_report, last_report, two_years_report = find_new_serving_members_reports()
        current_recruitment = 0
        last_recruitment = 0
        two_years_recruitment = 0

        # Analyze current year recruitment
        if current_report:
            current_members = extract_new_serving_members_data(current_report)
            if current_members:
                current_recruitment, _ = analyze_new_serving_members_word_based(current_members, all_people, all_groups)

        # Analyze last year recruitment
        if last_report:
            last_members = extract_new_serving_members_data(last_report)
            if last_members:
                last_recruitment, _ = analyze_new_serving_members_word_based(last_members, all_people, all_groups)

        # Analyze two years ago recruitment
        if two_years_report:
            two_years_members = extract_new_serving_members_data(two_years_report)
            if two_years_members:
                two_years_recruitment, _ = analyze_new_serving_members_word_based(two_years_members, all_people, all_groups)

        print(f"\n✅ Final recruitment counts:")
        print(f"   {current_year}: {current_recruitment} recruits")
        print(f"   {last_year}: {last_recruitment} recruits")
        print(f"   {two_years_year}: {two_years_recruitment} recruits")

        # Create combined dashboard
        html_file = create_serving_dashboard(
            current_cumulative, last_cumulative, two_years_cumulative,
            current_regulars, last_regulars, two_years_regulars,
            current_recruitment, last_recruitment, two_years_recruitment
        )

        if not html_file:
            print("❌ Failed to create dashboard")
            return

        # Open the HTML file automatically
        try:
            import webbrowser
            import os
            abs_path = os.path.abspath(html_file)
            webbrowser.open(f"file://{abs_path}")
            print(f"\n🌐 Interactive dashboard opened automatically!")
        except Exception as e:
            print(f"\n⚠️ Couldn't auto-open dashboard: {e}")

        print(f"\n🎉 USING GIFTS DASHBOARD CREATED!")
        print(f"📊 THREE-YEAR COMPARISON: Current, Last Year, and Two Years Ago")
        print(f"📄 HTML file (interactive): {html_file}")
        print(f"🖼️ PNG file (static): {html_file.replace('.html', '.png')}")
        print(f"\n✨ FEATURES:")
        print(f"   • Six charts showing serving participation trends:")
        print(f"     - 8:30 AM, 10:30 AM congregation cumulative servers")
        print(f"     - Combined congregation cumulative servers")
        print(f"     - Combined congregation % serving (with strategic plan targets)")
        print(f"     - Word-based ministry cumulative servers and recruitment")
        print(f"   • Year-over-year comparison with color-coded lines (Blue, Red, Green)")
        print(f"   • Strategic plan target lines on percentage chart")
        print(f"   • Statistics embedded in chart annotations")
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
