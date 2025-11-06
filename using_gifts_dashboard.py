import requests
import json
from datetime import datetime, timedelta
import webbrowser
import os
import time
from collections import defaultdict
import pandas as pd
from bs4 import BeautifulSoup
import re
import subprocess
import sys
import math

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'pandas', 'html2image']
    for package in packages:
        try:
            if package == 'beautifulsoup4':
                import bs4
            elif package == 'html2image':
                import html2image
            else:
                __import__(package)
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

# Install packages first
install_packages()

# Get API key
print("🎯 USING GIFTS MINISTRY DASHBOARD")
print("="*35)
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

# Word-based ministry volunteer positions
WORD_BASED_POSITIONS = {
    'Acoustic Guitar', 'Bible Reader', 'BSG Leader', 'Band Leader', 'Bass', 'Cajon', 'Comm. Celebrant',
    'Cornet', 'Drums', 'Electric Guitar', 'Flute', 'Gospel Story Teller', 'Jnr Kids Assist',
    'Jnr Kids Leader', 'Keyboard', 'Kids Church Assist', 'Kids Church Leader', 
    'Kids Club Assistant', 'Kids Talk', 'Oboe', 'Organ', 'Prayer Leader', 'Preacher',
    'Service Leader', 'Vocals', 'Youth Assist', 'Youth Group Talk', 'Youth Leader'
}

def make_request(endpoint, params=None):
    try:
        response = requests.post(f"{BASE_URL}/{endpoint}.json", auth=(API_KEY, ''), json=params or {}, timeout=30)
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

def is_person_excluded(person):
    """
    Check if a person should be excluded from dashboard calculations.
    Excludes deceased people and people with '*' categories.
    """
    # Check if person is deceased
    if person.get('deceased') == 1 or person.get('deceased') == '1':
        return True, "deceased"

    # Check if person is archived (optional - you may want to include or exclude these)
    if person.get('archived') == 1 or person.get('archived') == '1':
        return True, "archived"

    # Check status field as backup
    status = person.get('status', '').lower()
    if status == 'deceased':
        return True, "status_deceased"

    return False, None

def fetch_all_people():
    print("📋 Fetching people with demographics and departments...")
    all_people = []
    excluded_count = 0
    deceased_count = 0
    archived_count = 0

    page = 1
    while True:
        print(f"Page {page}...", end=" ")
        response = make_request('people/getAll', {
            'page': page, 
            'page_size': 1000,
            'fields': ['demographics', 'departments', 'locations']  # Added locations field for congregation assignment
        })
        if not response or not response.get('people'):
            print("Done")
            break
        people = response['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []

        # Filter out deceased and other excluded people (but don't show who is excluded)
        filtered_people = []
        for person in people:
            is_excluded, reason = is_person_excluded(person)
            if is_excluded:
                excluded_count += 1
                if reason == "deceased" or reason == "status_deceased":
                    deceased_count += 1
                elif reason == "archived":
                    archived_count += 1
                # REMOVED: Don't log excluded people individually
            else:
                filtered_people.append(person)

        all_people.extend(filtered_people)
        print(f" ({len(filtered_people)} active, {len(people) - len(filtered_people)} excluded)")

        if len(people) < 1000:
            break
        page += 1

    print(f"\nFinal totals:")
    print(f"  Active people: {len(all_people)}")
    print(f"  Deceased excluded: {deceased_count}")
    print(f"  Archived excluded: {archived_count}")
    print(f"  Total excluded: {excluded_count}")

    return all_people

def fetch_categories():
    print("📋 Fetching categories...")
    response = make_request('people/categories/getAll')
    if response and response.get('categories'):
        categories = response['categories'].get('category', [])
        if not isinstance(categories, list):
            categories = [categories] if categories else []
        return categories
    return []

def fetch_groups():
    print("📋 Fetching groups...")
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
        if len(groups) < 1000:
            break
        page += 1
    print(f"Total groups: {len(all_groups)}")
    return all_groups

def find_attendance_report_groups():
    """Find both current and last year attendance report groups"""
    print("\n📋 Searching for attendance report groups...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None, None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_group = None
    last_year_group = None

    for group in groups:
        group_name = group.get('name', '').lower()
        if 'report of last year service individual attendance' in group_name:
            last_year_group = group
            print(f"✅ Found last year group: {group.get('name')}")
        elif 'report of service individual attendance' in group_name and 'last year' not in group_name:
            current_group = group
            print(f"✅ Found current year group: {group.get('name')}")

    if not current_group:
        print("❌ Current year attendance report group not found")
    if not last_year_group:
        print("❌ Last year attendance report group not found")

    return current_group, last_year_group

def find_new_serving_members_report():
    """Find the 'Report of New Serving Members' group - FROM CODE WBM"""
    print("\n📋 Searching for 'Report of New Serving Members' group...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None
    
    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []
    
    for group in groups:
        group_name = group.get('name', '').lower()
        if 'report of new serving members' in group_name:
            print(f"✅ Found report group: {group.get('name')}")
            return group
    
    print("❌ 'Report of New Serving Members' group not found")
    return None

def extract_new_serving_members_data(group):
    """Extract new serving members data from the report group - FROM CODE WBM"""
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
                    print(f"📋 Headers found: {headers}")
                    
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
                                member_id = row_data.get('Member ID', '')  # Extract Member ID
                                if person_name:
                                    new_serving_members.append({
                                        'name': person_name,
                                        'member_id': member_id,  # Store Member ID
                                        'date_changed': row_data.get('Date', ''),
                                        'change_from': row_data.get('Change From', ''),
                                        'date_added': row_data.get('Date Added', ''),
                                        'preferred_name': row_data.get('Preferred Name', '')
                                    })
                    
                    print(f"✅ Extracted {len(new_serving_members)} new serving members")
                    
                    # Debug: Show what Member IDs we found
                    member_ids_found = [m['member_id'] for m in new_serving_members if m['member_id']]
                    print(f"📊 Member IDs found: {member_ids_found}")
                    
                    return new_serving_members
        
        print("❌ New serving members table not found in report")
        return []
        
    except Exception as e:
        print(f"❌ Error extracting new serving members data: {e}")
        return []

def get_detailed_person_info(person_id):
    """Get detailed person information including full volunteer positions - FROM CODE WBM"""
    try:
        response = make_request('people/getInfo', {
            'id': person_id,
            'fields': ['demographics', 'departments', 'volunteer', 'locations', 'family']
        })
        if response and response.get('person'):
            return response['person']
    except Exception as e:
        print(f"    ⚠️ Error getting detailed info for person {person_id}: {e}")
    return None

def parse_column_header(header):
    """Parse column headers like '8:30 AM' or '10:30 AM Communion 02/06/2024'"""

    # Extract time (look for patterns like 8:30 AM, 10:30 AM, etc.)
    time_match = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', header, re.IGNORECASE)
    if not time_match:
        return None

    time_str = f"{time_match.group(1)} {time_match.group(2).upper()}"

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
        'original_header': header
    }

def extract_attendance_data_from_group(group, year_label):
    """Extract attendance data from an attendance report group"""
    print(f"\n🔍 Extracting {year_label} attendance data...")

    if not group:
        print(f"❌ No {year_label} group provided")
        return None, None

    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print(f"❌ No URL found in {year_label} group")
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
                    return pd.DataFrame(attendance_records), headers

        print(f"❌ {year_label} attendance table not found")
        return None, None

    except Exception as e:
        print(f"❌ Error extracting {year_label} data: {e}")
        return None, None

def analyze_congregation_membership_from_df(attendance_df, headers, year_label, member_names):
    """Analyze which congregation each MEMBER primarily attends from a specific dataframe"""
    print(f"\n📊 Analyzing {year_label} congregation membership for members only...")

    if attendance_df is None or len(attendance_df) == 0:
        print(f"❌ No {year_label} attendance data available")
        return {}

    # Parse service columns
    service_columns = {}
    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue

        parsed = parse_column_header(header)
        if parsed and parsed['normalized_time'] in ['8:30', '10:30', '6:30']:
            if parsed['normalized_time'] not in service_columns:
                service_columns[parsed['normalized_time']] = []
            service_columns[parsed['normalized_time']].append(header)

    print(f"  Found {year_label} service columns:")
    for time, columns in service_columns.items():
        print(f"    {time}: {len(columns)} services")

    # For each MEMBER, count attendance at each service time
    congregation_assignments = {}

    for _, row in attendance_df.iterrows():
        first_name = row.get('First Name', '')
        last_name = row.get('Last Name', '')
        full_name = f"{first_name} {last_name}".strip()

        # FILTER: Only process if this person is a member
        if not full_name or full_name not in member_names:
            continue

        # Count attendance by service time
        attendance_counts = {'8:30': 0, '10:30': 0, '6:30': 0}

        for service_time, columns in service_columns.items():
            for column in columns:
                if column in row and row[column] == 'Y':
                    attendance_counts[service_time] += 1

        # Assign to congregation with most attendance
        total_attendance = sum(attendance_counts.values())
        if total_attendance > 0:
            primary_congregation = max(attendance_counts, key=attendance_counts.get)
            congregation_assignments[full_name] = primary_congregation

    print(f"  ✅ Assigned {len(congregation_assignments)} members from {year_label} data")
    for cong in ['8:30', '10:30', '6:30']:
        count = sum(1 for c in congregation_assignments.values() if c == cong)
        print(f"    {cong}: {count} people")

    return congregation_assignments

def analyze_congregation_from_location(people):
    """Analyze congregation preference from location field"""
    print(f"\n📍 Analyzing congregation from location data...")
    print(f"   Checking {len(people)} members for location preferences...")

    location_assignments = {}
    location_found = 0
    locations_checked = 0

    for person in people:
        # Check multiple possible field names for location
        location = (person.get('location', '') or 
                   person.get('locations', '') or
                   person.get('Location', '') or
                   person.get('Locations', ''))

        full_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()

        if not full_name:
            continue

        locations_checked += 1

        # Debug: Show first few location values
        if locations_checked <= 5:
            print(f"    Debug - {full_name}: location field = '{location}'")

        if not location:
            continue

        location_str = str(location).strip()
        location_lower = location_str.lower()

        # Check for congregation indicators in location (based on screenshot format)
        congregation = None
        if "st george's 8:30 am" in location_lower or "8:30 am" in location_lower:
            congregation = '8:30'
        elif "st george's 10:30 am" in location_lower or "10:30 am" in location_lower:
            congregation = '10:30'
        elif "st george's 6:30 pm" in location_lower or "6:30 pm" in location_lower:
            congregation = '6:30'
        elif '8:30' in location_lower:
            congregation = '8:30'
        elif '10:30' in location_lower:
            congregation = '10:30'
        elif '6:30' in location_lower:
            congregation = '6:30'
        elif '8.30' in location_lower:
            congregation = '8:30'
        elif '10.30' in location_lower:
            congregation = '10:30'
        elif '6.30' in location_lower:
            congregation = '6:30'

        if congregation:
            location_assignments[full_name] = congregation
            location_found += 1
            print(f"    📍 {full_name}: {congregation} (from location: '{location_str}')")

    print(f"  ✅ Checked {locations_checked} members, found {location_found} location-based assignments")
    for cong in ['8:30', '10:30', '6:30']:
        count = sum(1 for c in location_assignments.values() if c == cong)
        print(f"    {cong}: {count} people")

    return location_assignments

def create_two_tier_congregation_assignments(current_assignments, last_year_assignments):
    """Combine current and last year assignments with current year taking priority"""
    print(f"\n🔄 Creating two-tier congregation assignments...")

    # Priority order: current year > last year
    combined_assignments = {}

    # Tier 1: Current year assignments (highest priority)
    combined_assignments.update(current_assignments)
    tier1_count = len(current_assignments)

    # Tier 2: Last year assignments for people not found in current year
    tier2_added = 0
    for name, congregation in last_year_assignments.items():
        if name not in combined_assignments:
            combined_assignments[name] = congregation
            tier2_added += 1

    print(f"  ✅ Two-tier assignment results:")
    print(f"    Tier 1 (Current year): {tier1_count}")
    print(f"    Tier 2 (Last year): {tier2_added}")
    print(f"    Total assigned: {len(combined_assignments)}")

    # Show final distribution
    final_distribution = {}
    for cong in ['8:30', '10:30', '6:30']:
        count = sum(1 for c in combined_assignments.values() if c == cong)
        final_distribution[cong] = count
        print(f"    {cong}: {count} people")

    return combined_assignments

def categorize_person(category):
    if not category:
        return 'people'
    if category.endswith('*'):
        return 'excluded'
    elif category == 'RosteredMember_':
        return 'serving_member'
    elif category == 'Congregation_':
        return 'congregation_only'
    elif category.endswith('_'):
        return 'other_member'
    else:
        return 'people'

def extract_volunteer_positions(person):
    """Extract volunteer positions from person's departments"""
    positions = set()

    departments = person.get('departments', [])
    if not departments:
        return positions

    # Handle both single department and list of departments
    if isinstance(departments, dict) and 'department' in departments:
        dept_list = departments['department']
    else:
        dept_list = departments if isinstance(departments, list) else [departments]

    if not isinstance(dept_list, list):
        dept_list = [dept_list] if dept_list else []

    for dept in dept_list:
        if not isinstance(dept, dict):
            continue

        # Look for sub_departments
        sub_depts = dept.get('sub_departments', {})
        if isinstance(sub_depts, dict) and 'sub_department' in sub_depts:
            sub_dept_list = sub_depts['sub_department']
            if not isinstance(sub_dept_list, list):
                sub_dept_list = [sub_dept_list] if sub_dept_list else []

            for sub_dept in sub_dept_list:
                if not isinstance(sub_dept, dict):
                    continue

                # Look for positions
                position_data = sub_dept.get('positions', {})
                if isinstance(position_data, dict) and 'position' in position_data:
                    pos_list = position_data['position']
                    if not isinstance(pos_list, list):
                        pos_list = [pos_list] if pos_list else []

                    for pos in pos_list:
                        if isinstance(pos, dict):
                            pos_name = pos.get('name', '')
                            if pos_name:
                                positions.add(pos_name)

    return positions

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

def is_word_based_group(group_name):
    """Check if group qualifies for word-based ministry"""
    if not group_name:
        return False
    group_name_lower = group_name.lower()
    return not ('cherry picking' in group_name_lower or 'community care' in group_name_lower)

def is_group_leader(person_id, groups):
    """Check if person is a leader of any word-based ministry group"""
    for group in groups:
        group_name = group.get('name', '')

        if not is_word_based_group(group_name):
            continue

        if group.get('people') and group['people'].get('person'):
            group_members = group['people']['person']
            if not isinstance(group_members, list):
                group_members = [group_members] if group_members else []

            for member in group_members:
                if member.get('id') == person_id:
                    role = member.get('role', '').lower()
                    if 'leader' in role or 'assistant' in role:
                        return True
    return False

def calculate_service_load_concentration(people_in_congregation):
    """Calculate how concentrated the service load is (Pareto principle)"""
    serving_people = [p for p in people_in_congregation if p.get('serves')]

    if not serving_people:
        return 0

    # Count roles per person
    role_counts = []
    for person in serving_people:
        role_counts.append(len(person.get('positions', [])))

    if not role_counts or sum(role_counts) == 0:
        return 0

    # Sort by number of roles (descending)
    role_counts.sort(reverse=True)

    # Calculate what percentage of people handle 80% of roles
    total_roles = sum(role_counts)
    cumulative_roles = 0
    people_count = 0
    target_roles = total_roles * 0.8

    for roles in role_counts:
        cumulative_roles += roles
        people_count += 1
        if cumulative_roles >= target_roles:
            break

    concentration_percentage = (people_count / len(serving_people)) * 100
    return concentration_percentage

def calculate_visitor_to_serving_member_conversion(people, categories):
    """Calculate Visitor to Serving Member Conversion Rate"""
    print(f"\n📊 Calculating Visitor to Serving Member Conversion Rate...")

    # Create category lookup
    category_lookup = {}
    for cat in categories:
        category_lookup[cat.get('id', '')] = cat.get('name', '')

    # Current year boundaries
    current_year = datetime.now().year
    current_year_start = datetime(current_year, 1, 1)

    # Find people added this calendar year
    people_added_this_year = []

    for person in people:
        try:
            date_added = datetime.strptime(person.get('date_added', ''), '%Y-%m-%d %H:%M:%S')
        except:
            continue

        if date_added >= current_year_start:
            people_added_this_year.append(person)

    # Of those added this year, find who are now RosteredMember_
    rosteredmembers_added_this_year = []

    for person in people_added_this_year:
        category_id = person.get('category_id', '')
        category_name = category_lookup.get(category_id, '')

        if category_name == 'RosteredMember_':
            rosteredmembers_added_this_year.append(person)

    # Calculate conversion rate
    total_added = len(people_added_this_year)
    rosteredmembers_count = len(rosteredmembers_added_this_year)
    conversion_rate = (rosteredmembers_count / total_added * 100) if total_added > 0 else 0

    print(f"  ✅ Visitor to Serving Member Conversion Analysis:")
    print(f"    People added in {current_year}: {total_added}")
    print(f"    Now RosteredMember_: {rosteredmembers_count}")
    print(f"    Conversion rate: {conversion_rate:.1f}%")

    # List the people who converted for verification
    if rosteredmembers_added_this_year:
        print(f"  📋 People added in {current_year} who are now RosteredMember_:")
        for person in rosteredmembers_added_this_year:
            name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            date_added = person.get('date_added', '')
            print(f"    • {name} (added: {date_added})")

    return {
        'total_added_this_year': total_added,
        'rosteredmembers_this_year': rosteredmembers_count,
        'conversion_rate': conversion_rate
    }

def analyze_new_serving_members_word_based(new_serving_members, all_people, groups):
    """Analyze which new serving members are in word-based ministry using Member ID matching - FROM CODE WBM"""
    print(f"\n📖 Analyzing {len(new_serving_members)} new serving members for word-based ministry...")
    
    # Create lookup of all people by Member ID (much more reliable than name matching!)
    people_by_id = {}
    for person in all_people:
        person_id = person.get('id')
        if person_id:
            people_by_id[person_id] = person
    
    print(f"📊 Debug: Created ID lookup for {len(people_by_id)} people")
    
    new_word_based_count = 0
    new_word_based_details = []
    not_found_count = 0
    
    print(f"\n📋 Detailed analysis of new serving members:")
    
    for new_member in new_serving_members:
        member_name = new_member['name'].strip()
        member_id = new_member.get('member_id', '').strip()
        date_changed = new_member['date_changed']
        
        print(f"\n  🔍 Analyzing: {member_name} (ID: {member_id}) (became serving member: {date_changed})")
        
        # Use Member ID for direct lookup (much more reliable!)
        person_record = None
        if member_id and member_id in people_by_id:
            person_record = people_by_id[member_id]
            print(f"    ✅ Found person using Member ID: {member_id}")
        elif member_id:
            print(f"    ⚠️ Member ID '{member_id}' not found in API data")
            not_found_count += 1
            continue
        else:
            print(f"    ⚠️ No Member ID available in report for {member_name}")
            not_found_count += 1
            continue
        
        person_id = person_record.get('id')
        volunteer_flag = person_record.get('volunteer', '0')
        is_word_based = False
        reasons = []
        
        print(f"    📊 API Person ID: {person_id}")
        print(f"    📊 Volunteer flag: {volunteer_flag}")
        
        # Try to get more detailed person information if needed
        detailed_person = get_detailed_person_info(person_id)
        if detailed_person:
            person_to_analyze = detailed_person
            volunteer_flag = person_to_analyze.get('volunteer', volunteer_flag)
            print(f"    📊 Got detailed person info from API")
        else:
            person_to_analyze = person_record
        
        # Check volunteer positions
        if str(volunteer_flag) == '1':
            positions = extract_volunteer_positions(person_to_analyze)
            print(f"    📊 Extracted positions: {positions}")
            word_based_positions = positions.intersection(WORD_BASED_POSITIONS)
            print(f"    📊 Word-based positions found: {word_based_positions}")
            
            if word_based_positions:
                is_word_based = True
                reasons.append(f"Volunteer positions: {', '.join(word_based_positions)}")
                print(f"    ✅ Volunteer positions: {', '.join(word_based_positions)}")
            else:
                print(f"    ℹ️ Volunteer with non-word-based positions: {', '.join(positions) if positions else 'No positions found'}")
                # Debug: Show what we're looking for
                if not positions:
                    departments = person_to_analyze.get('departments', [])
                    print(f"    📊 Raw departments structure: {type(departments)} - {str(departments)[:100]}...")
        else:
            print(f"    ℹ️ Not marked as volunteer (flag = {volunteer_flag})")
        
        # Check group leadership
        if is_group_leader(person_id, groups):
            is_word_based = True
            reasons.append("Group leader")
            print(f"    ✅ Is a group leader")
        else:
            print(f"    ℹ️ Not a group leader")
        
        if is_word_based:
            new_word_based_count += 1
            reason_str = " & ".join(reasons)
            detail = {
                'name': member_name,
                'member_id': member_id,
                'date_changed': date_changed,
                'reasons': reason_str,
                'change_from': new_member['change_from']
            }
            new_word_based_details.append(detail)
            print(f"    🎯 RESULT: IN WORD-BASED MINISTRY - {reason_str}")
        else:
            print(f"    📊 RESULT: Not in word-based ministry")
    
    print(f"\n✅ Analysis complete:")
    print(f"   Total new serving members: {len(new_serving_members)}")
    print(f"   Found in Elvanto API: {len(new_serving_members) - not_found_count}")
    print(f"   Not found in API: {not_found_count}")
    print(f"   In word-based ministry: {new_word_based_count}")
    
    if new_word_based_details:
        print(f"\n📖 New serving members in word-based ministry this calendar year:")
        for detail in new_word_based_details:
            print(f"   • {detail['name']} (ID: {detail['member_id']}): {detail['reasons']} (from {detail['change_from']} on {detail['date_changed']})")
    
    return new_word_based_count, new_word_based_details

def calculate_metrics(people, categories, groups, congregation_assignments, current_df, current_headers, member_names):
    print("🔄 Calculating Using Gifts metrics...")

    # Initialize word-based breakdown by congregation (must be available for metrics calculation)
    word_based_by_congregation = {'8:30': 0, '10:30': 0, '6:30': 0, 'unassigned': 0}

    # Create category lookup
    category_lookup = {}
    for cat in categories:
        category_lookup[cat.get('id', '')] = cat.get('name', '')

    # Process people and assign congregations (deceased already filtered out)
    processed_people = []
    unassigned_count = 0
    unassigned_people = []  # NEW: Track people not assigned to congregations

    # Initialize word-based breakdown by congregation (must be available for metrics calculation)
    word_based_by_congregation = {'8:30': 0, '10:30': 0, '6:30': 0, 'unassigned': 0}
    word_based_recruited_this_year = 0  # Initialize here too

    # Create category lookup
    category_lookup = {}
    for cat in categories:
        category_lookup[cat.get('id', '')] = cat.get('name', '')

    print("📊 Processing people and assigning congregations...")

    # MOVED: Get accurate word-based ministry recruitment FIRST (before metrics calculation)
    print(f"\n📋 Getting accurate word-based ministry recruitment from 'Report of New Serving Members'...")
    
    # Find and extract new serving members report
    new_serving_report = find_new_serving_members_report()
    
    if new_serving_report:
        new_serving_members = extract_new_serving_members_data(new_serving_report)
        if new_serving_members:
            word_based_recruited_this_year, word_based_details = analyze_new_serving_members_word_based(
                new_serving_members, people, groups
            )
            print(f"✅ Found {word_based_recruited_this_year} new serving members in word-based ministry this calendar year")
            
            # Break down by congregation using LOCATION field
            print(f"📊 Breaking down word-based recruits by congregation using location data...")
            
            for detail in word_based_details:
                person_name = detail['name']
                member_id = detail['member_id']
                
                # Find the person in our people data to get their location
                person_congregation = 'unassigned'
                
                for person in people:
                    if person.get('id') == member_id:
                        # Check location field for congregation assignment
                        location = (person.get('location', '') or 
                                   person.get('locations', '') or
                                   person.get('Location', '') or
                                   person.get('Locations', ''))
                        
                        if location:
                            location_str = str(location).strip().lower()
                            
                            # Parse location for congregation
                            if "st george's 10:30 am" in location_str or "10:30 am" in location_str:
                                person_congregation = '10:30'
                            elif "st george's 8:30 am" in location_str or "8:30 am" in location_str:
                                person_congregation = '8:30'
                            elif "st george's 6:30 pm" in location_str or "6:30 pm" in location_str:
                                person_congregation = '6:30'
                            elif '10:30' in location_str:
                                person_congregation = '10:30'
                            elif '8:30' in location_str:
                                person_congregation = '8:30'
                            elif '6:30' in location_str:
                                person_congregation = '6:30'
                            elif '10.30' in location_str:
                                person_congregation = '10:30'
                            elif '8.30' in location_str:
                                person_congregation = '8:30'
                            elif '6.30' in location_str:
                                person_congregation = '6:30'
                        break
                
                # Count toward the appropriate congregation
                if person_congregation in word_based_by_congregation:
                    word_based_by_congregation[person_congregation] += 1
                    print(f"  📖 {person_name} -> {person_congregation} congregation")
                else:
                    word_based_by_congregation['unassigned'] += 1
                    print(f"  📖 {person_name} -> unassigned")
            
            print(f"📊 Word-based recruits by congregation:")
            for cong, count in word_based_by_congregation.items():
                print(f"   {cong}: {count} people")
        else:
            print("⚠️ No new serving members found in report")
    else:
        print("⚠️ Could not find 'Report of New Serving Members' group")

    # NOW process people and assign congregations (WBM data is ready)
    print(f"\n📊 Now processing people for congregation assignment...")
    word_based_people = []

    for person in people:
        # Skip contacts (already filtered deceased/archived in fetch_all_people)
        contact = person.get('contact', 0)
        if contact == 1 or contact == '1':
            continue

        category_id = person.get('category_id', '')
        category_name = category_lookup.get(category_id, '')
        person_type = categorize_person(category_name)
        person_id = person.get('id')

        # Only process members (serving members + congregation)
        if person_type not in ['serving_member', 'congregation_only']:
            continue

        try:
            date_added = datetime.strptime(person.get('date_added', ''), '%Y-%m-%d %H:%M:%S')
        except:
            date_added = datetime.now()

        # Assign congregation based on attendance data
        full_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
        primary_congregation = congregation_assignments.get(full_name)

        if not primary_congregation:
            unassigned_count += 1
            # NEW: Track unassigned people details
            unassigned_people.append({
                'name': full_name,
                'category': category_name,
                'email': person.get('email', ''),
                'phone': person.get('phone', '')
            })

        # Extract volunteer information
        is_volunteer = str(person.get('volunteer', '0')) == '1'
        positions = extract_volunteer_positions(person) if is_volunteer else set()
        word_based_positions = positions.intersection(WORD_BASED_POSITIONS)
        is_group_leader_bool = is_group_leader(person_id, groups)

        is_word_based = bool(word_based_positions) or is_group_leader_bool
        is_serving = is_volunteer  # Anyone marked as volunteer is serving

        # NEW: Log everyone included in word-based ministry tally
        if is_word_based:
            reasons = []
            if word_based_positions:
                reasons.append(f"Volunteer positions: {', '.join(word_based_positions)}")
            if is_group_leader_bool:
                reasons.append("Group leader")

            reason_str = " & ".join(reasons)
            print(f"  📖 {full_name}: {reason_str}")
            word_based_people.append(full_name)

        processed_people.append({
            **person,
            'category_name': category_name,
            'type': person_type,
            'date_added': date_added,
            'primary_congregation': primary_congregation,
            'serves': is_serving,
            'word_based': is_word_based,
            'positions': positions,
            'word_based_positions': word_based_positions,
            'full_name': full_name
        })

    print(f"✅ Processed {len(processed_people)} members")
    print(f"⚠️ Unassigned to congregation: {unassigned_count} members")
    print(f"📖 Total in word-based ministry: {len(word_based_people)} people")

    # Group by congregation
    congregations = {
        '8:30': [p for p in processed_people if p.get('primary_congregation') == '8:30'],
        '10:30': [p for p in processed_people if p.get('primary_congregation') == '10:30'],
        '6:30': [p for p in processed_people if p.get('primary_congregation') == '6:30'],
        'overall': processed_people
    }

    print(f"📊 Final congregation distribution:")
    assigned_total = 0
    for cong_name, cong_people in congregations.items():
        if cong_name != 'overall':
            print(f"   {cong_name}: {len(cong_people)} members")
            assigned_total += len(cong_people)
    print(f"   Total assigned: {assigned_total}")
    print(f"   Total overall: {len(congregations['overall'])}")
    print(f"   Verification: {assigned_total}/{len(congregations['overall'])} = {(assigned_total/len(congregations['overall'])*100):.1f}% assigned")

    # Calculate metrics for each congregation
    metrics = {}

    for cong_name, cong_people in congregations.items():
        if not cong_people:
            metrics[cong_name] = {
                'total_members': 0,
                'serving_count': 0,
                'serve_percentage': 0,
                'word_based_count': 0,
                'load_concentration': 0,
                'avg_roles_per_person': 0,
                'new_to_serving': 0,
                'new_word_based': 0
            }
            continue

        total_members = len(cong_people)
        serving_people = [p for p in cong_people if p.get('serves')]
        serving_count = len(serving_people)
        serve_percentage = (serving_count / total_members * 100) if total_members > 0 else 0

        word_based_people = [p for p in cong_people if p.get('word_based')]
        word_based_count = len(word_based_people)

        # Calculate load metrics
        load_concentration = calculate_service_load_concentration(cong_people)

        total_roles = sum(len(p.get('positions', [])) for p in serving_people)
        avg_roles = (total_roles / serving_count) if serving_count > 0 else 0

        # Calculate new people (added in this calendar year)
        current_year_start = datetime(datetime.now().year, 1, 1)
        new_people = [p for p in cong_people if p['date_added'] >= current_year_start]
        new_to_serving = len([p for p in new_people if p.get('serves')])
        # Use accurate word-based count from WBM report instead of date-based calculation
        new_word_based = word_based_by_congregation.get(cong_name, 0)
        
        # DEBUG: Show what's happening with word-based assignment
        print(f"🔍 DEBUG METRICS: {cong_name} congregation:")
        print(f"    word_based_by_congregation lookup for '{cong_name}': {new_word_based}")
        print(f"    Full word_based_by_congregation: {word_based_by_congregation}")

        metrics[cong_name] = {
            'total_members': total_members,
            'serving_count': serving_count,
            'serve_percentage': serve_percentage,
            'word_based_count': word_based_count,
            'load_concentration': load_concentration,
            'avg_roles_per_person': avg_roles,
            'new_to_serving': new_to_serving,
            'new_word_based': new_word_based
        }

    # NEW: Get accurate word-based ministry recruitment using "Report of New Serving Members" approach
    print(f"\n📋 Getting accurate word-based ministry recruitment from 'Report of New Serving Members'...")
    
    # Find and extract new serving members report
    new_serving_report = find_new_serving_members_report()
    word_based_recruited_this_year = 0
    
    if new_serving_report:
        new_serving_members = extract_new_serving_members_data(new_serving_report)
        if new_serving_members:
            word_based_recruited_this_year, word_based_details = analyze_new_serving_members_word_based(
                new_serving_members, people, groups
            )
            print(f"✅ Found {word_based_recruited_this_year} new serving members in word-based ministry this calendar year")
            
            # Break down by congregation using LOCATION field (more reliable than attendance analysis)
            print(f"📊 Breaking down word-based recruits by congregation using location data...")
            print(f"🔍 DEBUG: Found {len(word_based_details)} word-based details to process")
            
            for detail in word_based_details:
                person_name = detail['name']
                member_id = detail['member_id']
                
                print(f"\n🔍 DEBUG: Processing {person_name} (Member ID: {member_id})")
                
                # Find the person in our people data to get their location
                person_congregation = 'unassigned'
                location = 'not found'
                person_found = False
                
                for person in people:
                    if person.get('id') == member_id:
                        person_found = True
                        print(f"✅ DEBUG: Found person in people data")
                        
                        # Check multiple possible location field names and show what we find
                        location_fields = ['location', 'locations', 'Location', 'Locations']
                        available_fields = list(person.keys())
                        print(f"🔍 DEBUG: Available person fields: {available_fields}")
                        
                        for field in location_fields:
                            field_value = person.get(field, '')
                            if field_value:
                                print(f"📍 DEBUG: {field} = '{field_value}'")
                            else:
                                print(f"📍 DEBUG: {field} = (empty or not found)")
                        
                        # Check location field for congregation assignment
                        location = (person.get('location', '') or 
                                   person.get('locations', '') or
                                   person.get('Location', '') or
                                   person.get('Locations', ''))
                        
                        print(f"📍 DEBUG: Final location value = '{location}'")
                        
                        if location:
                            location_str = str(location).strip().lower()
                            print(f"📍 DEBUG: Location string (lowercase) = '{location_str}'")
                            
                            # Parse location for congregation (same logic as analyze_congregation_from_location)
                            if "st george's 10:30 am" in location_str or "10:30 am" in location_str:
                                person_congregation = '10:30'
                                print(f"✅ DEBUG: Matched 10:30 AM congregation")
                            elif "st george's 8:30 am" in location_str or "8:30 am" in location_str:
                                person_congregation = '8:30'
                                print(f"✅ DEBUG: Matched 8:30 AM congregation")
                            elif "st george's 6:30 pm" in location_str or "6:30 pm" in location_str:
                                person_congregation = '6:30'
                                print(f"✅ DEBUG: Matched 6:30 PM congregation")
                            elif '10:30' in location_str:
                                person_congregation = '10:30'
                                print(f"✅ DEBUG: Matched 10:30 via simple text match")
                            elif '8:30' in location_str:
                                person_congregation = '8:30'
                                print(f"✅ DEBUG: Matched 8:30 via simple text match")
                            elif '6:30' in location_str:
                                person_congregation = '6:30'
                                print(f"✅ DEBUG: Matched 6:30 via simple text match")
                            elif '10.30' in location_str:
                                person_congregation = '10:30'
                                print(f"✅ DEBUG: Matched 10.30 via simple text match")
                            elif '8.30' in location_str:
                                person_congregation = '8:30'
                                print(f"✅ DEBUG: Matched 8.30 via simple text match")
                            elif '6.30' in location_str:
                                person_congregation = '6:30'
                                print(f"✅ DEBUG: Matched 6.30 via simple text match")
                            else:
                                print(f"❌ DEBUG: No congregation match found in location string")
                        else:
                            print(f"❌ DEBUG: No location data found")
                        break
                
                if not person_found:
                    print(f"❌ DEBUG: Person with Member ID {member_id} not found in people data")
                
                # Count toward the appropriate congregation
                if person_congregation in word_based_by_congregation:
                    word_based_by_congregation[person_congregation] += 1
                    print(f"✅ DEBUG: Added to {person_congregation} congregation count")
                    print(f"  📖 {person_name} -> {person_congregation} congregation (location: {location})")
                else:
                    word_based_by_congregation['unassigned'] += 1
                    print(f"⚠️ DEBUG: Added to unassigned count")
                    print(f"  📖 {person_name} -> unassigned (location: {location})")
            
            print(f"\n🔍 DEBUG: Final word_based_by_congregation counts:")
            for cong, count in word_based_by_congregation.items():
                print(f"   {cong}: {count} people")
        else:
            print("⚠️ No new serving members found in report")
    else:
        print("⚠️ Could not find 'Report of New Serving Members' group")

    # Calculate additional overall metrics
    all_serving = [p for p in processed_people if p.get('serves')]
    all_word_based = [p for p in processed_people if p.get('word_based')]

    # Calculate service load balance (overall concentration measure)
    service_load_balance = calculate_service_load_concentration(processed_people)

    # Strategic plan progress
    # Word-based ministry: Use the ACCURATE count from new serving members report
    word_based_baseline = 46
    word_based_target = math.ceil(word_based_baseline * 1.10)  # 46 * 1.10 = 50.6 → 51 people
    word_based_needed_this_year = word_based_target - word_based_baseline  # 51 - 46 = 5 people
    word_based_recruitment_progress = (word_based_recruited_this_year / word_based_needed_this_year * 100) if word_based_needed_this_year > 0 else 0

    # NEW: Visitor to Serving Member Conversion Rate
    conversion_data = calculate_visitor_to_serving_member_conversion(people, categories)

    metrics['strategic'] = {
        'word_based_target': word_based_target,
        'word_based_baseline': word_based_baseline,
        'word_based_needed_this_year': word_based_needed_this_year,
        'word_based_recruited_this_year': word_based_recruited_this_year,  # Now accurate!
        'word_based_recruitment_progress': word_based_recruitment_progress,
        'visitor_conversion_target': 5.0,  # 5% target from strategic plan
        'visitor_conversion_actual': conversion_data['conversion_rate'],
        'visitors_total': conversion_data['total_added_this_year'],
        'visitors_now_serving': conversion_data['rosteredmembers_this_year']
    }

    return metrics, processed_people, unassigned_people

def generate_dashboard_html(metrics, processed_people, unassigned_people):
    """Generate the Using Gifts Ministry Dashboard HTML"""

    # Extract key numbers for top stats
    overall = metrics['overall']
    strategic = metrics['strategic']

    total_serving = overall['serving_count']
    total_word_based = overall['word_based_count']
    total_new_serving = overall['new_to_serving']
    service_load_balance = overall['load_concentration']

    # Colors for congregations
    colors = {
        '8:30': '#dc2626',    # Red
        '10:30': '#2563eb',   # Blue  
        '6:30': '#059669',    # Green
        'overall': '#7c3aed'  # Purple
    }

    return f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Using Gifts Ministry Dashboard</title>
<style>
* {{margin:0;padding:0;box-sizing:border-box}}
body {{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f8fafc;padding:15px;color:#1e293b}}
.container {{max-width:1000px;margin:0 auto;background:white;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);overflow:hidden}}
.header {{background:linear-gradient(135deg,#7c3aed,#3b82f6);color:white;padding:20px;text-align:center}}
.header h1 {{font-size:24px;font-weight:700;margin-bottom:5px}}
.header p {{font-size:14px;opacity:0.9}}

.stats-bar {{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;padding:20px;background:#f1f5f9}}
.stat-card {{background:white;padding:15px;border-radius:8px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.05)}}
.stat-value {{font-size:28px;font-weight:700;color:#1e293b;margin-bottom:5px}}
.stat-label {{font-size:11px;color:#64748b;font-weight:500;text-transform:uppercase;letter-spacing:0.5px}}

.main-content {{padding:20px}}
.section-title {{font-size:18px;font-weight:600;margin-bottom:15px;color:#1e293b}}

.overview-grid {{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:25px}}
.overview-card {{background:#f8fafc;padding:15px;border-radius:8px;text-align:center;border-left:4px solid}}
.overview-card.overall {{border-left-color:#7c3aed}}
.overview-card.c830 {{border-left-color:#dc2626}}
.overview-card.c1030 {{border-left-color:#2563eb}}
.overview-card.c630 {{border-left-color:#059669}}

.overview-percentage {{font-size:32px;font-weight:700;margin-bottom:5px}}
.overview-percentage.overall {{color:#7c3aed}}
.overview-percentage.c830 {{color:#dc2626}}
.overview-percentage.c1030 {{color:#2563eb}}
.overview-percentage.c630 {{color:#059669}}
.overview-label {{font-size:13px;color:#64748b;font-weight:500}}
.overview-detail {{font-size:11px;color:#94a3b8;margin-top:5px}}

.congregation-grid {{display:grid;grid-template-columns:repeat(3,1fr);gap:20px}}
.congregation-column {{background:#f8fafc;border-radius:8px;padding:15px}}
.congregation-header {{text-align:center;margin-bottom:15px;padding:10px;border-radius:6px;color:white;font-weight:600}}
.congregation-header.c830 {{background:#dc2626}}
.congregation-header.c1030 {{background:#2563eb}}
.congregation-header.c630 {{background:#059669}}

.metric-card {{background:white;border-radius:6px;padding:12px;margin-bottom:10px;border-left:3px solid #e2e8f0}}
.metric-title {{font-size:12px;color:#64748b;font-weight:500;margin-bottom:8px}}
.metric-value {{font-size:20px;font-weight:700;color:#1e293b;margin-bottom:4px}}
.metric-detail {{font-size:10px;color:#94a3b8}}

.strategic-targets {{background:#fef3c7;border:1px solid #fbbf24;border-radius:8px;padding:15px;margin-top:20px}}
.strategic-targets h3 {{color:#92400e;margin-bottom:10px;font-size:16px}}
.strategic-grid {{display:grid;grid-template-columns:repeat(2,1fr);gap:15px}}
.strategic-card {{background:white;padding:12px;border-radius:6px}}
.strategic-value {{font-size:18px;font-weight:700;color:#92400e}}
.strategic-label {{font-size:11px;color:#92400e;margin-top:3px}}

.footnote {{text-align:center;margin-top:20px;padding:15px;background:#f1f5f9;border-radius:8px}}
.footnote p {{font-size:10px;color:#64748b;line-height:1.4}}

@media print {{
    body {{padding:0}}
    .container {{box-shadow:none;border:1px solid #e2e8f0}}
    .stats-bar {{background:white;border-bottom:1px solid #e2e8f0}}
    .congregation-column {{background:white;border:1px solid #e2e8f0}}
    .strategic-targets {{background:white;border:1px solid #fbbf24}}
    .unassigned-section {{background:white;border:1px solid #fca5a5}}
}}
</style></head><body>

<div class="container">
<div class="header">
<h1>🎯 Using Gifts Ministry Dashboard</h1>
<p>Service Participation & Leadership Development • Generated {datetime.now().strftime('%B %d, %Y')}</p>
</div>

<div class="stats-bar">
<div class="stat-card">
<div class="stat-value">{total_serving}</div>
<div class="stat-label">Total in Serving Roles</div>
</div>
<div class="stat-card">
<div class="stat-value">{total_word_based}</div>
<div class="stat-label">Word-Based Ministry</div>
</div>
<div class="stat-card">
<div class="stat-value">{total_new_serving}</div>
<div class="stat-label">New to Serving This Calendar Year</div>
</div>
<div class="stat-card">
<div class="stat-value">{service_load_balance:.0f}%</div>
<div class="stat-label">Service Load Balance (% roles held by top 20%)</div>
</div>
</div>

<div class="main-content">
<div class="section-title">📊 Service Participation by Congregation</div>
<div class="overview-grid">
<div class="overview-card overall">
<div class="overview-percentage overall">{metrics['overall']['serve_percentage']:.0f}%</div>
<div class="overview-label">Overall Serve Percentage</div>
<div class="overview-detail">{metrics['overall']['serving_count']} of {metrics['overall']['total_members']} members</div>
</div>
<div class="overview-card c830">
<div class="overview-percentage c830">{metrics['8:30']['serve_percentage']:.0f}%</div>
<div class="overview-label">8:30 AM Serve Percentage</div>
<div class="overview-detail">{metrics['8:30']['serving_count']} of {metrics['8:30']['total_members']} members</div>
</div>
<div class="overview-card c1030">
<div class="overview-percentage c1030">{metrics['10:30']['serve_percentage']:.0f}%</div>
<div class="overview-label">10:30 AM Serve Percentage</div>
<div class="overview-detail">{metrics['10:30']['serving_count']} of {metrics['10:30']['total_members']} members</div>
</div>
<div class="overview-card c630">
<div class="overview-percentage c630">{metrics['6:30']['serve_percentage']:.0f}%</div>
<div class="overview-label">6:30 PM Serve Percentage</div>
<div class="overview-detail">{metrics['6:30']['serving_count']} of {metrics['6:30']['total_members']} members</div>
</div>
</div>

<div class="section-title">⚖️ Load Distribution & Growth by Congregation</div>
<div class="congregation-grid">
<div class="congregation-column">
<div class="congregation-header c830">8:30 AM Congregation</div>
<div class="metric-card">
<div class="metric-title">Service Load Balance (% roles held by top 20%)</div>
<div class="metric-value">{metrics['8:30']['load_concentration']:.0f}%</div>
<div class="metric-detail">Concentration among active servers</div>
</div>
<div class="metric-card">
<div class="metric-title">Average Roles per Person</div>
<div class="metric-value">{metrics['8:30']['avg_roles_per_person']:.1f}</div>
<div class="metric-detail">Roles per serving member</div>
</div>
<div class="metric-card">
<div class="metric-title">New to Serving</div>
<div class="metric-value">{metrics['8:30']['new_to_serving']}</div>
<div class="metric-detail">People added to serving this calendar year</div>
</div>
<div class="metric-card">
<div class="metric-title">Word-Based Ministry Growth</div>
<div class="metric-value">{metrics['8:30']['new_word_based']}</div>
<div class="metric-detail">New in word-based roles this calendar year</div>
</div>
</div>

<div class="congregation-column">
<div class="congregation-header c1030">10:30 AM Congregation</div>
<div class="metric-card">
<div class="metric-title">Service Load Balance (% roles held by top 20%)</div>
<div class="metric-value">{metrics['10:30']['load_concentration']:.0f}%</div>
<div class="metric-detail">Concentration among active servers</div>
</div>
<div class="metric-card">
<div class="metric-title">Average Roles per Person</div>
<div class="metric-value">{metrics['10:30']['avg_roles_per_person']:.1f}</div>
<div class="metric-detail">Roles per serving member</div>
</div>
<div class="metric-card">
<div class="metric-title">New to Serving</div>
<div class="metric-value">{metrics['10:30']['new_to_serving']}</div>
<div class="metric-detail">People added to serving (this calendar year)</div>
</div>
<div class="metric-card">
<div class="metric-title">Word-Based Ministry Growth</div>
<div class="metric-value">{metrics['10:30']['new_word_based']}</div>
<div class="metric-detail">New in word-based roles (this calendar year)</div>
</div>
</div>

<div class="congregation-column">
<div class="congregation-header c630">6:30 PM Congregation</div>
<div class="metric-card">
<div class="metric-title">Service Load Balance (% roles held by top 20%)</div>
<div class="metric-value">{metrics['6:30']['load_concentration']:.0f}%</div>
<div class="metric-detail">Concentration among active servers</div>
</div>
<div class="metric-card">
<div class="metric-title">Average Roles per Person</div>
<div class="metric-value">{metrics['6:30']['avg_roles_per_person']:.1f}</div>
<div class="metric-detail">Roles per serving member</div>
</div>
<div class="metric-card">
<div class="metric-title">New to Serving</div>
<div class="metric-value">{metrics['6:30']['new_to_serving']}</div>
<div class="metric-detail">People added to serving (this calendar year)</div>
</div>
<div class="metric-card">
<div class="metric-title">Word-Based Ministry Growth</div>
<div class="metric-value">{metrics['6:30']['new_word_based']}</div>
<div class="metric-detail">New in word-based roles (this calendar year)</div>
</div>
</div>
</div>

<div class="strategic-targets">
<h3>🎯 Strategic Plan Targets (Using Gifts Ministry Area)</h3>
<div class="strategic-grid">
<div class="strategic-card">
<div class="strategic-value">{strategic['word_based_recruited_this_year']}</div>
<div class="strategic-label">new people in word-based ministry this calendar year<br>({strategic['word_based_recruitment_progress']:.0f}% of this calendar year's target of {strategic['word_based_needed_this_year']} people recruited)</div>
</div>
<div class="strategic-card">
<div class="strategic-value">{strategic['visitor_conversion_actual']:.1f}%</div>
<div class="strategic-label">Visitor to Serving Member Conversion Rate<br>{strategic['visitors_now_serving']} new serving members out of {strategic['visitors_total']} visitors this calendar year<br>({"better" if strategic['visitor_conversion_actual'] >= strategic['visitor_conversion_target'] else "worse"} than this calendar year's target ratio of {strategic['visitor_conversion_target']:.1f}%)</div>
</div>
</div>
</div>

<div class="footnote">
<p><strong>Notes:</strong> Service Load Balance measures concentration of serving roles (lower = better distribution across more people). 
Word-based ministry includes teaching, music, preaching, and group leadership roles. 
Strategic targets from 2025-2029 Church Plan: 10% annual word-based ministry growth (target: 5 new recruits this year).
Visitor to Serving Member Conversion Rate measures the percentage of first-time visitors in {datetime.now().year} who are now Serving Members (strategic goal: 5% annually).
All "new to serving" and "new word-based" metrics calculated for calendar year {datetime.now().year}.</p>
</div>
</div>
</div>

</body></html>'''

def main():
    try:
        # Create outputs directory if it doesn't exist
        outputs_dir = "outputs"
        if not os.path.exists(outputs_dir):
            os.makedirs(outputs_dir)
            print(f"✅ Created '{outputs_dir}' directory")
        
        # Fetch all data
        people = fetch_all_people()
        categories = fetch_categories()
        groups = fetch_groups()

        # Find and extract attendance data from both years
        current_group, last_year_group = find_attendance_report_groups()
        if not current_group:
            print("❌ Cannot proceed without current year attendance report group")
            return

        # Extract current year attendance data
        current_df, current_headers = extract_attendance_data_from_group(current_group, "current year")
        if current_df is None:
            print("❌ Cannot proceed without current year attendance data")
            return

        # Extract last year attendance data (for fallback assignments)
        last_year_df, last_year_headers = extract_attendance_data_from_group(last_year_group, "last year")

        # Create list of member names for filtering attendance data 
        # Use the same logic as Code CD to process members correctly
        print("\n📋 Processing people to identify active members (using Code CD logic)...")
        category_lookup = {}
        for cat in categories:
            category_lookup[cat.get('id', '')] = cat.get('name', '')

        # Process people exactly like Code CD does
        processed_people_for_filtering = []

        for person in people:
            category_id = person.get('category_id', '')
            category_name = category_lookup.get(category_id, '')
            person_type = categorize_person(category_name)

            try:
                date_added = datetime.strptime(person.get('date_added', ''), '%Y-%m-%d %H:%M:%S')
            except:
                date_added = datetime.now()

            processed_people_for_filtering.append({
                **person,
                'category_name': category_name,
                'type': person_type,
                'date_added': date_added
            })

        # Filter by type exactly like Code CD
        all_people_active = [p for p in processed_people_for_filtering if p['type'] != 'excluded']
        serving_members = [p for p in processed_people_for_filtering if p['type'] == 'serving_member']
        congregation = [p for p in processed_people_for_filtering if p['type'] in ['serving_member', 'congregation_only']]

        # Use congregation for calculations (this maintains Code CD logic)
        members = congregation

        # Create member names set for attendance filtering
        member_names = set()
        for member in members:
            full_name = f"{member.get('firstname', '')} {member.get('lastname', '')}".strip()
            if full_name:
                member_names.add(full_name)

        print(f"✅ Code CD filtering results:")
        print(f"   All active people: {len(all_people_active)}")
        print(f"   Serving members: {len(serving_members)}")
        print(f"   Congregation members: {len(congregation)}")
        print(f"   Total members for analysis: {len(members)}")
        print(f"   Member names for attendance filtering: {len(member_names)}")

        # Analyze congregation membership from both datasets (filtered to members only)
        current_assignments = analyze_congregation_membership_from_df(current_df, current_headers, "current year", member_names)

        last_year_assignments = {}
        if last_year_df is not None:
            last_year_assignments = analyze_congregation_membership_from_df(last_year_df, last_year_headers, "last year", member_names)

        # Combine assignments with two-tier priority: current year > last year
        congregation_assignments = create_two_tier_congregation_assignments(current_assignments, last_year_assignments)

        if not congregation_assignments:
            print("❌ Cannot proceed without congregation assignments")
            return

        if not people:
            print("❌ Failed to fetch data. Check your API key.")
            return

        # Calculate metrics (pass attendance data for visitor analysis)
        print(f"\n📖 Identifying people in word-based ministry:")
        metrics, processed_people, unassigned_people = calculate_metrics(people, categories, groups, congregation_assignments, current_df, current_headers, member_names)

        # Generate dashboard
        print("\n🎨 Generating Using Gifts Ministry Dashboard...")
        dashboard_html = generate_dashboard_html(metrics, processed_people, unassigned_people)

        # Save dashboard with consistent name (no timestamp - will overwrite)
        html_filename = "using_gifts_dashboard.html"
        html_filepath = os.path.join(outputs_dir, html_filename)

        with open(html_filepath, 'w', encoding='utf-8') as f:
            f.write(dashboard_html)

        print(f"✅ Dashboard saved: {html_filepath}")

        # Generate PNG image of the dashboard
        print("🖼️ Generating PNG image...")
        try:
            from html2image import Html2Image
            hti = Html2Image()

            # Set up for high-quality image with consistent name (no timestamp - will overwrite)
            png_filename = "using_gifts_dashboard.png"
            png_filepath = os.path.join(outputs_dir, png_filename)

            # Generate PNG from HTML file
            hti.screenshot(
                html_file=html_filepath,
                save_as=png_filename,
                size=(1200, 1600)  # Width x Height - good for dashboard layout
            )
            
            # html2image saves in current directory, so move it to outputs
            if os.path.exists(png_filename):
                import shutil
                shutil.move(png_filename, png_filepath)

            print(f"✅ PNG image saved: {png_filepath}")
            print(f"📁 Both files ready in the '{outputs_dir}' directory!")

        except Exception as e:
            print(f"⚠️ PNG generation failed (HTML still available): {e}")
            print(f"💡 You can take a screenshot manually or try refreshing and running again")

        # Print summary
        print(f"\n📊 USING GIFTS MINISTRY SUMMARY:")
        print(f"   Overall: {metrics['overall']['serve_percentage']:.1f}% serving ({metrics['overall']['serving_count']} of {metrics['overall']['total_members']})")
        print(f"   8:30 AM: {metrics['8:30']['serve_percentage']:.1f}% serving ({metrics['8:30']['serving_count']} of {metrics['8:30']['total_members']})")
        print(f"   10:30 AM: {metrics['10:30']['serve_percentage']:.1f}% serving ({metrics['10:30']['serving_count']} of {metrics['10:30']['total_members']})")
        print(f"   6:30 PM: {metrics['6:30']['serve_percentage']:.1f}% serving ({metrics['6:30']['serving_count']} of {metrics['6:30']['total_members']})")
        print(f"   Word-Based Ministry: {metrics['overall']['word_based_count']} people")
        print(f"   🎯 IMPROVED Strategic Progress: {metrics['strategic']['word_based_recruitment_progress']:.0f}% of annual recruitment target ({metrics['strategic']['word_based_recruited_this_year']} of {metrics['strategic']['word_based_needed_this_year']} needed)")
        print(f"   📊 Accurate word-based ministry recruitment by congregation (using location data):")
        print(f"      8:30 AM: {metrics['8:30']['new_word_based']} new recruits this calendar year")
        print(f"      10:30 AM: {metrics['10:30']['new_word_based']} new recruits this calendar year") 
        print(f"      6:30 PM: {metrics['6:30']['new_word_based']} new recruits this calendar year")
        print(f"   📊 Using accurate 'Report of New Serving Members' data with location-based congregation assignment!")
        print(f"   Visitor to Serving Member Conversion: {metrics['strategic']['visitor_conversion_actual']:.1f}% (target: 5.0%) - {metrics['strategic']['visitors_now_serving']} of {metrics['strategic']['visitors_total']} first-time visitors in {datetime.now().year} are now Serving Members")

        # Verification math
        individual_total = metrics['8:30']['total_members'] + metrics['10:30']['total_members'] + metrics['6:30']['total_members']
        unassigned_count = metrics['overall']['total_members'] - individual_total
        assignment_rate = (individual_total / metrics['overall']['total_members'] * 100) if metrics['overall']['total_members'] > 0 else 0

        print(f"\n🔍 VERIFICATION:")
        print(f"   Individual congregations total: {individual_total}")
        print(f"   Overall total: {metrics['overall']['total_members']}")
        print(f"   Unassigned: {unassigned_count} ({(unassigned_count/metrics['overall']['total_members']*100):.1f}%)")
        print(f"   Assignment success rate: {assignment_rate:.1f}%")
        print(f"   Two-tier system: Current year attendance + Last year fallback for active members only")
        print(f"   Note: Deceased and archived members automatically excluded from all calculations")
        print(f"   Unassigned members: {len(unassigned_people)} people not assigned to any congregation")
        if len(unassigned_people) > 0:
            print(f"   Unassigned members list:")
            for person in unassigned_people[:10]:  # Show first 10
                print(f"     • {person['name']} ({person['category']})")
            if len(unassigned_people) > 10:
                print(f"     • ... and {len(unassigned_people) - 10} more (see dashboard for full list)")

        # Open dashboard
        file_path = os.path.abspath(html_filepath)
        webbrowser.open(f"file://{file_path}")
        print(f"\n🌐 Dashboard opened!")
        print(f"📋 Files created in '{outputs_dir}' directory:")
        print(f"   • {html_filename} (interactive HTML)")
        print(f"   • {png_filename} (image for Google Docs)")
        print(f"💡 New runs will overwrite these files automatically")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
