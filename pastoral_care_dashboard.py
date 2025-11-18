import subprocess
import sys


# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'pandas', 'requests']
    for package in packages:
        try:
            if package == 'beautifulsoup4':
                import bs4
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
from bs4 import BeautifulSoup
import re

print("🏛️ ST GEORGE'S MAGILL - PASTORAL CARE REPORT")
print("="*50)
print("Identifying members who haven't attended in the last 4 weeks")
print("+ Recent newcomers who attended for the first time in last 6 weeks")

# Get API key
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


def is_person_excluded(person):
    """
    Check if a person should be excluded from dashboard calculations.
    Excludes deceased people and people with '*' categories.
    This is the EXACT same function from Code CD.
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
    """
    Fetch ALL people first, then filter out deceased - using Code CD approach
    """
    print("📋 Fetching all people (using Code CD deceased filtering approach)...")
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
            'fields': ['demographics', 'departments']  # API will automatically include deceased, archived, status fields
        })
        if not response or not response.get('people'):
            print("Done")
            break
        people = response['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []

        # Filter out deceased and other excluded people (EXACT Code CD approach)
        filtered_people = []
        for person in people:
            is_excluded, reason = is_person_excluded(person)
            if is_excluded:
                excluded_count += 1
                if reason == "deceased" or reason == "status_deceased":
                    deceased_count += 1
                elif reason == "archived":
                    archived_count += 1
                # Log excluded people for verification
                name = f"{person.get('firstname', '')} {person.get('lastname', '')}"
                print(f"\n   Excluding {name} ({reason})", end="")
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


def categorize_person(category):
    """From Code CD - categorize person by their category"""
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


def fetch_congregation_members():
    """
    Get congregation members using Code CD approach:
    1. Fetch ALL people (already deceased-filtered)
    2. Filter by category
    3. Filter by other criteria
    """
    print("\n👥 Filtering for congregation members (Congregation_ and RosteredMember_)...")

    # Get all living people first (deceased already filtered out)
    all_people = fetch_all_people()
    if not all_people:
        return []

    # Get categories to create a lookup
    categories_response = make_request('people/categories/getAll')
    if not categories_response:
        return []

    categories = categories_response['categories'].get('category', [])
    if not isinstance(categories, list):
        categories = [categories] if categories else []

    # Create category lookup
    category_lookup = {}
    for cat in categories:
        category_lookup[cat.get('id', '')] = cat.get('name', '')

    print(f"✅ Created lookup for {len(category_lookup)} total categories")

    # Now filter the already-living people by category and other criteria
    congregation_members = []
    excluded_counts = {
        'wrong_category': 0,
        'contact_only': 0,
        'no_name': 0
    }

    for person in all_people:
        # Check if person has a proper name
        first_name = (person.get('firstname') or '').strip()
        last_name = (person.get('lastname') or '').strip()
        if not first_name and not last_name:
            excluded_counts['no_name'] += 1
            continue

        # Check contact flag (contact-only people are not full members)
        contact = person.get('contact', 0) 
        if contact == 1 or str(contact).lower() == 'true':
            excluded_counts['contact_only'] += 1
            continue

        # Check category using Code CD logic
        category_id = person.get('category_id', '')
        category_name = category_lookup.get(category_id, '')
        person_type = categorize_person(category_name)

        # Only include serving members and congregation members
        if person_type not in ['serving_member', 'congregation_only']:
            excluded_counts['wrong_category'] += 1
            continue

        # This person passes all filters - they are a living congregation member
        congregation_members.append(person)

    print(f"\n✅ Found {len(congregation_members)} LIVING congregation members after filtering")
    print(f"   Excluded {excluded_counts['wrong_category']} people with wrong categories")
    print(f"   Excluded {excluded_counts['contact_only']} contact-only people")
    print(f"   Excluded {excluded_counts['no_name']} people with no name")

    # Show breakdown by category
    by_category = {}
    for person in congregation_members:
        category_id = person.get('category_id', '')
        category_name = category_lookup.get(category_id, 'Unknown')
        by_category[category_name] = by_category.get(category_name, 0) + 1

    print(f"\n📊 Final breakdown by category (living members only):")
    for cat_name, count in by_category.items():
        print(f"   • {cat_name}: {count} members")

    return congregation_members


def find_current_attendance_report():
    """Find the current year individual attendance report group"""
    print("\n📋 Searching for current attendance report...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return None

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    # Look for current year attendance report (not "last year")
    for group in groups:
        group_name = group.get('name', '').lower()
        if ('report of service individual attendance' in group_name and 
            'last year' not in group_name):
            print(f"✅ Found attendance report: {group.get('name')}")
            return group

    print("❌ Current year attendance report group not found")
    return None


def extract_attendance_data(group):
    """Extract attendance data from the report group"""
    print("\n🔍 Extracting attendance data from report...")

    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print("❌ No URL found in attendance report group")
        return None, None

    print(f"✅ Found report URL")

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
                    print(f"✅ Found attendance table with {len(headers)} columns")

                    # Extract all attendance records
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

                    print(f"✅ Extracted {len(attendance_records)} attendance records")
                    return pd.DataFrame(attendance_records), headers

        print("❌ Attendance table not found in report")
        return None, None

    except Exception as e:
        print(f"❌ Error extracting attendance data: {e}")
        return None, None


def parse_recent_service_columns(headers, num_sundays=4):
    """Parse and identify all services from the most recent N Sundays"""
    print(f"\n📅 Identifying all services from the {num_sundays} most recent Sundays...")

    today = datetime.now().date()  # Use date() for easier comparison
    print(f"   Today's date: {today}")

    all_past_services = []
    unparseable_count = 0
    future_count = 0
    non_sunday_count = 0

    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue

        # Parse service header (looking for date and time patterns)
        # Examples: "9:30 AMMorning Prayer 14/01" or "Communion 2nd Order 02/06/2024 8:30 AM"

        # Extract time (look for patterns like 8:30 AM, 10:30 AM, etc.)
        time_match = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', header, re.IGNORECASE)
        if not time_match:
            continue

        time_str = f"{time_match.group(1)} {time_match.group(2).upper()}"

        # Extract date - handle both DD/MM and DD/MM/YYYY formats
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
                # For DD/MM format, assume current year
                year = datetime.now().year

            service_date = datetime(year, month, day)
            service_date_only = service_date.date()

            # IMPORTANT: Only include services that have already happened
            if service_date_only >= today:
                future_count += 1
                continue

            # Check if it's a Sunday
            if service_date.weekday() != 6:  # Sunday = 6
                non_sunday_count += 1
                continue

            # Extract service name
            service_name = header
            service_name = re.sub(r'\d{1,2}/\d{1,2}(/\d{4})?\s*\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
            service_name = service_name.strip()

            all_past_services.append({
                'header': header,
                'date': service_date,
                'date_only': service_date_only,
                'time': time_str,
                'service_name': service_name
            })

        except ValueError:
            # Invalid date
            unparseable_count += 1
            continue

    # Group services by Sunday date
    sundays = {}
    for service in all_past_services:
        sunday_date = service['date_only']
        if sunday_date not in sundays:
            sundays[sunday_date] = []
        sundays[sunday_date].append(service)

    # Sort Sunday dates (most recent first) and take the requested number
    sorted_sundays = sorted(sundays.keys(), reverse=True)
    recent_sundays = sorted_sundays[:num_sundays]

    # Collect all services from those recent Sundays
    recent_services = []
    for sunday_date in recent_sundays:
        recent_services.extend(sundays[sunday_date])

    # Sort services by date and time for display
    recent_services.sort(key=lambda x: (x['date'], x['time']))

    print(f"✅ Found {len(all_past_services)} total past Sunday services across {len(sundays)} Sundays")
    print(f"   Selected {len(recent_services)} services from {len(recent_sundays)} most recent Sundays")
    print(f"   Excluded {future_count} future services")
    print(f"   Excluded {non_sunday_count} non-Sunday services")
    print(f"   Couldn't parse {unparseable_count} headers")

    # Show the Sundays and services we selected
    if recent_services:
        print(f"\n📅 Selected services from {len(recent_sundays)} most recent Sundays:")

        current_sunday = None
        for svc in recent_services:
            if svc['date_only'] != current_sunday:
                current_sunday = svc['date_only']
                days_ago = (today - current_sunday).days
                print(f"\n   📍 {current_sunday.strftime('%a %d %b %Y')} ({days_ago} days ago):")

            print(f"      • {svc['time']} - {svc['service_name']}")

        print(f"\n✅ Total: {len(recent_services)} services from {len(recent_sundays)} Sundays")
    else:
        print("❌ No recent Sunday services found!")

    return recent_services


def parse_all_service_columns_for_newcomers(headers):
    """Parse ALL service columns to identify newcomers from the entire attendance report"""
    print(f"\n📅 Parsing ALL services for newcomer analysis...")

    today = datetime.now().date()
    six_weeks_ago = today - timedelta(weeks=6)
    
    print(f"   Today's date: {today}")
    print(f"   Six weeks ago: {six_weeks_ago}")

    all_services = []
    unparseable_count = 0
    future_count = 0
    non_sunday_count = 0

    for header in headers:
        # Skip obvious non-service columns
        header_lower = header.lower()
        if any(skip in header_lower for skip in ['first name', 'last name', 'category', 'email', 'phone']):
            continue

        # Parse service header (looking for date and time patterns)
        time_match = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', header, re.IGNORECASE)
        if not time_match:
            continue

        time_str = f"{time_match.group(1)} {time_match.group(2).upper()}"

        # Extract date - handle both DD/MM and DD/MM/YYYY formats
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
                # For DD/MM format, assume current year
                year = datetime.now().year

            service_date = datetime(year, month, day)
            service_date_only = service_date.date()

            # Only include services that have already happened
            if service_date_only >= today:
                future_count += 1
                continue

            # Check if it's a Sunday
            if service_date.weekday() != 6:  # Sunday = 6
                non_sunday_count += 1
                continue

            # Extract service name
            service_name = header
            service_name = re.sub(r'\d{1,2}/\d{1,2}(/\d{4})?\s*\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
            service_name = service_name.strip()

            all_services.append({
                'header': header,
                'date': service_date,
                'date_only': service_date_only,
                'time': time_str,
                'service_name': service_name
            })

        except ValueError:
            # Invalid date
            unparseable_count += 1
            continue

    # Sort services by date
    all_services.sort(key=lambda x: x['date'])

    print(f"✅ Found {len(all_services)} total Sunday services")
    print(f"   Excluded {future_count} future services")
    print(f"   Excluded {non_sunday_count} non-Sunday services")
    print(f"   Couldn't parse {unparseable_count} headers")

    return all_services


def identify_newcomers(members, attendance_df, all_services):
    """Identify newcomers who attended for the first time in the last 6 weeks"""
    print(f"\n🆕 Identifying newcomers who attended for the first time in the last 6 weeks...")

    today = datetime.now().date()
    six_weeks_ago = today - timedelta(weeks=6)

    # Create a set of known member names for quick lookup
    member_names = set()
    for member in members:
        first_name = member.get('firstname', '').strip()
        last_name = member.get('lastname', '').strip()
        full_name = f"{first_name} {last_name}".strip().lower()
        if full_name:
            member_names.add(full_name)

    print(f"   Known members: {len(member_names)}")

    # Get all service headers sorted by date
    service_headers = [svc['header'] for svc in all_services]
    
    newcomers = []
    processed_people = set()  # To avoid duplicates

    # Go through each person in the attendance report
    for _, row in attendance_df.iterrows():
        first_name = str(row.get('First Name', '')).strip()
        last_name = str(row.get('Last Name', '')).strip()
        full_name = f"{first_name} {last_name}".strip()
        full_name_lower = full_name.lower()

        # Skip if no name or already processed
        if not full_name or full_name_lower in processed_people:
            continue
        
        processed_people.add(full_name_lower)

        # Skip if this person is already a known member
        if full_name_lower in member_names:
            continue

        # Find their first attendance
        first_attendance_date = None
        attended_services = []

        for service in all_services:
            header = service['header']
            if header in row and str(row[header]).strip().upper() == 'Y':
                attended_services.append(service)
                if first_attendance_date is None:
                    first_attendance_date = service['date_only']

        # Check if they are a newcomer (first attendance in last 6 weeks)
        if first_attendance_date and first_attendance_date >= six_weeks_ago:
            days_ago = (today - first_attendance_date).days
            
            newcomers.append({
                'name': full_name,
                'first_name': first_name,
                'last_name': last_name,
                'first_attendance': first_attendance_date,
                'days_ago': days_ago,
                'total_attendances': len(attended_services),
                'attended_services': attended_services
            })

    # Sort newcomers by most recent first
    newcomers.sort(key=lambda x: x['first_attendance'], reverse=True)

    print(f"✅ Found {len(newcomers)} newcomers who attended for the first time in the last 6 weeks")

    if newcomers:
        print(f"\n🆕 Newcomers list:")
        for newcomer in newcomers:
            first_date = newcomer['first_attendance'].strftime('%d %b %Y')
            print(f"   • {newcomer['name']} - First attended: {first_date} ({newcomer['days_ago']} days ago) - {newcomer['total_attendances']} total attendances")

    return newcomers


def identify_missing_members(members, attendance_df, recent_services):
    """Identify members who haven't attended any recent services"""
    print(f"\n🔍 Cross-referencing {len(members)} confirmed living members with attendance data...")

    # Get categories for lookup
    categories_response = make_request('people/categories/getAll')
    category_lookup = {}
    if categories_response:
        categories = categories_response['categories'].get('category', [])
        if not isinstance(categories, list):
            categories = [categories] if categories else []
        for cat in categories:
            category_lookup[cat.get('id', '')] = cat.get('name', '')

    # Create a lookup of members by name
    member_lookup = {}
    category_check = {}  # For debugging

    for member in members:
        first_name = member.get('firstname', '').strip()
        last_name = member.get('lastname', '').strip()
        full_name = f"{first_name} {last_name}".strip()

        # Get category name from lookup
        category_id = member.get('category_id', '')
        category_name = category_lookup.get(category_id, 'Unknown')

        # Debug: Track categories we're actually processing
        category_check[category_name] = category_check.get(category_name, 0) + 1

        if full_name and full_name != "":
            member_lookup[full_name.lower()] = {
                'id': member.get('id'),
                'first_name': first_name,
                'last_name': last_name,
                'category': category_name,
                'category_id': category_id,  # Keep for debugging
                'email': member.get('email', ''),
                'phone': member.get('phone', '')
            }

    print(f"✅ Created lookup for {len(member_lookup)} confirmed living members")
    print(f"📊 Category distribution in members being checked:")
    for cat_name, count in sorted(category_check.items()):
        print(f"   • {cat_name}: {count} members")

    # Check attendance for each member
    missing_members = []
    recent_service_headers = [svc['header'] for svc in recent_services]

    print(f"📊 Checking attendance across {len(recent_service_headers)} recent services...")

    found_in_report_count = 0
    not_found_in_report_count = 0

    for member_name, member_info in member_lookup.items():
        attended_any = False

        # Look for this member in the attendance data
        # Try different name matching approaches
        name_variations = [
            member_name,
            f"{member_info['first_name'].lower()} {member_info['last_name'].lower()}",
            f"{member_info['last_name'].lower()}, {member_info['first_name'].lower()}",
            f"{member_info['first_name'].lower()}{member_info['last_name'].lower()}"
        ]

        member_row = None
        for _, row in attendance_df.iterrows():
            row_first = str(row.get('First Name', '')).strip().lower()
            row_last = str(row.get('Last Name', '')).strip().lower()
            row_full = f"{row_first} {row_last}".strip()

            if row_full in name_variations or any(var == row_full for var in name_variations):
                member_row = row
                break

        if member_row is not None:
            found_in_report_count += 1
            # Check if they attended any recent service
            for service_header in recent_service_headers:
                if service_header in member_row:
                    attendance_value = str(member_row[service_header]).strip().upper()
                    if attendance_value == 'Y':
                        attended_any = True
                        break
        else:
            not_found_in_report_count += 1

        if not attended_any:
            missing_members.append({
                'name': f"{member_info['first_name']} {member_info['last_name']}",
                'first_name': member_info['first_name'],
                'last_name': member_info['last_name'],
                'category': member_info['category'],
                'email': member_info['email'],
                'phone': member_info['phone'],
                'id': member_info['id'],
                'found_in_report': member_row is not None
            })

    print(f"📋 Found {len(missing_members)} members who haven't attended recently")
    print(f"   {found_in_report_count} members found in attendance report")
    print(f"   {not_found_in_report_count} members not found in attendance report")

    # Final verification: Only include people who were found in the attendance report
    # (People not found in report are likely inactive or have data inconsistencies)
    verified_members = []
    excluded_not_found = 0

    target_categories = {'Congregation_', 'RosteredMember_'}

    for member in missing_members:
        # Must be found in attendance report (this excludes inactive people)
        if not member['found_in_report']:
            excluded_not_found += 1
            continue

        # This member passes all checks
        verified_members.append(member)

    if excluded_not_found > 0:
        print(f"✅ Excluded {excluded_not_found} people not found in attendance report (likely inactive)")

    missing_members = verified_members
    print(f"✅ Final report will include {len(missing_members)} confirmed living members who haven't attended recently")

    # Final verification: check categories of missing members
    missing_categories = {}
    for member in missing_members:
        cat = member['category']
        missing_categories[cat] = missing_categories.get(cat, 0) + 1

    print(f"📊 Categories of missing members:")
    for cat_name, count in sorted(missing_categories.items()):
        print(f"   • {cat_name}: {count} members")

    return missing_members


def generate_pastoral_care_report(missing_members, recent_services, newcomers=None, num_sundays=4):
    """Generate HTML report for pastoral care follow-up"""
    print("\n📄 Generating pastoral care report...")
    
    import os
    
    # Create outputs directory if it doesn't exist
    outputs_dir = 'outputs'
    os.makedirs(outputs_dir, exist_ok=True)

    # Calculate how many unique Sundays are represented
    unique_sundays = set()
    for svc in recent_services:
        unique_sundays.add(svc['date_only'])

    # Sort by category and then by name
    missing_members.sort(key=lambda x: (x['category'], x['last_name'], x['first_name']))

    # Group by category
    by_category = {}
    for member in missing_members:
        category = member['category'] or 'Unknown'
        if category not in by_category:
            by_category[category] = []
        by_category[category].append(member)

    # Generate services summary grouped by Sunday (compact format)
    services_html = ""
    if recent_services:
        services_html = "<ul style='margin: 3px 0; padding-left: 15px;'>"

        # Group services by Sunday for better display
        sunday_groups = {}
        for svc in recent_services:
            sunday_date = svc['date_only']
            if sunday_date not in sunday_groups:
                sunday_groups[sunday_date] = []
            sunday_groups[sunday_date].append(svc)

        # Sort Sundays most recent first
        for sunday_date in sorted(sunday_groups.keys(), reverse=True):
            days_ago = (datetime.now().date() - sunday_date).days
            services_html += f"<li style='font-size: 10px; margin: 0; line-height: 1.2;'><strong>{sunday_date.strftime('%d %b')}</strong> ({days_ago}d ago) - {len(sunday_groups[sunday_date])} services</li>"

        services_html += "</ul>"

    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>St George's Magill - Pastoral Care Report</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.4;
            color: #2c3e50;
            max-width: 1000px;
            margin: 0 auto;
            padding: 15px;
            background: #f5f7fa;
            font-size: 12px;
        }}
        .header {{
            background: linear-gradient(135deg, #5e72e4 0%, #825ee4 100%);
            color: white;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            margin-bottom: 15px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .header h1 {{
            margin: 0;
            font-size: 18px;
            font-weight: 600;
        }}
        .header p {{
            margin: 3px 0 0 0;
            font-size: 11px;
            opacity: 0.95;
        }}
        .summary {{
            background: white;
            padding: 12px;
            border-radius: 8px;
            margin-bottom: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            border-left: 4px solid #5e72e4;
        }}
        .summary h2 {{
            color: #2c3e50;
            margin: 0 0 8px 0;
            font-size: 14px;
            font-weight: 600;
        }}
        .summary p {{
            margin: 2px 0;
            font-size: 10px;
            color: #546e7a;
        }}
        .summary ul {{
            margin: 3px 0;
            padding-left: 15px;
        }}
        .summary li {{
            font-size: 10px;
            margin: 0;
            line-height: 1.2;
            color: #546e7a;
        }}
        .alert {{
            background: #fff3e0;
            border: 1px solid #ffb74d;
            border-left: 4px solid #ff9800;
            color: #e65100;
            padding: 8px;
            border-radius: 6px;
            margin: 8px 0;
            font-size: 11px;
        }}
        .success {{
            background: #e8f5e9;
            border: 1px solid #81c784;
            border-left: 4px solid #4caf50;
            color: #2e7d32;
            padding: 8px;
            border-radius: 6px;
            margin: 8px 0;
            font-size: 11px;
        }}
        .newcomer-alert {{
            background: #e3f2fd;
            border: 1px solid #64b5f6;
            border-left: 4px solid #2196f3;
            color: #1565c0;
            padding: 8px;
            border-radius: 6px;
            margin: 8px 0;
            font-size: 11px;
        }}
        .category-section {{
            background: white;
            margin-bottom: 10px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .category-header {{
            background: #5e72e4;
            color: white;
            padding: 6px 12px;
            font-size: 13px;
            font-weight: 600;
        }}
        .newcomer-header {{
            background: #2196f3;
            color: white;
            padding: 6px 12px;
            font-size: 13px;
            font-weight: 600;
        }}
        .member-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        .member-table th {{
            background: #f5f7fa;
            padding: 6px 10px;
            text-align: left;
            font-weight: 600;
            color: #2c3e50;
            border-bottom: 2px solid #e1e8ed;
            font-size: 11px;
        }}
        .member-table td {{
            padding: 6px 10px;
            border-bottom: 1px solid #f0f3f5;
            font-size: 11px;
        }}
        .member-table tr:hover {{
            background: #f8fafb;
        }}
        .contact-info {{
            font-size: 10px;
            color: #78909c;
        }}
        .newcomer-info {{
            font-size: 10px;
            color: #1976d2;
        }}
        .footer {{
            text-align: center;
            margin-top: 15px;
            padding: 8px;
            color: #90a4ae;
            font-size: 8px;
        }}
        @media print {{
            body {{ 
                background: white; 
                font-size: 9px;
                padding: 8px;
            }}
            .header {{ 
                background: #5e72e4 !important; 
                padding: 10px;
            }}
            .header h1 {{ font-size: 16px; }}
            .summary {{ padding: 8px; }}
            .summary h2 {{ font-size: 12px; }}
            .category-header, .newcomer-header {{ padding: 4px 8px; font-size: 11px; }}
            .member-table th, .member-table td {{ 
                padding: 3px 6px; 
                font-size: 9px;
            }}
            .footer {{ font-size: 7px; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>🏛️ St George's Magill - Pastoral Care Report</h1>
        <p>Members Needing Follow-up + Recent Newcomers - {datetime.now().strftime('%d %b %Y')}</p>
    </div>

    <div class="summary">
        <h2>📊 Summary</h2>
        <p><strong>Members Not Attending:</strong> {len(missing_members)} people • <strong>Services:</strong> {len(recent_services)} across {len(unique_sundays)} Sundays</p>
        <p><strong>Recent Newcomers:</strong> {len(newcomers) if newcomers else 0} people (first attended in last 6 weeks)</p>
        <p style="font-size: 10px; margin: 3px 0 1px 0;"><strong>Recent Sundays:</strong></p>
        {services_html}
    </div>
"""

    # Newcomers section (show first)
    if newcomers and len(newcomers) > 0:
        html_content += f"""
    <div class="newcomer-alert">
        <strong>🆕 Newcomers Opportunity:</strong> {len(newcomers)} new people attended for the first time in the last 6 weeks. Consider follow-up contact!
    </div>

    <div class="category-section">
        <div class="newcomer-header">
            🆕 Recent Newcomers ({len(newcomers)} people) - First Time Attendees
        </div>
        <table class="member-table">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>First Attendance & Details</th>
                </tr>
            </thead>
            <tbody>
"""
        for newcomer in newcomers:
            first_date = newcomer['first_attendance'].strftime('%d %b %Y')
            days_ago = newcomer['days_ago']
            total_attendances = newcomer['total_attendances']
            
            # Show recent services they attended
            recent_services_attended = []
            for service in newcomer['attended_services'][-3:]:  # Last 3 services
                service_date = service['date_only'].strftime('%d %b')
                recent_services_attended.append(f"{service_date} ({service['time']})")
            
            services_text = ", ".join(recent_services_attended)
            if len(newcomer['attended_services']) > 3:
                services_text += f" + {len(newcomer['attended_services']) - 3} more"

            html_content += f"""
            <tr>
                <td><strong>{newcomer['name']}</strong></td>
                <td class="newcomer-info">
                    📅 First: {first_date} ({days_ago} days ago)<br>
                    📊 Total: {total_attendances} attendance{'s' if total_attendances != 1 else ''}<br>
                    🏛️ Recent: {services_text}
                </td>
            </tr>
"""

        html_content += """
            </tbody>
        </table>
    </div>
"""

    # Missing members section
    if len(missing_members) == 0:
        html_content += """
    <div class="success">
        <strong>🎉 Excellent!</strong> All active congregation members have attended at least one recent service.
    </div>
"""
    else:
        html_content += f"""
    <div class="alert">
        <strong>🚨 Action Required:</strong> {len(missing_members)} members haven't attended any of the recent Sunday services.
    </div>
"""

        # Add each category section
        for category, members in by_category.items():
            html_content += f"""
    <div class="category-section">
        <div class="category-header">
            {category} ({len(members)} members)
        </div>
        <table class="member-table">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Contact Information</th>
                </tr>
            </thead>
            <tbody>
"""

            for member in members:
                contact_parts = []
                if member['email']:
                    contact_parts.append(f"📧 {member['email']}")
                if member['phone']:
                    contact_parts.append(f"📞 {member['phone']}")

                contact_info = "<br>".join(contact_parts) if contact_parts else "<em>No contact info</em>"

                html_content += f"""
                <tr>
                    <td><strong>{member['name']}</strong></td>
                    <td class="contact-info">{contact_info}</td>
                </tr>
"""

            html_content += """
            </tbody>
        </table>
    </div>
"""

    html_content += f"""
    <div class="footer">
        <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} • Newcomers: first attendance in last 6 weeks • For pastoral care use only</p>
    </div>
</body>
</html>"""

    # Use consistent filenames (no timestamp) in outputs directory
    html_filename = os.path.join(outputs_dir, 'pastoral_care_report.html')
    png_filename = os.path.join(outputs_dir, 'pastoral_care_report.png')
    
    # Save HTML file
    with open(html_filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"✅ HTML report saved: {html_filename}")

    # Generate PNG image
    print("🖼️ Generating PNG image...")
    try:
        from html2image import Html2Image
        hti = Html2Image()

        # Generate PNG from HTML file
        hti.screenshot(
            html_file=html_filename,
            save_as='pastoral_care_report.png',
            size=(1200, 1600)  # Width x Height
        )
        
        # html2image saves in current directory, so move it to outputs
        if os.path.exists('pastoral_care_report.png'):
            import shutil
            shutil.move('pastoral_care_report.png', png_filename)
            print(f"✅ PNG image saved: {png_filename}")
        else:
            print(f"⚠️ PNG file was not created (html2image may not be installed)")
            
    except ImportError:
        print(f"⚠️ PNG export skipped: html2image not installed")
        print(f"   To enable PNG export, install with: pip install html2image")
    except Exception as e:
        print(f"⚠️ PNG export failed: {e}")

    return html_filename


def main():
    """Main execution function"""
    try:
        print("🚀 Starting pastoral care analysis using Code CD deceased filtering approach...")
        print("   📋 Finding members who haven't attended in last 4 weeks")
        print("   🆕 Finding newcomers who attended for first time in last 6 weeks")

        # Step 1: Get congregation members (using Code CD approach)
        members = fetch_congregation_members()
        if not members:
            print("❌ No congregation members found")
            return

        # Step 2: Find attendance report
        attendance_group = find_current_attendance_report()
        if not attendance_group:
            print("❌ Cannot proceed without attendance report")
            return

        # Step 3: Extract attendance data
        attendance_df, headers = extract_attendance_data(attendance_group)
        if attendance_df is None:
            print("❌ Cannot proceed without attendance data")
            return

        # Step 4: Identify all services from the 4 most recent Sundays
        recent_services = parse_recent_service_columns(headers, num_sundays=4)
        if not recent_services:
            print("❌ No recent Sunday services found in data")
            return

        # Step 5: Parse ALL services for newcomer analysis
        all_services = parse_all_service_columns_for_newcomers(headers)

        # Step 6: Find missing members
        missing_members = identify_missing_members(members, attendance_df, recent_services)

        # Step 7: Identify newcomers who attended for the first time in last 6 weeks
        newcomers = identify_newcomers(members, attendance_df, all_services)

        # Step 8: Generate report (including newcomers)
        report_file = generate_pastoral_care_report(missing_members, recent_services, newcomers)

        print(f"\n🎉 PASTORAL CARE ANALYSIS COMPLETE!")
        print(f"📄 Report: {report_file}")
        print(f"👥 Members checked: {len(members)}")

        # Calculate unique Sundays
        unique_sundays = set()
        for svc in recent_services:
            unique_sundays.add(svc['date_only'])

        print(f"📅 Analyzed: {len(recent_services)} services from {len(unique_sundays)} most recent Sundays")
        print(f"🚨 Members needing follow-up: {len(missing_members)}")
        print(f"🆕 Recent newcomers: {len(newcomers)} (first attended in last 6 weeks)")
        print(f"✨ Used Code CD deceased filtering approach - should eliminate all deceased people")

        if missing_members:
            print(f"\n📋 SUMMARY BY CATEGORY:")
            by_category = {}
            for member in missing_members:
                category = member['category'] or 'Unknown'
                by_category[category] = by_category.get(category, 0) + 1

            for category, count in by_category.items():
                print(f"   • {category}: {count} members")
        else:
            print("\n🎉 Excellent! All confirmed living congregation members have attended at least one service on at least one recent Sunday.")

        if newcomers:
            print(f"\n🆕 NEWCOMERS SUMMARY:")
            for newcomer in newcomers[:5]:  # Show first 5
                first_date = newcomer['first_attendance'].strftime('%d %b %Y')
                print(f"   • {newcomer['name']} - First: {first_date} ({newcomer['days_ago']} days ago) - {newcomer['total_attendances']} attendances")
            if len(newcomers) > 5:
                print(f"   • ... and {len(newcomers) - 5} more newcomers")
        else:
            print("\n🆕 No newcomers found in the last 6 weeks.")

        # Open the report automatically
        try:
            import webbrowser
            import os
            file_path = os.path.abspath(report_file)
            webbrowser.open(f"file://{file_path}")
            print(f"\n🌐 Report opened automatically!")
        except Exception as e:
            print(f"\n⚠️ Couldn't auto-open report: {e}")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
