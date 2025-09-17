import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'pandas', 'requests', 'plotly', 'kaleido']
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

install_packages()

import requests
import json
from datetime import datetime, timedelta
import pandas as pd
from bs4 import BeautifulSoup
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import numpy as np
import re
import webbrowser

print("🎯 NEXTGEN DASHBOARD - Kids, Youth & Conversions Analytics")
print("=" * 60)

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

def demographic_names(person):
    """
    Return a set of the person's demographic names in lowercase.
    Handles both shapes returned by Elvanto:
      {"demographics": {"demographic": [{"name":"Youth"}, ...]}}
      {"demographics": ["Youth","Children"]}
    """
    demos = person.get('demographics')
    if not demos:
        return set()

    # If it's a dict with 'demographic' inside
    if isinstance(demos, dict):
        items = demos.get('demographic', [])
    else:
        items = demos

    if not isinstance(items, list):
        items = [items]

    names = set()
    for item in items:
        if isinstance(item, dict):
            n = (item.get('name') or item.get('id') or '').strip().lower()
        else:
            n = str(item).strip().lower()
        if n:
            names.add(n)
    return names

def get_age_category(person):
    names = demographic_names(person)

    # match common variants
    if any(('children' in n) or ('child' in n) or ('kids' in n) for n in names):
        return 'children'
    if any(('youth' in n) or ('teen' in n) for n in names):
        return 'youth'
    if any('adult' in n for n in names):
        return 'adult'
    return None

def is_kids_age(person):
    """Check if person is kids age (5-12) based on demographic category"""
    age_category = get_age_category(person)
    return age_category == 'children'

def is_youth_age(person):
    """Check if person is youth age (13-17) based on demographic category"""
    age_category = get_age_category(person)
    return age_category == 'youth'

def fetch_custom_fields():
    """Fetch all custom fields to identify potential date professed fields"""
    print("🔍 Fetching custom fields to identify 'Date Professed' field...")
    
    response = make_request('people/customFields/getAll', {})
    if not response:
        print("   ❌ Failed to fetch custom fields")
        return []
    
    custom_fields_data = response.get('custom_fields', {})
    custom_fields = custom_fields_data.get('custom_field', [])
    
    if not isinstance(custom_fields, list):
        custom_fields = [custom_fields] if custom_fields else []
    
    print(f"   📋 Found {len(custom_fields)} custom fields:")
    
    professed_candidates = []
    
    for field in custom_fields:
        field_id = field.get('id', '')
        field_name = field.get('name', '')
        field_type = field.get('type', '')
        
        print(f"      • '{field_name}' (ID: {field_id}, Type: {field_type})")
        
        # Look for fields that might contain professed/decision dates
        name_lower = field_name.lower()
        if any(keyword in name_lower for keyword in ['professed', 'profession', 'decision', 'conversion', 'faith', 'saved', 'born again']):
            professed_candidates.append({
                'id': field_id,
                'name': field_name,
                'type': field_type
            })
            print(f"         🎯 POTENTIAL PROFESSED FIELD!")
    
    if professed_candidates:
        print(f"   ✅ Found {len(professed_candidates)} potential professed fields")
    else:
        print(f"   ⚠️ No obvious professed fields found - will check all text fields")
        # If no obvious candidates, include all text fields as potential professed fields
        for field in custom_fields:
            if field.get('type') in ['text', 'date', 'string']:
                professed_candidates.append({
                    'id': field.get('id', ''),
                    'name': field.get('name', ''),
                    'type': field.get('type', '')
                })
    
    return professed_candidates

def get_date_professed(person, professed_custom_fields):
    """Extract date professed from custom fields"""
    for field in professed_custom_fields:
        field_id = field['id']
        custom_field_key = f"custom_{field_id}"
        
        if custom_field_key in person:
            field_value = person[custom_field_key]
            
            # Handle nested structure (common in Elvanto)
            if isinstance(field_value, dict):
                field_value = field_value.get('value') or field_value.get('text', '')
            
            if field_value:
                field_value_str = str(field_value).strip()
                if field_value_str and field_value_str != '0':
                    # Try to parse as date
                    parsed_date = parse_date_string(field_value_str)
                    if parsed_date:
                        return parsed_date
    
    return None

def parse_date_string(date_string):
    """Parse various date string formats into datetime object"""
    if not date_string or date_string.strip() == '':
        return None
    
    date_string = str(date_string).strip()
    
    # Common date formats to try
    formats = [
        '%Y-%m-%d',      # 2024-01-15
        '%d/%m/%Y',      # 15/01/2024
        '%m/%d/%Y',      # 01/15/2024
        '%d-%m-%Y',      # 15-01-2024
        '%Y/%m/%d',      # 2024/01/15
        '%d/%m/%y',      # 15/01/24
        '%m/%d/%y',      # 01/15/24
        '%Y%m%d',        # 20240115
        '%d %B %Y',      # 15 January 2024
        '%B %d, %Y',     # January 15, 2024
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    
    # If standard formats fail, try regex parsing for flexible formats
    try:
        import re
        
        # Pattern: DD/MM/YYYY or MM/DD/YYYY or YYYY/MM/DD
        number_pattern = r'(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})'
        match = re.search(number_pattern, date_string)
        
        if match:
            part1, part2, part3 = match.groups()
            
            # Try to determine which is year, month, day
            parts = [int(part1), int(part2), int(part3)]
            
            # If one number is > 31, it's probably the year
            year_candidates = [p for p in parts if p > 31]
            if len(year_candidates) == 1:
                year = year_candidates[0]
                remaining = [p for p in parts if p != year]
                
                # If one of remaining is > 12, it's probably day
                if remaining[0] > 12:
                    day, month = remaining[0], remaining[1] 
                elif remaining[1] > 12:
                    month, day = remaining[0], remaining[1]
                else:
                    # Ambiguous - assume DD/MM format (common in Australia)
                    day, month = remaining[0], remaining[1]
                
                # Validate ranges
                if 1 <= month <= 12 and 1 <= day <= 31 and 1900 <= year <= 2030:
                    return datetime(year, month, day)
        
    except (ValueError, IndexError):
        pass
    
    return None

def get_all_people_with_demographics():
    """Get all people with demographics and custom fields (Code GD approach)"""
    print("\n👥 Fetching all people with demographics and custom fields...")
    
    # First, discover custom fields that might contain professed dates
    professed_custom_fields = fetch_custom_fields()
    
    # Build fields array including custom fields (NO date_of_birth - it doesn't exist)
    fields_to_request = ['demographics', 'departments', 'locations', 'family']
    
    # Add custom field IDs to the request
    for field in professed_custom_fields:
        custom_field_name = f"custom_{field['id']}"
        fields_to_request.append(custom_field_name)
        print(f"   📋 Adding custom field to request: {custom_field_name} ('{field['name']}')")
    
    print(f"   📋 Total fields to request: {len(fields_to_request)}")
    
    all_people = []
    page = 1
    
    while True:
        print(f"Page {page}...", end=" ")
        response = make_request('people/getAll', {
            'page': page,
            'page_size': 1000,
            'fields': fields_to_request
        })
        
        if not response:
            break
        
        people = response['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []
        
        if not people:
            print("Done")
            break
        
        all_people.extend(people)
        print(f"({len(people)} people)")
        page += 1
        
        if len(people) < 1000:
            break
    
    print(f"✅ Retrieved {len(all_people)} people with demographics and custom fields")
    
    # Return both people and custom fields for later use
    return all_people, professed_custom_fields

def get_conversions(people_data, professed_custom_fields):
    """Get NextGen conversions by year from Date Professed custom field (children & youth only)"""
    print("\n✝️ Fetching NextGen conversions data from custom fields...")
    
    conversions_by_year = {}
    people_checked = 0
    professed_fields_found = 0
    nextgen_conversions_found = 0
    
    # Debug: Track demographic parsing
    demo_debug_count = 0
    
    for person in people_data:
        people_checked += 1
        
        date_professed = get_date_professed(person, professed_custom_fields)
        if date_professed:
            professed_fields_found += 1
            person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            
            # Debug demographic parsing for people with conversions
            age_category = get_age_category(person)
            demo_debug_count += 1
            
            if demo_debug_count <= 5:  # Show first 5 for debugging
                print(f"   🔍 DEBUG: {person_name} has Date Professed, checking demographics...")
                demo_names = demographic_names(person)
                print(f"      Demographics found: {demo_names}")
                print(f"      → Parsed category: '{age_category}'")
                print(f"      → is_kids_age: {is_kids_age(person)}")
                print(f"      → is_youth_age: {is_youth_age(person)}")
            
            # FILTER: Only include children and youth for NextGen dashboard
            if is_kids_age(person) or is_youth_age(person):
                nextgen_conversions_found += 1
                year = date_professed.year
                
                if year not in conversions_by_year:
                    conversions_by_year[year] = []
                
                conversions_by_year[year].append({
                    'name': person_name,
                    'date': date_professed,
                    'category': age_category
                })
                
                print(f"   ✅ NextGen conversion: {person_name} ({age_category}) on {date_professed.strftime('%d/%m/%Y')}")
            else:
                print(f"   ➖ Conversion found but not NextGen: {person_name} (category: '{age_category}')")
    
    # Sort conversions by date within each year
    for year in conversions_by_year:
        conversions_by_year[year].sort(key=lambda x: x['date'])
    
    print(f"   📊 Processing summary:")
    print(f"      • Total people checked: {people_checked}")
    print(f"      • Total professed fields found: {professed_fields_found}")
    print(f"      • NextGen conversions (children & youth only): {nextgen_conversions_found}")
    print(f"✅ Found NextGen conversions in {len(conversions_by_year)} years")
    
    for year, convs in sorted(conversions_by_year.items()):
        kids_count = sum(1 for c in convs if c['category'] == 'children')
        youth_count = sum(1 for c in convs if c['category'] == 'youth')
        print(f"   {year}: {len(convs)} conversions ({kids_count} children, {youth_count} youth)")
    
    return conversions_by_year

def fetch_people_categories():
    """Fetch people categories to resolve category_id -> name"""
    print("📋 Fetching people categories...")
    
    response = make_request('people/categories/getAll', {})
    if not response:
        print("   ❌ Failed to fetch people categories")
        return {}
    
    categories_data = response.get('categories', {})
    categories = categories_data.get('category', [])
    
    if not isinstance(categories, list):
        categories = [categories] if categories else []
    
    categories_by_id = {}
    for cat in categories:
        cat_id = cat.get('id')
        cat_name = cat.get('name', '')
        if cat_id:
            categories_by_id[cat_id] = cat_name
    
    print(f"   ✅ Found {len(categories_by_id)} people categories")
    return categories_by_id

def find_attendance_report_groups():
    """Find the 6 'Individual Attendance' reports (Group & Service) for
    two years ago, last year, and current year, regardless of wording order."""
    print("\n📋 Searching for attendance report groups...")

    resp = make_request('groups/getAll', {'page_size': 1000})
    if not resp:
        return None

    groups = resp.get('groups', {}).get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    current_year = datetime.now().year
    yr_map = {
        'two_years_ago': current_year - 2,
        'last_year':     current_year - 1,
        'current':       current_year
    }

    # bucket we'll fill
    report_groups = {
        'group_reports':   {'two_years_ago': None, 'last_year': None, 'current': None},
        'service_reports': {'two_years_ago': None, 'last_year': None, 'current': None}
    }

    def norm(s):
        return re.sub(r'\s+', ' ', (s or '').strip().lower())

    def detect_period(nm):
        # explicit phrases first
        if any(x in nm for x in ['two years ago', '2 years ago', 'two-years-ago']):
            return 'two_years_ago'
        if any(x in nm for x in ['last year', 'previous year', 'prev year', 'prev. year']):
            return 'last_year'
        # fallback: current year if neither above
        return 'current'

    def is_group_report(nm):
        return 'group' in nm and 'individual' in nm and 'attendance' in nm

    def is_service_report(nm):
        return 'service' in nm and 'individual' in nm and 'attendance' in nm

    print(f"   🔍 Scanning {len(groups)} groups for attendance reports...")
    
    for g in groups:
        name = norm(g.get('name', ''))
        if 'report' not in name or 'individual' not in name or 'attendance' not in name:
            continue

        period = detect_period(name)
        
        if is_group_report(name):
            if not report_groups['group_reports'][period]:
                report_groups['group_reports'][period] = g
                print(f"   ✅ Group report ({period}): {g.get('name')}")
            else:
                print(f"   ⚠️ Multiple group reports for {period}, using first: {g.get('name')}")
        
        elif is_service_report(name):
            if not report_groups['service_reports'][period]:
                report_groups['service_reports'][period] = g
                print(f"   ✅ Service report ({period}): {g.get('name')}")
            else:
                print(f"   ⚠️ Multiple service reports for {period}, using first: {g.get('name')}")

    # validate
    missing = []
    for rtype in ['group_reports', 'service_reports']:
        for period in ['two_years_ago', 'last_year', 'current']:
            if not report_groups[rtype][period]:
                missing.append(f"{rtype}.{period}")
    
    if missing:
        print(f"   ⚠️ Missing reports: {missing}")
    
    total_found = sum(1 for rtype in report_groups.values() for r in rtype.values() if r)
    print(f"   📊 Found {total_found}/6 expected reports")
    
    return report_groups if total_found > 0 else None

def download_attendance_report_data(group, year_key):
    """Download HTML content from attendance report group URL"""
    if not group:
        print(f"   ❌ No {year_key} group provided")
        return None
        
    group_name = group.get('name', 'Unknown')
    print(f"   📥 Downloading {year_key} data: {group_name}")
    
    # Extract URL from group location fields
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field]).strip()
            print(f"      📍 Found URL in {field}")
            break
    
    if not report_url:
        print(f"   ❌ No download URL found in group fields")
        return None
    
    try:
        print(f"   📡 Downloading report data...")
        response = requests.get(report_url, timeout=60)
        if response.status_code == 200:
            print(f"   ✅ Downloaded {len(response.text)} characters")
            return response.text
        else:
            print(f"   ❌ HTTP Error {response.status_code}")
            return None
    except Exception as e:
        print(f"   ❌ Download failed: {e}")
        return None

def normalize_name_from_html(html_name):
    """Normalize name from HTML to match against directory names"""
    if not html_name:
        return ""
    
    # Remove HTML entities and normalize
    normalized = html_name.replace('&nbsp;', ' ').replace('&amp;', '&')
    
    # Remove common suffixes that appear in HTML but not directory
    suffixes = [' (Child)', ' (Youth)', ' (Adult)', ' (Member)']
    for suffix in suffixes:
        if normalized.endswith(suffix):
            normalized = normalized[:-len(suffix)].strip()
    
    return normalized.lower()

def match_html_name_to_person(html_name, name_lookup):
    """Match an HTML attendance name to a person from the group roster"""
    normalized_html = normalize_name_from_html(html_name)
    
    # Try exact match first
    if normalized_html in name_lookup:
        return name_lookup[normalized_html]
    
    # Try parsing "Last, First" format and converting to "First Last"
    if ',' in normalized_html:
        parts = normalized_html.split(',', 1)
        if len(parts) == 2:
            last_name = parts[0].strip()
            first_name = parts[1].strip()
            first_last_format = f"{first_name} {last_name}"
            
            if first_last_format in name_lookup:
                return name_lookup[first_last_format]
    
    # Try parsing "First Last" format and converting to "Last, First"  
    else:
        parts = normalized_html.split()
        if len(parts) >= 2:
            first_name = parts[0]
            last_name = ' '.join(parts[1:])
            last_first_format = f"{last_name}, {first_name}"
            
            if last_first_format in name_lookup:
                return name_lookup[last_first_format]
    
    # Try partial matching (first name prefix)
    for name_key, person_info in name_lookup.items():
        if normalized_html.startswith(person_info['firstname'].lower()) and person_info['lastname'].lower() in normalized_html:
            print(f"   🔍 Partial match: '{html_name}' → {person_info['full_name']}")
            return person_info
    
    print(f"   ❓ No match found for: '{html_name}'")
    return None

def get_group_members_by_name(group_name):
    """Get group roster as name_key -> person_info dict"""
    print(f"\n👥 Getting members of group: {group_name}")

    # Fetch all groups and find the one we want
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        print("   ❌ Failed to get groups list")
        return {}

    groups_data = response.get('groups', {})
    groups = groups_data.get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    print(f"   🔍 Searching through {len(groups)} groups for: {group_name}")

    target_group = None
    for group in groups:
        if not isinstance(group, dict):
            continue
        current_group_name = group.get('name', '')
        if current_group_name.lower() == group_name.lower():
            target_group = group
            print(f"   ✅ Found exact match: {current_group_name}")
            break
        elif group_name.lower() in current_group_name.lower():
            target_group = group
            print(f"   ✅ Found partial match: {current_group_name}")
            # keep looking for an exact match

    if not target_group:
        print(f"   ❌ Group not found: {group_name}")
        return {}

    group_id = target_group.get('id')
    if not group_id:
        print(f"   ❌ No group ID found for: {group_name}")
        return {}

    print(f"   📋 Getting members for group ID: {group_id}")

    # Ask for roster
    response = make_request('groups/getInfo', {'id': group_id, 'fields': ['people']})
    if not response:
        print(f"   ❌ Failed to get group info for: {group_name}")
        return {}

    group_obj = response.get('group', {})
    if isinstance(group_obj, list):  # rare variant
        group_obj = group_obj[0] if group_obj else {}

    people_node = group_obj.get('people', None)

    # --- COERCE to a list of person dicts, regardless of structure ---
    people_list = []
    if people_node:
        if isinstance(people_node, list):
            people_list = people_node
        elif isinstance(people_node, dict):
            person_data = people_node.get('person', [])
            if isinstance(person_data, list):
                people_list = person_data
            elif isinstance(person_data, dict):
                people_list = [person_data]

    print(f"   📊 Found {len(people_list)} members")

    # Build lookup dict
    name_lookup = {}
    for person in people_list:
        if not isinstance(person, dict):
            continue
        
        person_id = person.get('id', '')
        first = (person.get('firstname') or '').strip()
        last = (person.get('lastname') or '').strip()
        
        if first and last:
            full_name = f"{first} {last}"
            
            person_info = {
                'id': person_id,
                'firstname': first,
                'lastname': last,
                'full_name': full_name
            }
            
            # Store under multiple key formats for flexible matching
            name_lookup[f"{first} {last}".lower()] = person_info
            name_lookup[f"{last}, {first}".lower()] = person_info

    print(f"   ✅ Built name lookup with {len(name_lookup)} entries")
    return name_lookup

def extract_attendance_data_from_group(group, year_key):
    """Extract attendance data from 'Report of Group Individual Attendance' pages.
    Returns dict like: {'kids_club': {...}, 'youth_group': {...}, 'buzz': {...}, etc}"""
    
    print(f"\n🔍 Extracting {year_key} group attendance data...")
    
    # Download the HTML report
    html_content = download_attendance_report_data(group, year_key)
    if not html_content:
        print(f"   ❌ Could not download {year_key} group report data")
        return {}
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find the main table
    table = soup.find('table')
    if not table:
        print(f"   ❌ No table found in {year_key} report")
        return {}
    
    rows = table.find_all('tr')
    if not rows:
        print(f"   ❌ No rows found in {year_key} table")
        return {}
    
    print(f"   📊 Found {len(rows)} rows in attendance table")
    
    # Find weekly date columns in header
    header_row = rows[0] if rows else None
    date_columns = []
    date_labels = []
    
    if header_row:
        cells = header_row.find_all(['th', 'td'])
        for i, cell in enumerate(cells):
            cell_text = cell.get_text().strip()
            # Look for date patterns (dd/mm format)
            if re.match(r'\d{1,2}/\d{1,2}', cell_text):
                date_columns.append(i)
                date_labels.append(cell_text)
        print(f"   📅 Found {len(date_columns)} date columns: {date_labels}")
    
    # Parse group sections
    group_data = {}
    current_group_name = None
    current_group_rows = []
    
    for row_idx, row in enumerate(rows[1:], 1):  # Skip header
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        # Convert cells to text for easier processing
        row_data = [cell.get_text().strip() for cell in cells]
        
        # Check if this is a group header row (contains group name)
        first_cell = cells[0]
        is_group_header = False
        
        # Check for styling that indicates group header (dark background, bold text, etc.)
        style = first_cell.get('style', '').lower()
        class_attr = ' '.join(first_cell.get('class', [])).lower()
        
        if ('background' in style and ('black' in style or 'dark' in style)) or \
           'header' in class_attr or \
           (len(row_data) > 0 and any(group_word in row_data[0].lower() 
                                      for group_word in ['kids club', 'youth group', 'buzz', 'jkc', 'kids church'])):
            is_group_header = True
            
        if is_group_header:
            # Save previous group if exists
            if current_group_name and current_group_rows:
                processed_group = process_group_section(current_group_name, current_group_rows, date_columns, date_labels, year_key)
                if processed_group:
                    group_data[processed_group['section_key']] = processed_group
            
            # Start new group
            current_group_name = row_data[0] if row_data else f"Unknown Group {row_idx}"
            current_group_rows = []
            print(f"   📋 Found group header: {current_group_name}")
            
        else:
            # This is a person row - add to current group
            if current_group_name and len(row_data) >= 3:  # Must have name + some data
                current_group_rows.append(row_data)
    
    # Process the final group
    if current_group_name and current_group_rows:
        processed_group = process_group_section(current_group_name, current_group_rows, date_columns, date_labels, year_key)
        if processed_group:
            group_data[processed_group['section_key']] = processed_group
    
    print(f"   ✅ Extracted data for {len(group_data)} groups: {list(group_data.keys())}")
    return group_data

def process_group_section(group_name, rows, date_columns, date_labels, year_key):
    """Process a single group section and return structured data"""
    
    # Determine section key from group name
    group_lower = group_name.lower()
    if 'kids club' in group_lower:
        section_key = 'kids_club'
    elif 'youth group' in group_lower or 'youth' in group_lower:
        section_key = 'youth_group'  
    elif 'buzz' in group_lower:
        section_key = 'buzz'
    elif 'jkc' in group_lower or 'junior kids church' in group_lower:
        section_key = 'jkc'
    elif 'kids church' in group_lower and 'junior' not in group_lower:
        section_key = 'kids_church'
    else:
        section_key = re.sub(r'[^a-z0-9_]', '_', group_lower)
    
    print(f"      📊 Processing {group_name} as '{section_key}' with {len(rows)} people")
    
    # Count weekly attendance across all date columns
    weekly_counts = []
    for date_idx in date_columns:
        weekly_total = 0
        for row in rows:
            if date_idx < len(row):
                cell_value = row[date_idx].strip()
                # Count 'Y' (yes) attendance marks
                if cell_value.upper() == 'Y':
                    weekly_total += 1
        weekly_counts.append(weekly_total)
    
    # Track people with >= 2 attendance
    people_attendance = {}
    matched_people = {}  # Store by person ID when available
    
    for row in rows:
        if len(row) < 2:  # Need at least name + some data
            continue
            
        person_name = row[0].strip()
        if not person_name or person_name.lower() in ['first name', 'name']:
            continue
        
        # Count attendance across date columns
        attendance_count = 0
        for date_idx in date_columns:
            if date_idx < len(row):
                cell_value = row[date_idx].strip()
                if cell_value.upper() == 'Y':
                    attendance_count += 1
        
        # Store people with >= 2 attendance
        if attendance_count >= 2:
            people_attendance[person_name] = attendance_count
            
            # Try to match to person ID (for more accurate metrics)
            # This would require group roster lookup - simplified for now
            matched_people[f"unknown_id_{len(matched_people)}"] = {
                'name': person_name,
                'attendance_count': attendance_count
            }
    
    result = {
        'section_key': section_key,
        'group_name': group_name,
        'people': people_attendance,
        'matched_people': matched_people,
        'weekly_counts': weekly_counts,
        'date_labels': date_labels
    }
    
    print(f"         ✅ {len(people_attendance)} people with ≥2 attendance")
    return result

def parse_service_attendance(group, year_key, year):
    """Parse service attendance data from service report group"""
    print(f"\n🔍 Parsing {year_key} service attendance data...")
    
    try:
        # Download the HTML report
        html_content = download_attendance_report_data(group, year_key)
        if not html_content:
            return {}
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find the attendance table
        table = soup.find('table')
        if not table:
            print(f"   ❌ No table found in {year_key} service report")
            return {}
        
        rows = table.find_all('tr')
        if len(rows) < 2:  # Need header + at least one data row
            print(f"   ❌ Insufficient rows in {year_key} service table")
            return {}
        
        # Parse header to find service columns
        header_row = rows[0]
        headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
        
        service_columns = []
        for i, header in enumerate(headers):
            if '10:30' in header and 'AM' in header.upper():
                service_columns.append(('10:30 AM', i))
            elif '6:30' in header and 'PM' in header.upper():
                service_columns.append(('6:30 PM', i))
        
        if not service_columns:
            print(f"   ❌ No 10:30 AM or 6:30 PM service columns found")
            return {}
        
        print(f"   📊 Found service columns: {[s[0] for s in service_columns]}")
        
        # Process attendance data
        service_attendance = {'10:30 AM': {}, '6:30 PM': {}}
        
        for row in rows[1:]:  # Skip header
            cells = row.find_all(['td', 'th'])
            if len(cells) <= max(col_idx for _, col_idx in service_columns):
                continue
                
            # Get person name from first column
            person_name = cells[0].get_text(strip=True)
            if not person_name or person_name.lower() in ['first name', 'name', '']:
                continue
            
            # Count attendance for each service
            for service_time, col_idx in service_columns:
                if col_idx < len(cells):
                    cell_value = cells[col_idx].get_text(strip=True)
                    # Count 'Y' marks
                    y_count = cell_value.upper().count('Y')
                    if y_count > 0:
                        if person_name not in service_attendance[service_time]:
                            service_attendance[service_time][person_name] = 0
                        service_attendance[service_time][person_name] += y_count
        
        # Filter to people with >= 2 attendance and prepare result
        ppl_1030 = set()
        ppl_630 = set()
        ppl_overall = set()
        
        for person_name, count in service_attendance['10:30 AM'].items():
            if count >= 2:
                ppl_1030.add(person_name)
                ppl_overall.add(person_name)
        
        for person_name, count in service_attendance['6:30 PM'].items():
            if count >= 2:
                ppl_630.add(person_name)
                ppl_overall.add(person_name)
        
        result = {
            '10:30 AM': len(ppl_1030),
            '6:30 PM': len(ppl_630),
            'overall': len(ppl_overall),
            'people_10_30': ppl_1030,
            'people_6_30': ppl_630,
            'people_overall': ppl_overall
        }
        
        print(f"   📊 Found {len(service_columns)} service columns")
        print(f"   ✅ Service attendance (≥2 times): 10:30 AM: {result['10:30 AM']}, 6:30 PM: {result['6:30 PM']}, Overall: {result['overall']}")
        return result
    
    except Exception as e:
        print(f"❌ Error parsing {year_key} service data: {e}")
        return {}

def build_global_name_lookup(people_data):
    """
    Map normalized names to person IDs using the full directory.
    Keys: "first last" and "last, first" (lowercased, trimmed).
    """
    lookup = {}
    for p in people_data:
        pid = p.get('id')
        first = (p.get('firstname') or '').strip()
        last  = (p.get('lastname')  or '').strip()
        if pid and first and last:
            lookup[f"{first} {last}".lower()] = pid
            lookup[f"{last}, {first}".lower()] = pid
    return lookup

def name_to_id(display_name, global_lookup):
    """
    Use normalize_name_from_html() and the global lookup to resolve a display name to a person ID.
    Tries "first last" and "last, first" variants.
    """
    normalized = normalize_name_from_html(display_name)
    if normalized in global_lookup:
        return global_lookup[normalized]
    
    # Try flipping first/last order
    parts = normalized.split()
    if len(parts) >= 2:
        first_part = parts[0]
        last_parts = ' '.join(parts[1:])
        flipped = f"{last_parts}, {first_part}"
        if flipped in global_lookup:
            return global_lookup[flipped]
    
    return None

def is_serving_by_category(person, categories_by_id):
    """
    Treat any category whose name begins with 'RosteredMember_' as serving.
    """
    cid = person.get('category_id')
    if not cid:
        return False
    name = categories_by_id.get(cid, '') or ''
    name_l = name.strip().lower()
    return name_l.startswith('rosteredmember_')

def build_rostered_ids(people_data, categories_by_id):
    """
    Identify people who 'serve' for the Serving% numerator.
    Uses both People Category 'RosteredMember_' (via category_id) and Departments.
    """
    print("\n🎯 Building rostered/serving member list...")

    # Method 1: People Categories starting with 'RosteredMember_'
    picked_by_category = []
    for person in people_data:
        if is_serving_by_category(person, categories_by_id):
            name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            if name:
                picked_by_category.append(name)

    # Method 2: Departments (people with any department assignment)
    picked_by_department = []
    for person in people_data:
        departments = person.get('departments')
        has_department = False
        if departments:
            if isinstance(departments, dict):
                dept_list = departments.get('department', [])
            else:
                dept_list = departments
            
            if not isinstance(dept_list, list):
                dept_list = [dept_list] if dept_list else []
            
            if dept_list and any(d for d in dept_list if d):  # Any non-empty department
                has_department = True
        
        if has_department:
            name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            if name and name not in picked_by_category:
                picked_by_department.append(name)

    # Method 3: Demographic fallback (people with 'serving' demographic)
    picked_by_demo = []
    for person in people_data:
        demo_names = demographic_names(person)
        if any('serving' in d for d in demo_names):
            name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            if name and name not in picked_by_category and name not in picked_by_department:
                picked_by_demo.append(name)

    # Combine all methods
    rostered_ids = set()
    
    # Add people by category
    global_lookup = build_global_name_lookup(people_data)
    for name in picked_by_category:
        pid = name_to_id(name, global_lookup)
        if pid:
            rostered_ids.add(pid)
    
    # Add people by department
    for name in picked_by_department:
        pid = name_to_id(name, global_lookup)
        if pid:
            rostered_ids.add(pid)
    
    # Add people by demographic
    for name in picked_by_demo:
        pid = name_to_id(name, global_lookup)
        if pid:
            rostered_ids.add(pid)

    print(f"   🧾 Rostered by Category: {len(picked_by_category)}")
    if picked_by_category:
        print("      e.g.", ", ".join(picked_by_category[:8]) + ("..." if len(picked_by_category) > 8 else ""))
    print(f"   🧾 Rostered by Department: {len(picked_by_department)}")
    if picked_by_department:
        print("      e.g.", ", ".join(picked_by_department[:8]) + ("..." if len(picked_by_department) > 8 else ""))
    print(f"   🧾 Rostered by Demographic fallback: {len(picked_by_demo)}")
    if picked_by_demo:
        print("      e.g.", ", ".join(picked_by_demo[:8]) + ("..." if len(picked_by_demo) > 8 else ""))

    print(f"   ✅ Total rostered detected: {len(rostered_ids)}")
    return rostered_ids

def calculate_rolling_average(values, window=4):
    """Calculate rolling average for smoothing data"""
    if not values or len(values) < window:
        return values
    
    rolling_values = []
    for i in range(len(values)):
        start_idx = max(0, i + 1 - window)
        recent_values = values[start_idx:i+1]
        rolling_avg = sum(recent_values) / len(recent_values)
        rolling_values.append(rolling_avg)
    
    return rolling_values

def calculate_metrics(group_weekly_data, service_attendance_data, people_data, conversions, categories_by_id):
    """
    Calculates:
      • Kids Club %: (Kids Club ≥2) / (Kids Church (older Sunday) ≥2 children)
      • Youth Group %: (Youth Group ≥2) / (service-attending youth ≥2)
      • Serving %: Eligible NextGen (Kids Church≥2 children ∪ Youth with ≥6 services if available, else ≥2) ∩ rostered

    Uses People Category 'RosteredMember_' (via category_id) and Departments to detect serving.
    """
    print("\n📈 Calculating NextGen metrics...")

    import re

    # ---- Demographic cohorts ----
    def get_demo_category(person):
        return get_age_category(person)  # your existing helper

    children_ids_all = set()
    youth_ids_all = set()
    people_index = {}  # id -> {'first','last','full'}

    for p in (people_data or []):
        pid = p.get('id')
        if not pid:
            continue
        first = (p.get('firstname') or '').strip()
        last  = (p.get('lastname') or '').strip()
        people_index[pid] = {'first': first, 'last': last, 'full': f"{first} {last}".strip()}
        cat = get_demo_category(p)
        if cat == 'children':
            children_ids_all.add(pid)
        elif cat == 'youth':
            youth_ids_all.add(pid)

    print(f"   📊 Found {len(children_ids_all)} Children and {len(youth_ids_all)} Youth by demographics")

    # ---- Name mapping for resolving service names -> IDs ----
    def _name_keys(first, last):
        a = f"{first} {last}".strip().lower()
        b = f"{last}, {first}".strip().lower() if first and last else ''
        return {a, b} if b else {a}

    global_name_map = {}
    for pid, rec in people_index.items():
        for k in _name_keys(rec['first'], rec['last']):
            if k:
                global_name_map.setdefault(k, set()).add(pid)

    def _ids_from_section(section_blob):
        """Return IDs with attendance >= 2 from a group section (prefers matched_people; else resolve via section roster)."""
        out = set()
        if not isinstance(section_blob, dict):
            return out

        matched = section_blob.get('matched_people')
        if isinstance(matched, dict) and matched:
            for pid, info in matched.items():
                try:
                    if int(info.get('attendance_count', 0)) >= 2:
                        out.add(pid)
                except Exception:
                    continue
            return out

        # Fallback: resolve raw 'people' names using live roster of the section
        people_map = section_blob.get('people') or {}
        if isinstance(people_map, dict) and people_map:
            title = section_blob.get('group_name') or ''
            section_lookup = get_group_members_by_name(title) if title else {}
            for raw_name, cnt in people_map.items():
                try:
                    cnt_int = int(cnt)
                except Exception:
                    cnt_int = 0
                if cnt_int < 2:
                    continue
                mp = match_html_name_to_person(raw_name, section_lookup)
                if mp and mp.get('id'):
                    out.add(mp['id'])
        return out

    def _ids_from_service_names(name_set, restrict_ids=None):
        """Map service attendee display-names -> person IDs; optionally restrict to a cohort (e.g., youth_ids_all)."""
        out = set()
        for raw_name in name_set:
            norm_name = normalize_name_from_html(raw_name)
            matched_ids = global_name_map.get(norm_name, set())
            for mid in matched_ids:
                if restrict_ids is None or mid in restrict_ids:
                    out.add(mid)
        return out

    # ---- Serving detection ----
    rostered_ids = build_rostered_ids(people_data, categories_by_id)

    # ---- Initialize metrics structure ----
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]

    metrics = {
        'church_kids_in_kids_club': {},
        'church_youth_in_youth_group': {},
        'kids_youth_serving': {},
        'conversions': conversions,
        'weekly_attendance': group_weekly_data  # pass through
    }

    # ---- Per-year calculations ----
    for i, year in enumerate(years):
        year_key = ['two_years_ago', 'last_year', 'current'][i]
        
        print(f"\n   📅 Calculating metrics for {year} ({year_key})...")

        # GROUP DATA
        gwd = group_weekly_data.get(year_key, {})
        
        # SERVICE DATA  
        svc_data = service_attendance_data.get(year_key, {})

        # 1) KIDS CLUB PARTICIPATION
        kc_ids = _ids_from_section(gwd.get('kids_club', {}))
        
        # Denominator: Children who attended Kids Church ≥2 times (older Sunday program)
        kc_denom_ids = _ids_from_section(gwd.get('kids_church', {}))
        kc_denom_ids &= children_ids_all  # safety filter
        
        if kc_denom_ids:
            kc_pct = (len(kc_ids) / len(kc_denom_ids)) * 100
            metrics['church_kids_in_kids_club'][year] = {
                'count': len(kc_ids),
                'total': len(kc_denom_ids), 
                'percentage': kc_pct
            }
            print(f"      🧒 Kids Club: {len(kc_ids)}/{len(kc_denom_ids)} = {kc_pct:.1f}%")
        else:
            metrics['church_kids_in_kids_club'][year] = {'count': 0, 'total': 0, 'percentage': 0.0}
            print("      🧒 Kids Club: 0/0 (no Kids Church data)")

        # 2) YOUTH GROUP PARTICIPATION
        yg_ids = _ids_from_section(gwd.get('youth_group', {}))
        
        # Denominator: Youth who attended services ≥2 times
        youth_names = svc_data.get('people_overall', set())
        youth_service_ids = _ids_from_service_names(youth_names, restrict_ids=youth_ids_all)
        
        if youth_service_ids:
            yg_pct = (len(yg_ids) / len(youth_service_ids)) * 100
            metrics['church_youth_in_youth_group'][year] = {
                'count': len(yg_ids),
                'total': len(youth_service_ids),
                'percentage': yg_pct
            }
            print(f"      🎯 Youth Group: {len(yg_ids)}/{len(youth_service_ids)} = {yg_pct:.1f}%")
        else:
            metrics['church_youth_in_youth_group'][year] = {'count': 0, 'total': 0, 'percentage': 0.0}
            print("      🎯 Youth Group: 0/0 (no service data)")

        # 3) SERVING PERCENTAGE
        # Denominator: Kids Church≥2 children ∪ Youth with ≥6 services (if available, else ≥2)
        
        # Try ≥6 services for youth (more stringent)
        youth_names_min6 = svc_data.get('people_overall_min6') or set()
        youth_service_ids_min6 = _ids_from_service_names(youth_names_min6, restrict_ids=youth_ids_all)
        youth_for_serving = youth_service_ids_min6 if youth_service_ids_min6 else youth_service_ids

        eligible_nextgen_ids = set(kc_denom_ids) | set(youth_for_serving)
        serving_ids_this_year = eligible_nextgen_ids & rostered_ids

        print(f"      🧮 Eligible for Serving (KC≥2 ∪ Youth services≥{'6' if youth_service_ids_min6 else '2'}): {len(eligible_nextgen_ids)}")

        if eligible_nextgen_ids:
            pct = (len(serving_ids_this_year) / len(eligible_nextgen_ids)) * 100
            metrics['kids_youth_serving'][year] = {
                'count': len(serving_ids_this_year),
                'total': len(eligible_nextgen_ids),
                'percentage': pct
            }
            print(f"      🎯 Serving (real): {len(serving_ids_this_year)}/{len(eligible_nextgen_ids)} = {pct:.1f}%")
        else:
            metrics['kids_youth_serving'][year] = {'count': 0, 'total': 0, 'percentage': 0.0}
            print("      🎯 Serving (real): 0/0")

    return metrics

def create_dashboard(metrics):
    """
    Build the NextGen dashboard.

    This version understands the new weekly_attendance structure:
      metrics['weekly_attendance'][year_key][section_key] is a dict like:
        {
          'people': { 'Name': count>=2, ... },
          'matched_people': { person_id: {name, attendance_count}, ... },
          'group_name': 'Kids Club',
          'weekly_counts': [int, int, ...],   # Y's per date column
          'date_labels':  [str, str, ...]     # headers for those columns
        }
    If an older list-based series is encountered, it still works.
    """
    print("\n🎨 Creating NextGen Dashboard...")

    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    year_labels = [str(y) for y in years]

    # Colors
    colors = {
        'primary': '#0ea5e9',      # Sky blue
        'secondary': '#8b5cf6',    # Purple
        'accent': '#f59e0b',       # Amber
        'success': '#10b981',      # Emerald
        'warning': '#f97316',      # Orange
        'danger': '#ef4444',       # Red
        'text': '#1e293b',         # Slate 800
        'text_light': '#64748b',   # Slate 500
        'background': 'white',
        'grid': 'rgba(148, 163, 184, 0.2)',
        'current_year': '#0ea5e9',
        'last_year': '#ef4444',
        'two_years_ago': '#10b981'
    }

    # Safe getters
    def pct_list(key):
        return [metrics.get(key, {}).get(y, {}).get('percentage', 0) for y in years]

    def count_list(key):
        return [metrics.get(key, {}).get(y, {}).get('count', 0) for y in years]

    def conversions_count(y):
        return len(metrics.get('conversions', {}).get(y, []))

    # Figure & subplots with clearer titles
    fig = make_subplots(
        rows=3, cols=3,
        subplot_titles=[
            "Kids Club Participation — Church Children (%)",
            "Youth Group Participation — Church Youth (%)",
            "Serving Rate — Eligible NextGen (%)",
            "NextGen Conversions per Year",
            "Weekly Attendance — Kids Club & Youth Group",
            "Weekly Attendance — Buzz Playgroup (Tue)",
            "Kids Club Participants — Church Children (Count)",
            "Youth Group Participants — Church Youth (Count)",
            "Key Metrics Summary"
        ],
        specs=[[{"type": "bar"}, {"type": "bar"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "scatter"}, {"type": "scatter"}],
               [{"type": "bar"}, {"type": "bar"}, {"type": "table"}]],
        vertical_spacing=0.12,
        horizontal_spacing=0.08
    )

    # Row 1: Percentages
    kids_club_values = pct_list('church_kids_in_kids_club')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=kids_club_values, name="Kids Club %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in kids_club_values],
            textposition='outside',
            showlegend=False
        ),
        row=1, col=1
    )

    youth_group_values = pct_list('church_youth_in_youth_group')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=youth_group_values, name="Youth Group %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in youth_group_values],
            textposition='outside',
            showlegend=False
        ),
        row=1, col=2
    )

    serving_values = pct_list('kids_youth_serving')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=serving_values, name="Serving %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in serving_values],
            textposition='outside',
            showlegend=False
        ),
        row=1, col=3
    )

    # Row 2: Conversions and Weekly Trends
    conversion_values = [conversions_count(y) for y in years]
    fig.add_trace(
        go.Bar(
            x=year_labels, y=conversion_values, name="Conversions",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[str(v) for v in conversion_values],
            textposition='outside',
            showlegend=False
        ),
        row=2, col=1
    )

    # Weekly attendance for Kids Club & Youth Group (current year)
    weekly_data = metrics.get('weekly_attendance', {})
    current_weekly = weekly_data.get('current', {})
    
    # Kids Club weekly
    kc_section = current_weekly.get('kids_club', {})
    if kc_section and kc_section.get('weekly_counts'):
        kc_weeks = list(range(1, len(kc_section['weekly_counts']) + 1))
        kc_counts = kc_section['weekly_counts']
        kc_smooth = calculate_rolling_average(kc_counts)
        
        fig.add_trace(
            go.Scatter(
                x=kc_weeks, y=kc_counts, mode='lines+markers',
                name='Kids Club', line=dict(color=colors['primary'], width=2),
                marker=dict(size=6), showlegend=True
            ),
            row=2, col=2
        )
        
        fig.add_trace(
            go.Scatter(
                x=kc_weeks, y=kc_smooth, mode='lines',
                name='Kids Club (trend)', line=dict(color=colors['primary'], width=3, dash='dot'),
                showlegend=True
            ),
            row=2, col=2
        )

    # Youth Group weekly  
    yg_section = current_weekly.get('youth_group', {})
    if yg_section and yg_section.get('weekly_counts'):
        yg_weeks = list(range(1, len(yg_section['weekly_counts']) + 1))
        yg_counts = yg_section['weekly_counts']
        yg_smooth = calculate_rolling_average(yg_counts)
        
        fig.add_trace(
            go.Scatter(
                x=yg_weeks, y=yg_counts, mode='lines+markers',
                name='Youth Group', line=dict(color=colors['secondary'], width=2),
                marker=dict(size=6), showlegend=True
            ),
            row=2, col=2
        )
        
        fig.add_trace(
            go.Scatter(
                x=yg_weeks, y=yg_smooth, mode='lines',
                name='Youth Group (trend)', line=dict(color=colors['secondary'], width=3, dash='dot'),
                showlegend=True
            ),
            row=2, col=2
        )

    # Buzz Playgroup weekly
    buzz_section = current_weekly.get('buzz', {})
    if buzz_section and buzz_section.get('weekly_counts'):
        buzz_weeks = list(range(1, len(buzz_section['weekly_counts']) + 1))
        buzz_counts = buzz_section['weekly_counts']
        buzz_smooth = calculate_rolling_average(buzz_counts)
        
        fig.add_trace(
            go.Scatter(
                x=buzz_weeks, y=buzz_counts, mode='lines+markers',
                name='Buzz Playgroup', line=dict(color=colors['accent'], width=2),
                marker=dict(size=6), showlegend=True
            ),
            row=2, col=3
        )
        
        fig.add_trace(
            go.Scatter(
                x=buzz_weeks, y=buzz_smooth, mode='lines',
                name='Buzz (trend)', line=dict(color=colors['accent'], width=3, dash='dot'),
                showlegend=True
            ),
            row=2, col=3
        )

    # Row 3: Counts and Summary
    kids_club_counts = count_list('church_kids_in_kids_club')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=kids_club_counts, name="Kids Club Count",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[str(v) for v in kids_club_counts],
            textposition='outside',
            showlegend=False
        ),
        row=3, col=1
    )

    youth_group_counts = count_list('church_youth_in_youth_group')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=youth_group_counts, name="Youth Group Count",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[str(v) for v in youth_group_counts],
            textposition='outside',
            showlegend=False
        ),
        row=3, col=2
    )

    # Summary table
    current_year_idx = years[-1]
    summary_data = []
    
    # Get current year metrics
    kc_metric = metrics.get('church_kids_in_kids_club', {}).get(current_year_idx, {})
    yg_metric = metrics.get('church_youth_in_youth_group', {}).get(current_year_idx, {})
    serving_metric = metrics.get('kids_youth_serving', {}).get(current_year_idx, {})
    
    summary_data.extend([
        ["Kids Club Participation", f"{kc_metric.get('count', 0)}/{kc_metric.get('total', 0)}", f"{kc_metric.get('percentage', 0):.1f}%"],
        ["Youth Group Participation", f"{yg_metric.get('count', 0)}/{yg_metric.get('total', 0)}", f"{yg_metric.get('percentage', 0):.1f}%"],
        ["NextGen Serving", f"{serving_metric.get('count', 0)}/{serving_metric.get('total', 0)}", f"{serving_metric.get('percentage', 0):.1f}%"],
        ["Conversions This Year", str(conversions_count(current_year_idx)), ""],
        ["Conversions Last Year", str(conversions_count(years[-2])) if len(years) >= 2 else "0", ""]
    ])

    fig.add_trace(
        go.Table(
            header=dict(values=["Metric", "Count", "Percentage"], 
                       fill_color=colors['primary'], font=dict(color='white', size=12)),
            cells=dict(values=[[row[0] for row in summary_data],
                              [row[1] for row in summary_data], 
                              [row[2] for row in summary_data]],
                      fill_color='white', font=dict(color=colors['text'], size=11)),
            showlegend=False
        ),
        row=3, col=3
    )

    # Layout updates
    fig.update_layout(
        title={
            'text': f"🎯 NextGen Dashboard — Kids, Youth & Conversions Analytics",
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 24, 'color': colors['text']}
        },
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        font=dict(family="Arial, sans-serif", size=10, color=colors['text']),
        plot_bgcolor=colors['background'],
        paper_bgcolor=colors['background'],
        height=1200,
        width=1600
    )

    # Update axes
    for i in range(1, 4):
        for j in range(1, 4):
            if (i == 2 and j in [2, 3]) or (i == 3 and j == 3):  # Skip line charts and table
                continue
            fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor=colors['grid'], row=i, col=j)
            fig.update_xaxes(showgrid=False, row=i, col=j)

    # Save files
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'NextGen_Dashboard_{timestamp}.html'
    
    try:
        fig.write_html(filename, div_id="nextgen-dashboard", include_plotlyjs=True)
        print(f"✅ Dashboard saved as: {filename}")
        
        # Open in browser
        try:
            webbrowser.open(f'file://{os.path.abspath(filename)}')
        except Exception:
            pass
        
        # Also save as PNG
        try:
            png_filename = f'NextGen_Dashboard_{timestamp}.png'
            fig.write_image(png_filename, width=1600, height=1400, scale=2)
            print(f"✅ Also saved as PNG: {png_filename}")
        except Exception as e:
            print(f"⚠️ PNG save failed: {e}")
    except Exception as e:
        print(f"❌ Failed to save dashboard: {e}")
        return None

    return filename

def main():
    """Main execution function"""
    print("🚀 Starting NextGen Dashboard Analysis...")

    # Step 1: Get all people with demographics and custom fields
    people_data, professed_custom_fields = get_all_people_with_demographics()
    if not people_data:
        print("❌ Failed to fetch people data")
        return

    # Step 1b: Load People Categories (needed to resolve category_id -> name)
    categories_by_id = fetch_people_categories()

    # Step 2: Get conversions
    conversions = get_conversions(people_data, professed_custom_fields)

    # Step 3: Find report groups (3 group reports + 3 service reports)
    report_groups = find_attendance_report_groups()
    if not report_groups:
        print("❌ Could not find report groups")
        return

    # Step 4: Extract group weekly attendance data (Kids Club, Youth Group, Buzz/JKC/KC)
    group_weekly_data = {}
    for year_key in ['two_years_ago', 'last_year', 'current']:
        grp = report_groups['group_reports'].get(year_key)
        if grp:
            group_weekly_data[year_key] = extract_attendance_data_from_group(grp, year_key)

    # Step 5: Extract service attendance data (10:30 AM and 6:30 PM)
    service_attendance_data = {}
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    year_keys = ['two_years_ago', 'last_year', 'current']

    for i, yk in enumerate(year_keys):
        svc_grp = report_groups['service_reports'].get(yk)
        if svc_grp:
            service_attendance_data[yk] = parse_service_attendance(svc_grp, yk, years[i])

    # Step 6: Calculate metrics (⚠️ now 5 args, passes categories_by_id)
    metrics = calculate_metrics(
        group_weekly_data,
        service_attendance_data,
        people_data,
        conversions,
        categories_by_id
    )

    # Step 7: Create dashboard
    dashboard_file = create_dashboard(metrics)

    print("\n🎯 NextGen Dashboard Complete!")
    print(f"📊 Dashboard saved as: {dashboard_file}")
    print("\nKey Metrics Summary:")

    if years[-1] in metrics['church_kids_in_kids_club']:
        kc = metrics['church_kids_in_kids_club'][years[-1]]
        print(f"• Church Children in Kids Club: {kc.get('count',0)}/{kc.get('total',0)} = {kc.get('percentage',0):.1f}%")

    if years[-1] in metrics['church_youth_in_youth_group']:
        yg = metrics['church_youth_in_youth_group'][years[-1]]
        print(f"• Church Youth in Youth Group: {yg.get('count',0)}/{yg.get('total',0)} = {yg.get('percentage',0):.1f}%")

    if years[-1] in metrics['kids_youth_serving']:
        sv = metrics['kids_youth_serving'][years[-1]]
        print(f"• Kids & Youth Serving: {sv.get('count',0)}/{sv.get('total',0)} = {sv.get('percentage',0):.1f}%")

    conv_now  = len(metrics['conversions'].get(years[-1], []))
    conv_prev = len(metrics['conversions'].get(years[-2], [])) if len(years) >= 2 else 0
    print(f"• Conversions This Year: {conv_now}")
    print(f"• Conversions Last Year: {conv_prev}")

if __name__ == "__main__":
    main()
