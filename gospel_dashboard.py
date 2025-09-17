#!/usr/bin/env python3
"""
Gospel Chart Dashboard
Uses Code GP pattern to access real attendance data from Elvanto reports
- Gospel Course Attendance: Real data from 'Ever Attended Taste and See' attendance reports
- Decisions: Real data from people with 'date professed' field
- Historical data: Current year and last year (2025 and 2024 currently)
"""

import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'plotly', 'requests', 'pandas']
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
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import re
from collections import defaultdict
from bs4 import BeautifulSoup
import webbrowser
import os

print("🏛️ ST GEORGE'S MAGILL - GOSPEL CHART DASHBOARD")
print("="*50)
print("📊 Real attendance data from Elvanto reports")
print()

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

def make_request(endpoint, params=None):
    """Make authenticated request to Elvanto API"""
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
                return None
        else:
            print(f"   HTTP Error {response.status_code}: {response.text[:200]}")
            return None
    except Exception as e:
        print(f"   Request failed: {e}")
        return None

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
        print(f"   ✅ Found {len(professed_candidates)} potential 'Date Professed' fields:")
        for candidate in professed_candidates:
            print(f"      🎯 '{candidate['name']}' (ID: {candidate['id']}, Type: {candidate['type']})")
    else:
        print(f"   ❌ No obvious 'Date Professed' custom fields found")
        print(f"   💡 The field might have a different name - check the list above")
    
    return professed_candidates

def get_date_professed(person, professed_custom_fields=None):
    """Extract Date Professed from person's custom fields (it's a custom datepicker field)"""
    person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
    
    # Date Professed is a custom field - search custom fields first
    if professed_custom_fields:
        for custom_field in professed_custom_fields:
            custom_field_name = f"custom_{custom_field['id']}"
            custom_field_value = person.get(custom_field_name)
            
            if custom_field_value:
                print(f"   🔍 Found '{custom_field['name']}' for {person_name}: '{custom_field_value}' (Type: {custom_field['type']})")
                
                # Handle datepicker fields specifically
                if custom_field['type'] == 'datepicker':
                    parsed_date = parse_datepicker_field(custom_field_value)
                else:
                    parsed_date = parse_date_robust(custom_field_value)
                
                if parsed_date:
                    print(f"   ✅ Parsed date for {person_name}: {parsed_date.strftime('%d/%m/%Y')}")
                    return parsed_date
                else:
                    print(f"   ❌ Could not parse date for {person_name}: '{custom_field_value}'")
    
    # Fallback: try demographics (though it's unlikely to be there)
    if person.get('demographics'):
        demographics = person['demographics']

        if isinstance(demographics, dict) and 'demographic' in demographics:
            demo_list = demographics['demographic']
            if not isinstance(demo_list, list):
                demo_list = [demo_list] if demo_list else []

            for demo in demo_list:
                if isinstance(demo, dict):
                    demo_name = demo.get('name', '').lower().strip()
                    
                    if any(keyword in demo_name for keyword in ['professed', 'profession', 'decision', 'conversion']):
                        date_value = demo.get('value', '').strip()
                        
                        if date_value:
                            print(f"   🔍 Found professed field (demographics) for {person_name}: '{demo_name}' = '{date_value}'")
                            
                            parsed_date = parse_date_robust(date_value)
                            if parsed_date:
                                print(f"   ✅ Parsed date for {person_name}: {parsed_date.strftime('%d/%m/%Y')}")
                                return parsed_date
                            else:
                                print(f"   ❌ Could not parse date for {person_name}: '{date_value}'")
    
    return None

def parse_datepicker_field(date_value):
    """Parse datepicker field value - might be in ISO format or other formats"""
    if not date_value:
        return None
    
    date_string = str(date_value).strip()
    
    # Datepicker fields often return ISO format dates
    iso_formats = [
        '%Y-%m-%d',           # 2024-03-31
        '%Y-%m-%dT%H:%M:%S',  # 2024-03-31T00:00:00
        '%Y-%m-%d %H:%M:%S',  # 2024-03-31 00:00:00
    ]
    
    # Try ISO formats first
    for fmt in iso_formats:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    
    # If ISO parsing fails, fall back to robust parsing
    return parse_date_robust(date_string)
    """Extract Date Professed from person's demographics with robust parsing"""
    person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
    
    if not person.get('demographics'):
        return None

    demographics = person['demographics']

    if isinstance(demographics, dict) and 'demographic' in demographics:
        demo_list = demographics['demographic']
        if not isinstance(demo_list, list):
            demo_list = [demo_list] if demo_list else []

        for demo in demo_list:
            if isinstance(demo, dict):
                demo_name = demo.get('name', '').lower().strip()
                
                # More flexible field name matching
                if any(keyword in demo_name for keyword in ['professed', 'profession', 'decision', 'conversion']):
                    date_value = demo.get('value', '').strip()
                    
                    if date_value:
                        print(f"   🔍 Found professed field for {person_name}: '{demo_name}' = '{date_value}'")
                        
                        # Robust date parsing with multiple formats and separators
                        parsed_date = parse_date_robust(date_value)
                        if parsed_date:
                            print(f"   ✅ Parsed date for {person_name}: {parsed_date.strftime('%d/%m/%Y')}")
                            return parsed_date
                        else:
                            print(f"   ❌ Could not parse date for {person_name}: '{date_value}'")
    
    return None

def parse_date_robust(date_string):
    """Robust date parsing with multiple formats and error handling"""
    if not date_string:
        return None
    
    # Clean the date string
    date_string = date_string.strip()
    
    # Try multiple date formats with different separators
    formats_to_try = [
        # DD/MM/YYYY formats
        '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d %m %Y',
        # MM/DD/YYYY formats  
        '%m/%d/%Y', '%m-%d-%Y', '%m.%d.%Y', '%m %d %Y',
        # YYYY-MM-DD formats
        '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y %m %d',
        # DD-MM-YYYY formats
        '%d-%m-%Y', '%d/%m/%Y',
        # With abbreviated month names
        '%d %b %Y', '%d %B %Y', '%b %d %Y', '%B %d %Y',
        # Alternative formats
        '%d-%b-%Y', '%d/%b/%Y', '%b-%d-%Y', '%b/%d/%Y'
    ]
    
    for fmt in formats_to_try:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    
    # If standard parsing fails, try to extract numbers and guess format
    try:
        # Look for patterns like "31/03/2024", "2024-03-31", etc.
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

def fetch_all_people():
    """Fetch ALL people from Elvanto using people/getAll with custom fields"""
    print("👥 Fetching ALL people from Elvanto using people/getAll...")
    
    # First, discover custom fields that might contain professed dates
    professed_custom_fields = fetch_custom_fields()
    
    # Build fields array including custom fields
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
        print(f"   People page {page}...", end=" ")
        response = make_request('people/getAll', {
            'page': page,
            'page_size': 1000,
            'fields': fields_to_request
        })
        
        if not response or not response.get('people'):
            print("Done")
            break
        
        people = response['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []
        
        # Include ALL people - no filtering
        all_people.extend(people)
        print(f"({len(people)} people)")
        
        if len(people) < 1000:  # Break when fewer than page_size returned
            break
        page += 1
    
    print(f"   Total people: {len(all_people)}")
    
    # Debug: Show what we're getting, including custom fields
    if all_people:
        sample_person = all_people[0]
        available_fields = list(sample_person.keys())
        print(f"   📊 Available fields: {available_fields}")
        
        # Check for custom fields in the response
        custom_fields_in_response = [field for field in available_fields if field.startswith('custom_')]
        if custom_fields_in_response:
            print(f"   🎯 Custom fields found in response: {len(custom_fields_in_response)}")
            for custom_field in custom_fields_in_response:
                print(f"      • {custom_field}")
        else:
            print(f"   ❌ No custom fields found in response")
        
        # Check demographics and custom field data availability
        demographics_count = 0
        custom_field_data_count = 0
        sample_custom_data = []
        
        for person in all_people[:20]:  # Check first 20 people
            if person.get('demographics'):
                demographics_count += 1
            
            # Check for any custom field data
            for field_name in available_fields:
                if field_name.startswith('custom_') and person.get(field_name):
                    custom_field_data_count += 1
                    if len(sample_custom_data) < 3:
                        person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}"
                        field_value = person.get(field_name)
                        sample_custom_data.append(f"{person_name}: {field_name} = {field_value}")
                    break
        
        print(f"   📊 People with demographics: {demographics_count}/20")
        print(f"   📊 People with custom field data: {custom_field_data_count}/20")
        
        if sample_custom_data:
            print(f"   📊 Sample custom field data:")
            for sample in sample_custom_data:
                print(f"      {sample}")
    
    return all_people, professed_custom_fields

def find_taste_and_see_attendance_reports():
    """Find generic attendance report groups (Code GP pattern)"""
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
    """Download attendance data from a report group (Code GP pattern)"""
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

def parse_attendance_data(html_content, year_label):
    """Parse attendance data from HTML report - only for 'Ever Attended Taste and See' group"""
    if not html_content:
        return {}
    
    print(f"📊 Parsing attendance data for {year_label}...")
    
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    
    if not table:
        print("   ❌ No table found in report")
        return {}
    
    rows = table.find_all('tr')
    if not rows:
        print("   ❌ No rows found in table")
        return {}
    
    # Find the "Ever Attended Taste and See" section
    print("   🔍 Looking for 'Ever Attended Taste and See' section...")
    
    target_group_found = False
    target_group_start_row = None
    target_group_end_row = None
    
    for i, row in enumerate(rows):
        # Check if this is a group header row (black background, white text)
        row_text = row.get_text().strip()
        
        # Look for "Ever Attended Taste and See" header
        if 'ever attended taste and see' in row_text.lower():
            target_group_found = True
            target_group_start_row = i + 1  # Start parsing from next row
            print(f"   ✅ Found 'Ever Attended Taste and See' at row {i}")
            print(f"   📝 Group header text: '{row_text}'")
            continue
        
        # If we found our target group and hit another group header, stop
        if target_group_found and target_group_start_row is not None:
            # Check if this looks like another group header (has minimal cells, different styling)
            cells = row.find_all(['td', 'th'])
            if len(cells) <= 2 and row_text and not any(char.isdigit() for char in row_text):
                # This might be the start of another group
                target_group_end_row = i
                print(f"   🛑 Found next group at row {i}: '{row_text}'")
                break
    
    if not target_group_found:
        print("   ❌ 'Ever Attended Taste and See' section not found in report")
        print("   📋 Available sections found:")
        for i, row in enumerate(rows):
            row_text = row.get_text().strip()
            cells = row.find_all(['td', 'th'])
            if len(cells) <= 2 and row_text and 'attended' not in row_text.lower():
                print(f"      • Row {i}: '{row_text}'")
        return {}
    
    # Parse only the rows belonging to "Ever Attended Taste and See"
    end_row = target_group_end_row if target_group_end_row else len(rows)
    target_rows = rows[target_group_start_row:end_row]
    
    print(f"   📊 Parsing rows {target_group_start_row} to {end_row-1} ({len(target_rows)} rows)")
    
    # Find headers within the target section (should be first row of the section)
    if not target_rows:
        print("   ❌ No data rows found in target section")
        return {}
    
    # First row should be headers
    header_row = target_rows[0]
    headers = [cell.get_text().strip() for cell in header_row.find_all(['th', 'td'])]
    print(f"   📋 Target section headers: {headers[:5]}...")
    
    # Find the "Attended" column
    attended_column_index = 1  # Default to column 1
    for i, header in enumerate(headers):
        if 'attended' in header.lower():
            attended_column_index = i
            print(f"   📍 Found 'Attended' column at index {i}")
            break
    
    # Parse data rows (skip the header row)
    attendees = set()
    data_rows = target_rows[1:]  # Skip header
    
    print(f"   🔍 Processing {len(data_rows)} data rows in target section...")
    
    for row_idx, row in enumerate(data_rows):
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        
        row_data = [cell.get_text().strip() for cell in cells]
        
        # Need at least name and attended columns
        if len(row_data) <= attended_column_index:
            continue
            
        person_name = row_data[0].strip()
        attended_count = row_data[attended_column_index].strip()
        
        if person_name and attended_count:
            try:
                # Convert attended count to integer
                attended_num = int(attended_count)
                if attended_num > 0:
                    attendees.add(person_name)
                    print(f"   ✅ {person_name}: attended {attended_num} times")
                else:
                    print(f"   ➖ {person_name}: attended {attended_num} times (not counted)")
            except ValueError:
                # If can't convert to int, skip this row
                print(f"   ⚠️ Could not parse attendance for {person_name}: '{attended_count}'")
                continue
    
    print(f"   🎯 RESULT: {len(attendees)} unique people attended in {year_label} (Ever Attended Taste and See only)")
    
    if attendees:
        print("   👥 People who attended:")
        for name in sorted(list(attendees)):
            print(f"      • {name}")
    else:
        print("   👥 No attendees found in 'Ever Attended Taste and See' section")
    
    return {'attendees': attendees, 'count': len(attendees)}

def analyze_decisions_by_year(people, current_year, last_year, two_years_ago, professed_custom_fields=None):
    """Analyze people who professed faith in current, last year, and two years ago"""
    print(f"📊 Analyzing decisions (professed faith) for {two_years_ago}, {last_year} and {current_year}...")
    
    decisions_by_year = {}
    professed_details = []
    people_checked = 0
    people_with_demographics = 0
    professed_fields_found = 0
    all_professed_dates = []  # Track ALL professed dates found
    
    for person in people:
        people_checked += 1
        
        if person.get('demographics'):
            people_with_demographics += 1
        
        date_professed = get_date_professed(person, professed_custom_fields)
        if date_professed:
            professed_fields_found += 1
            person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
            
            # Record ALL professed dates for debugging
            all_professed_dates.append({
                'name': person_name,
                'date': date_professed,
                'year': date_professed.year
            })
            
            year = date_professed.year
            if year in [current_year, last_year, two_years_ago]:
                if year not in decisions_by_year:
                    decisions_by_year[year] = 0
                decisions_by_year[year] += 1
                professed_details.append({
                    'name': person_name,
                    'date': date_professed,
                    'year': year
                })
    
    print(f"   📊 Processing summary:")
    print(f"      • Total people checked: {people_checked}")
    print(f"      • People with demographics data: {people_with_demographics}")
    print(f"      • Professed fields found: {professed_fields_found}")
    print(f"      • Matching target years: {len(professed_details)}")
    
    # Show ALL professed dates found (for debugging)
    if all_professed_dates:
        print(f"   🔍 ALL professed dates found ({len(all_professed_dates)} total):")
        for detail in sorted(all_professed_dates, key=lambda x: x['date']):
            in_target = "✅" if detail['year'] in [current_year, last_year, two_years_ago] else "➖"
            print(f"      {in_target} {detail['year']}: {detail['name']} ({detail['date'].strftime('%d/%m/%Y')})")
    else:
        print(f"   ❌ NO professed dates found in entire database!")
        print(f"      This suggests the search/parsing is not working correctly.")
    
    if professed_details:
        print(f"   🎯 People matching target years ({two_years_ago}, {last_year}, {current_year}):")
        for detail in sorted(professed_details, key=lambda x: x['date']):
            print(f"      ✅ {detail['year']}: {detail['name']} ({detail['date'].strftime('%d/%m/%Y')})")
    else:
        print(f"   ❌ No professed faith records found in target years")
        if professed_fields_found > 0:
            print(f"      (But found {professed_fields_found} professed dates in other years - see above)")
    
    return decisions_by_year

def get_user_projections(current_year):
    """Allow user to customize projections for future years"""
    
    # Default projections from Staff Retreat Report
    default_projections = {
        current_year + 1: {'gospel': 21, 'decisions': 3},
        current_year + 2: {'gospel': 30, 'decisions': 6}, 
        current_year + 3: {'gospel': 34, 'decisions': 7},
        current_year + 4: {'gospel': 37, 'decisions': 7}
    }
    
    # Show current projections
    print("\n📊 PROJECTION SETTINGS")
    print("="*50)
    print("Current projections (from Staff Retreat Report):")
    print()
    print("📈 Gospel Course Attendance:")
    for year in range(current_year + 1, current_year + 5):
        print(f"   {year}: {default_projections[year]['gospel']} people")
    
    print()
    print("📈 Decisions (Professed Faith):")
    for year in range(current_year + 1, current_year + 5):
        print(f"   {year}: {default_projections[year]['decisions']} people")
    
    print()
    
    # Ask if user wants to customize
    while True:
        choice = input("Do you want to customize these projections? (y/n): ").strip().lower()
        if choice in ['y', 'yes']:
            break
        elif choice in ['n', 'no']:
            print("✅ Using default projections from Staff Retreat Report")
            return default_projections
        else:
            print("Please enter 'y' for yes or 'n' for no")
    
    # Get custom projections
    print("\n🎯 CUSTOM PROJECTIONS")
    print("="*30)
    print("Enter your custom projections for the next 4 years:")
    print()
    
    custom_projections = {}
    
    # Get Gospel Course Attendance projections
    print("📈 Gospel Course Attendance:")
    for i, year in enumerate(range(current_year + 1, current_year + 5), 1):
        while True:
            try:
                value = input(f"   {year} (Year {i}): ").strip()
                gospel_count = int(value)
                if gospel_count < 0:
                    print("   Please enter a non-negative number")
                    continue
                custom_projections[year] = {'gospel': gospel_count}
                break
            except ValueError:
                print("   Please enter a valid number")
    
    print()
    
    # Get Decisions projections  
    print("📈 Decisions (Professed Faith):")
    for i, year in enumerate(range(current_year + 1, current_year + 5), 1):
        while True:
            try:
                value = input(f"   {year} (Year {i}): ").strip()
                decisions_count = int(value)
                if decisions_count < 0:
                    print("   Please enter a non-negative number")
                    continue
                custom_projections[year]['decisions'] = decisions_count
                break
            except ValueError:
                print("   Please enter a valid number")
    
    # Confirm custom projections
    print()
    print("✅ Your custom projections:")
    print("📈 Gospel Course Attendance:")
    for year in range(current_year + 1, current_year + 5):
        print(f"   {year}: {custom_projections[year]['gospel']} people")
    
    print("📈 Decisions (Professed Faith):")
    for year in range(current_year + 1, current_year + 5):
        print(f"   {year}: {custom_projections[year]['decisions']} people")
    
    print()
    
    return custom_projections

def create_combined_data(decisions_real, attendance_real, current_year, last_year, two_years_ago, custom_projections=None):
    """Combine real historical data with projections (default or custom)"""
    print("📊 Combining real data with projections...")
    
    # Use custom projections if provided, otherwise use default Staff Retreat Report projections
    if custom_projections:
        projections = custom_projections
        print("   Using custom user projections")
    else:
        # Default projections from Staff Retreat Report
        projections = {
            current_year + 1: {'gospel': 21, 'decisions': 3},
            current_year + 2: {'gospel': 30, 'decisions': 6},
            current_year + 3: {'gospel': 34, 'decisions': 7},
            current_year + 4: {'gospel': 37, 'decisions': 7}
        }
        print("   Using default Staff Retreat Report projections")
    
    # Combine historical and projection data (now includes three years of historical data)
    all_years = list(range(two_years_ago, current_year + 5))  # Extended to include projection years
    
    gospel_data = {}
    decisions_data = {}
    
    for year in all_years:
        if year <= current_year:
            # Use real data for historical years
            gospel_data[year] = attendance_real.get(year, 0)
            decisions_data[year] = decisions_real.get(year, 0)
        else:
            # Use projections for future years
            if year in projections:
                gospel_data[year] = projections[year]['gospel']
                decisions_data[year] = projections[year]['decisions']
            else:
                gospel_data[year] = 0
                decisions_data[year] = 0
    
    return gospel_data, decisions_data
    """Combine real historical data with projections (default or custom)"""
    print("📊 Combining real data with projections...")
    
    # Use custom projections if provided, otherwise use default Staff Retreat Report projections
    if custom_projections:
        projections = custom_projections
        print("   Using custom user projections")
    else:
        # Default projections from Staff Retreat Report
        projections = {
            current_year + 1: {'gospel': 21, 'decisions': 3},
            current_year + 2: {'gospel': 30, 'decisions': 6},
            current_year + 3: {'gospel': 34, 'decisions': 7},
            current_year + 4: {'gospel': 37, 'decisions': 7}
        }
        print("   Using default Staff Retreat Report projections")
    
    # Combine historical and projection data (now includes three years of historical data)
    all_years = list(range(two_years_ago, current_year + 5))  # Extended to include projection years
    
    gospel_data = {}
    decisions_data = {}
    
    for year in all_years:
        if year <= current_year:
            # Use real data for historical years
            gospel_data[year] = attendance_real.get(year, 0)
            decisions_data[year] = decisions_real.get(year, 0)
        else:
            # Use projections for future years
            if year in projections:
                gospel_data[year] = projections[year]['gospel']
                decisions_data[year] = projections[year]['decisions']
            else:
                gospel_data[year] = 0
                decisions_data[year] = 0
    
    return gospel_data, decisions_data

def create_chart(gospel_data, decisions_data, current_year):
    """Create the Gospel Chart with exact Staff Retreat Report styling"""
    print("🎨 Creating Gospel Chart...")
    
    # Exact colors from the original chart
    BLUE_COLOR = '#4A90E2'      # Historical data
    ORANGE_COLOR = '#F5A623'    # Projected data
    
    # Create subplot figure (2 charts side by side)
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=['Gospel course attendance', 'Decisions'],
        horizontal_spacing=0.15
    )
    
    years = list(gospel_data.keys())
    
    # Determine colors for each year (blue for historical, orange for projected)
    colors_gospel = []
    colors_decisions = []
    
    for year in years:
        if year <= current_year:  # Historical data
            colors_gospel.append(BLUE_COLOR)
            colors_decisions.append(BLUE_COLOR)
        else:  # Projected data
            colors_gospel.append(ORANGE_COLOR)
            colors_decisions.append(ORANGE_COLOR)
    
    # Add Gospel course attendance chart (left panel)
    fig.add_trace(
        go.Bar(
            x=years,
            y=list(gospel_data.values()),
            name='Gospel Course',
            marker_color=colors_gospel,
            showlegend=False
        ),
        row=1, col=1
    )
    
    # Add Decisions chart (right panel)
    fig.add_trace(
        go.Bar(
            x=years,
            y=list(decisions_data.values()),
            name='Decisions',
            marker_color=colors_decisions,
            showlegend=False
        ),
        row=1, col=2
    )
    
    # Update layout with simplified title
    fig.update_layout(
        title={
            'text': '<b>Gospel Chart</b>',
            'x': 0.5,
            'y': 0.95,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': {'size': 24, 'family': 'Arial', 'color': 'black'}
        },
        width=1000,
        height=500,
        plot_bgcolor='white',
        paper_bgcolor='white',
        font={'family': 'Arial', 'size': 12, 'color': 'black'},
        margin={'t': 80, 'b': 80, 'l': 80, 'r': 80}
    )
    
    # Update x-axes
    fig.update_xaxes(
        showgrid=False,
        showline=True,
        linewidth=1,
        linecolor='lightgray',
        tickfont={'size': 11}
    )
    
    # Update y-axes with appropriate ranges
    max_gospel = max(gospel_data.values()) if gospel_data.values() else 40
    max_decisions = max(decisions_data.values()) if decisions_data.values() else 8
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='lightgray',
        showline=True,
        linewidth=1,
        linecolor='lightgray',
        tickfont={'size': 11},
        range=[0, max(45, max_gospel + 5)],  # Gospel course range
        row=1, col=1
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='lightgray',
        showline=True,
        linewidth=1,
        linecolor='lightgray',
        tickfont={'size': 11},
        range=[0, max(9, max_decisions + 1)],   # Decisions range
        row=1, col=2
    )
    
    # Update subplot titles
    fig.layout.annotations[0].update(font={'size': 14, 'color': 'black'})
    fig.layout.annotations[1].update(font={'size': 14, 'color': 'black'})
    
    return fig

def main():
    """Main execution function"""
    try:
        current_year = datetime.now().year
        last_year = current_year - 1
        two_years_ago = current_year - 2
        
        print(f"🚀 Starting Gospel Chart Creation for {two_years_ago}, {last_year} and {current_year}...")
        print()
        
        # Step 1: Get all people for decisions analysis
        all_people, professed_custom_fields = fetch_all_people()
        if not all_people:
            print("❌ Failed to fetch people data")
            return
        
        # Step 2: Find the generic attendance reports (Code GP pattern)
        current_year_report, last_year_report, two_years_ago_report = find_taste_and_see_attendance_reports()
        
        # Step 3: Download and parse attendance data from all available reports
        attendance_by_year = {}
        
        if current_year_report:
            current_data = download_attendance_report_data(current_year_report)
            current_parsed = parse_attendance_data(current_data, f"{current_year} data")
            if current_parsed:
                attendance_by_year[current_year] = current_parsed['count']
        
        if last_year_report:
            last_data = download_attendance_report_data(last_year_report)
            last_parsed = parse_attendance_data(last_data, f"{last_year} data")
            if last_parsed:
                attendance_by_year[last_year] = last_parsed['count']
                
        if two_years_ago_report:
            two_years_ago_data = download_attendance_report_data(two_years_ago_report)
            two_years_ago_parsed = parse_attendance_data(two_years_ago_data, f"{two_years_ago} data")
            if two_years_ago_parsed:
                attendance_by_year[two_years_ago] = two_years_ago_parsed['count']
        
        print()
        
        # Step 4: Analyze decisions data
        decisions_by_year = analyze_decisions_by_year(all_people, current_year, last_year, two_years_ago, professed_custom_fields)
        
        print()
        print("📊 Real Data Summary:")
        print(f"   Gospel Course Attendance:")
        for year in [two_years_ago, last_year, current_year]:
            count = attendance_by_year.get(year, 0)
            print(f"      {year}: {count} people")
        
        print(f"   Decisions (people who professed):")
        for year in [two_years_ago, last_year, current_year]:
            count = decisions_by_year.get(year, 0)
            print(f"      {year}: {count} people")
        
        # Step 5: Get user projections (after data analysis is complete)
        user_projections = get_user_projections(current_year)
        
        # Step 6: Combine with projections
        gospel_data, decisions_data = create_combined_data(
            decisions_by_year, attendance_by_year, current_year, last_year, two_years_ago, user_projections
        )
        
        print()
        print("📊 Combined Data (Real + Projections):")
        print("   Gospel Course Attendance:")
        for year, count in gospel_data.items():
            status = "Real" if year <= current_year else "Projected"
            print(f"      {year}: {count:2d} ({status})")
        
        print("   Decisions:")
        for year, count in decisions_data.items():
            status = "Real" if year <= current_year else "Projected"
            print(f"      {year}: {count:2d} ({status})")
        print()
        
        # Step 7: Create and save chart
        fig = create_chart(gospel_data, decisions_data, current_year)
        
        # Save as HTML file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"gospel_chart_{timestamp}.html"
        
        fig.write_html(filename)
        print(f"✅ Chart saved as: {filename}")
        
        # Automatically open the chart in browser
        try:
            file_path = os.path.abspath(filename)
            webbrowser.open(f'file://{file_path}')
            print(f"🌐 Chart opened in your default browser")
        except Exception as e:
            print(f"⚠️  Could not auto-open chart: {e}")
            print(f"   You can manually open: {filename}")
        
        # Save as PNG for high-quality output
        try:
            png_filename = f"gospel_chart_{timestamp}.png"
            fig.write_image(png_filename, width=1000, height=500, scale=2)
            print(f"✅ High-quality PNG saved as: {png_filename}")
        except Exception as e:
            print(f"⚠️  PNG export failed (install kaleido for PNG support): {e}")
        
        print()
        print("🎯 SUCCESS! Gospel Chart created with real Elvanto attendance reports")
        print("📈 Features:")
        print(f"   • Real attendance data from generic attendance reports")
        print(f"   • Parsed 'Ever Attended Taste and See' section specifically")
        print(f"   • Real decisions data from 'date professed' field")
        print(f"   • Historical data: {two_years_ago}, {last_year} and {current_year}")
        print(f"   • Projections: 2025-2029")
        print("   • Blue bars for historical, orange for projected")
        
    except KeyboardInterrupt:
        print("\nProcess cancelled by user")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
