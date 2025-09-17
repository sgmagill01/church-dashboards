import requests
import json
from datetime import datetime, timedelta
import webbrowser
import os
from bs4 import BeautifulSoup
import math
import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'requests']
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

# Get API key
print("🏛️ ST GEORGE'S CHURCH HEALTH DASHBOARD - WORD-BASED MINISTRY EDITION")
print("="*65)
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

# Word-based ministry volunteer positions (from WBM/UG classification)
WORD_BASED_POSITIONS = {
    'Acoustic Guitar', 'BSG Leader', 'Band Leader', 'Bass', 'Cajon', 'Comm. Celebrant',
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
            'fields': ['demographics', 'departments']  # API will automatically include deceased, archived, status fields
        })
        if not response or not response.get('people'):
            print("Done")
            break
        people = response['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []

        # Filter out deceased and other excluded people
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
    print("📋 Fetching groups with categories and people...")
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

# ===== WBM SUCCESSFUL FUNCTIONS INTEGRATED =====

def find_new_serving_members_report():
    """Find the 'Report of New Serving Members' group (from WBM)"""
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
    """Extract new serving members data from the report group (from WBM)"""
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

def extract_volunteer_positions(person):
    """Extract volunteer positions from person's departments (from WBM/UG)"""
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

def is_word_based_group(group_name):
    """Check if group qualifies for word-based ministry (from WBM/UG)"""
    if not group_name:
        return False
    group_name_lower = group_name.lower()
    return not ('cherry picking' in group_name_lower or 'community care' in group_name_lower)

def is_group_leader(person_id, groups):
    """Check if person is a leader of any word-based ministry group (from WBM/UG)"""
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

def get_detailed_person_info(person_id):
    """Get detailed person information including full volunteer positions (from WBM)"""
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

def analyze_new_serving_members_word_based(new_serving_members, all_people, groups):
    """Analyze which new serving members are in word-based ministry using Member ID matching (from WBM)"""
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

# ===== END WBM INTEGRATION =====

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

def get_date_professed(person):
    """Extract Date Professed from person's demographics"""
    if not person.get('demographics'):
        return None

    demographics = person['demographics']

    if isinstance(demographics, dict) and 'demographic' in demographics:
        demo_list = demographics['demographic']
        if not isinstance(demo_list, list):
            demo_list = [demo_list] if demo_list else []

        for demo in demo_list:
            if isinstance(demo, dict):
                demo_name = demo.get('name', '').lower()
                if 'date professed' in demo_name or 'professed' in demo_name:
                    date_value = demo.get('value', '')
                    if date_value:
                        try:
                            for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']:
                                try:
                                    return datetime.strptime(date_value, fmt)
                                except:
                                    continue
                        except:
                            pass
    return None

def calculate_word_based_ministry_ratio(members, groups):
    """Calculate the percentage of serving members in word-based ministry"""
    print("📖 Calculating word-based ministry participation...")

    word_based_members = set()
    word_based_details = []  # Track details for console output

    for person in members:
        person_id = person.get('id')
        person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}"
        is_word_based = False
        reasons = []  # Track why this person qualifies

        # Check volunteer positions
        if str(person.get('volunteer', '0')) == '1':
            positions = extract_volunteer_positions(person)
            word_based_positions = positions.intersection(WORD_BASED_POSITIONS)
            if word_based_positions:
                is_word_based = True
                reasons.append(f"Volunteer positions: {', '.join(word_based_positions)}")

        # Check group leadership
        if is_group_leader(person_id, groups):
            is_word_based = True
            reasons.append("Group leader")

        if is_word_based:
            word_based_members.add(person_id)
            reason_str = " & ".join(reasons)
            word_based_details.append(f"  📖 {person_name}: {reason_str}")
            print(f"  - {person_name}: {reason_str}")

    word_based_count = len(word_based_members)
    total_members = len(members)
    ratio = (word_based_count / total_members) * 100 if total_members > 0 else 0

    print(f"  - {word_based_count} out of {total_members} serving members in word-based ministry")
    print(f"  - Word-based ministry ratio: {ratio:.1f}%")

    # Print complete list for console
    print(f"\n📖 Complete list of people in word-based ministry ({word_based_count} people):")
    for detail in sorted(word_based_details):
        print(detail)

    return ratio, word_based_count

def calculate_metrics(people, categories, groups):
    print("🔄 Calculating metrics...")

    # ===== NEW: WBM INTEGRATION FOR ACCURATE WORD-BASED MINISTRY RECRUITMENT =====
    
    # Find and extract new serving members using WBM approach
    print("\n🎯 USING WBM APPROACH FOR WORD-BASED MINISTRY RECRUITMENT...")
    new_serving_report = find_new_serving_members_report()
    word_based_recruited_this_year = 0
    word_based_recruitment_details = []
    
    if new_serving_report:
        new_serving_members = extract_new_serving_members_data(new_serving_report)
        if new_serving_members:
            word_based_recruited_this_year, word_based_recruitment_details = analyze_new_serving_members_word_based(
                new_serving_members, people, groups
            )
        else:
            print("⚠️ No new serving members found in report")
    else:
        print("⚠️ Could not find New Serving Members report - word-based recruitment count will be 0")
    
    # ===== END WBM INTEGRATION =====

    # Create category lookup
    category_lookup = {}
    for cat in categories:
        category_lookup[cat.get('id', '')] = cat.get('name', '')

    # Process people
    processed_people = []
    professed_in_past_year = []

    for person in people:
        category_id = person.get('category_id', '')
        category_name = category_lookup.get(category_id, '')
        person_type = categorize_person(category_name)

        try:
            date_added = datetime.strptime(person.get('date_added', ''), '%Y-%m-%d %H:%M:%S')
        except:
            date_added = datetime.now()

        # Check for Date Professed
        date_professed = get_date_professed(person)

        processed_people.append({
            **person,
            'category_name': category_name,
            'type': person_type,
            'date_added': date_added,
            'date_professed': date_professed
        })

        # Count those who professed in the past year
        if date_professed:
            one_year_ago = datetime.now() - timedelta(days=366)
            if date_professed >= one_year_ago:
                professed_in_past_year.append(person)

    print(f"Found {len(professed_in_past_year)} people who professed in the past year")

    # Filter by type
    all_people = [p for p in processed_people if p['type'] != 'excluded']
    serving_members = [p for p in processed_people if p['type'] == 'serving_member']
    congregation = [p for p in processed_people if p['type'] in ['serving_member', 'congregation_only']]

    # Use congregation for calculations (this maintains existing logic)
    members = congregation

    if len(serving_members) == 0:
        print("No 'RosteredMember_' categories found.")
    if len(congregation) == 0:
        print("No congregation members found. Using all people as baseline.")
        congregation = [p for p in processed_people if p['type'] != 'excluded']
        members = congregation

    # Date calculations
    now = datetime.now()
    one_year_ago = now - timedelta(days=366)
    two_years_ago = now - timedelta(days=732)

    # Current year stats
    added_in_past_year = [p for p in all_people if p['date_added'] >= one_year_ago]
    new_members_in_past_year = [p for p in added_in_past_year if p['type'] in ['serving_member', 'congregation_only']]

    # Previous year stats for comparison
    added_in_previous_year = [p for p in all_people if two_years_ago <= p['date_added'] < one_year_ago]
    new_members_in_previous_year = [p for p in added_in_previous_year if p['type'] in ['serving_member', 'congregation_only']]

    # Calculate ratios
    visitor_ratio = (len(added_in_past_year) / len(members)) * 100 if members else 0
    stay_ratio = (len(new_members_in_past_year) / len(added_in_past_year)) * 100 if added_in_past_year else 0

    # Calculate serve stats (replacing volunteer terminology)
    serving = [m for m in members if str(m.get('volunteer', '0')) == '1']
    serve_ratio = (len(serving) / len(members)) * 100 if members else 0

    # Count serving members who were added in the past year
    new_serving_past_year = len([v for v in serving if v['date_added'] >= one_year_ago])

    # Calculate group stats
    print(f"Analyzing {len(groups)} groups for Bible Study membership...")
    bible_study_groups = [g for g in groups if is_bible_study_group(g)]
    print(f"Found {len(bible_study_groups)} Bible Study groups (category: Bible Study Groups_)")

    # Create lookup for person ID to name
    person_lookup = {}
    for p in members:
        person_id = p.get('id')
        name = f"{p.get('firstname', '')} {p.get('lastname', '')}".strip()
        person_lookup[person_id] = name

    members_in_bible_study_set = set()
    member_ids = {p.get('id') for p in members}
    all_bible_study_members = []  # To collect all names

    for group in bible_study_groups:
        group_name = group.get('name', 'Unknown')
        people_count = 0
        group_member_names = []

        if group.get('people') and group['people'].get('person'):
            group_members = group['people']['person']
            if not isinstance(group_members, list):
                group_members = [group_members] if group_members else []

            for person in group_members:
                person_id = person.get('id')
                if person_id in member_ids:
                    members_in_bible_study_set.add(person_id)
                    people_count += 1
                    person_name = person_lookup.get(person_id, 'Unknown Person')
                    group_member_names.append(person_name)
                    all_bible_study_members.append(person_name)

        if people_count > 0:
            print(f"   - {group_name}: {people_count} members")
            for name in sorted(group_member_names):
                print(f"     • {name}")

    members_in_bible_study = len(members_in_bible_study_set)
    group_ratio = (members_in_bible_study / len(members)) * 100 if members else 0

    print(f"\nTotal unique members in Bible Study groups: {members_in_bible_study}")
    print("All Bible Study Group Members:")
    for name in sorted(set(all_bible_study_members)):  # Remove duplicates and sort
        print(f"  • {name}")

    # Calculate word-based ministry ratio (replaces regularity)
    word_ministry_ratio, word_ministry_count = calculate_word_based_ministry_ratio(serving_members, groups)

    # Compare to baseline from strategic plan (March 2025: 46 people in word-based ministry)
    # Source: Draft plan 2025-2029 v6.docx, "Using Gifts" section measurement
    march_2025_word_ministry_count = 46
    word_ministry_change = word_ministry_count - march_2025_word_ministry_count

    # Calculate year-over-year changes (12-month comparisons for visitor and stay metrics)
    people_added_change = len(added_in_past_year) - len(added_in_previous_year)
    new_members_change = len(new_members_in_past_year) - len(new_members_in_previous_year)

    # For group members added - we can calculate people in groups who were added in past year
    group_members_added = len([p for p in all_people 
                              if p.get('id') in members_in_bible_study_set 
                              and p['date_added'] >= one_year_ago])

    # Target calculation for strategic plan (using WBM data!)
    word_based_baseline = 46
    word_based_target = math.ceil(word_based_baseline * 1.10)  # 46 * 1.10 = 50.6 → 51 people
    word_based_needed_this_year = word_based_target - word_based_baseline  # 51 - 46 = 5 people
    # Calculate recruitment progress using WBM data
    word_based_recruitment_progress = (word_based_recruited_this_year / word_based_needed_this_year * 100) if word_based_needed_this_year > 0 else 0

    return {
        'total_people': len(all_people),
        'total_members': len(members),
        'serving_members': len(serving_members),
        'congregation': len(congregation),
        'professing_members': len(professed_in_past_year),
        'visitor_ratio': visitor_ratio,
        'stay_ratio': stay_ratio,
        'word_ministry_ratio': word_ministry_ratio,  # Replaces regularity_ratio
        'group_ratio': group_ratio,
        'serve_ratio': serve_ratio,  # Updated from volunteer_ratio
        'people_added': len(added_in_past_year),
        'added_members': len(new_members_in_past_year),
        'word_ministry_count': word_ministry_count,  # Replaces members_attended
        'members_in_groups': members_in_bible_study,
        'current_serving': len(serving),  # Updated from current_volunteers
        # Year-over-year changes
        'people_added_change': people_added_change,
        'new_members_change': new_members_change,
        'word_ministry_change': word_ministry_change,  # Replaces regularity_change
        'group_members_added': group_members_added,
        'new_serving': new_serving_past_year,  # Updated from new_volunteers
        # Strategic plan metrics (FIXED WITH WBM DATA!)
        'word_based_recruited_this_year': word_based_recruited_this_year,  # NOW ACCURATE!
        'word_based_recruitment_progress': word_based_recruitment_progress,  # NOW ACCURATE!
        'word_based_target': word_based_target,
        'word_based_needed_this_year': word_based_needed_this_year,
        'word_based_recruitment_details': word_based_recruitment_details  # Added for detailed output
    }

def generate_dashboard_html(metrics):
    # Metric definitions
    definitions = {
        'visitor': 'The number of people added to the database in the past year, expressed as a percentage of current members. This measures our ability to attract new people.',
        'stay': 'The percentage of people added in the past year who have become members. This measures our effectiveness at integrating newcomers into church life.',
        'word_ministry': 'The percentage of serving members who participate in word-based ministry (teaching, music, preaching, group leadership). This measures leadership development and spiritual engagement.',
        'group': 'The percentage of members who participate in a Bible Study group (includes kids club, youth group, and adult Bible studies). This measures depth of spiritual formation and community connection.',
        'serve': 'The percentage of members who serve in any capacity. This measures active participation in church ministry.'  # Updated from 'volunteer'
    }

    # Format change indicators
    # Note: visitor_ratio and stay_ratio use 12-month year-over-year comparisons
    # word_ministry now uses "this calendar year" instead of "March 2025 baseline"
    def format_change(value, metric_type="year_over_year", prefix="", suffix="", is_percentage=False):
        if value > 0:
            sign = "up"
            color = "#4ade80"
        elif value < 0:
            sign = "down"
            color = "#f87171"
            value = abs(value)
        else:
            if metric_type == "calendar_year":  # Updated from "march_baseline"
                return '<span style="color: #64748b; font-size: 12px;">unchanged this calendar year</span>'
            else:
                return '<span style="color: #64748b; font-size: 12px;">unchanged from previous year</span>'

        if metric_type == "calendar_year":  # Updated from "march_baseline"
            return f'<span style="color: {color}; font-size: 12px;">{sign} {int(value)} this calendar year</span>'
        elif is_percentage:
            return f'<span style="color: {color}; font-size: 12px;">{sign} {value:.1f}% from previous year</span>'
        else:
            return f'<span style="color: {color}; font-size: 12px;">{sign} {int(value)} from previous year</span>'

    def format_added(value, label):
        if label == 'serving members':  # Updated from 'volunteer members'
            return f'<span style="color: #4ade80; font-size: 12px;">{int(value)} people added to database (past year) who are serving members</span>'
        elif label == 'group members':
            return f'<span style="color: #4ade80; font-size: 12px;">{int(value)} people added to database (past year) who are in groups</span>'
        return f'<span style="color: #4ade80; font-size: 12px;">{int(value)} {label} added in the last year</span>'

    return f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Church Health Dashboard</title>
<style>
* {{margin:0;padding:0;box-sizing:border-box}}
body {{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0a0e27;min-height:100vh;padding:20px}}
.header {{display:flex;align-items:center;margin-bottom:30px;gap:20px}}
.logo {{width:70px;height:70px;background:linear-gradient(135deg,#6366f1,#8b5cf6);border-radius:20px;display:flex;align-items:center;justify-content:center;font-size:32px;color:white;font-weight:bold;box-shadow:0 10px 40px rgba(99,102,241,0.3)}}
.title-container {{background:rgba(255,255,255,0.05);padding:20px 30px;border-radius:20px;backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1)}}
.title {{font-size:28px;font-weight:700;color:#ffffff;margin:0;letter-spacing:-0.5px}}
.subtitle {{font-size:14px;color:#94a3b8;margin-top:5px}}
.stats-bar {{background:rgba(255,255,255,0.05);padding:25px 35px;border-radius:20px;display:flex;gap:50px;align-items:center;backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1);margin-bottom:35px}}
.stat-item {{text-align:center}}
.stat-label {{font-size:13px;color:#94a3b8;font-weight:500;margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px}}
.stat-value {{font-size:36px;font-weight:700;color:#ffffff}}
.dashboard {{display:grid;grid-template-columns:repeat(5,1fr);gap:25px;margin-bottom:35px}}
.metric-card {{background:rgba(255,255,255,0.05);border-radius:24px;padding:35px;text-align:center;backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1);position:relative;transition:all 0.3s ease}}
.metric-card:hover {{transform:translateY(-5px);border-color:rgba(255,255,255,0.2);background:rgba(255,255,255,0.08)}}
.metric-card::before {{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:24px 24px 0 0}}
.metric-card.visitor::before {{background:linear-gradient(90deg,#3b82f6,#60a5fa)}}
.metric-card.stay::before {{background:linear-gradient(90deg,#f97316,#fb923c)}}
.metric-card.word-ministry::before {{background:linear-gradient(90deg,#ec4899,#f472b6)}}
.metric-card.group::before {{background:linear-gradient(90deg,#eab308,#facc15)}}
.metric-card.serve::before {{background:linear-gradient(90deg,#22c55e,#4ade80)}}
.gauge-container {{position:relative;width:140px;height:140px;margin:0 auto 25px}}
.gauge {{width:140px;height:140px;border-radius:50%;background:#1e293b;position:relative;display:flex;align-items:center;justify-content:center;box-shadow:inset 0 0 20px rgba(0,0,0,0.3)}}
.gauge-ring {{position:absolute;width:140px;height:140px;border-radius:50%}}
.gauge-center {{width:100px;height:100px;background:#0a0e27;border-radius:50%;display:flex;align-items:center;justify-content:center;position:relative;z-index:2;box-shadow:0 5px 20px rgba(0,0,0,0.5)}}
.gauge-percentage {{font-size:32px;font-weight:700;color:#ffffff}}
.metric-title {{font-size:22px;font-weight:600;color:#ffffff;margin-bottom:15px;text-transform:capitalize}}
.metric-definition {{font-size:13px;color:#64748b;line-height:1.5;margin-bottom:20px;padding:0 10px;min-height:60px}}
.metric-description {{font-size:14px;color:#94a3b8;line-height:1.4;font-weight:500}}
.summary-cards {{display:grid;grid-template-columns:repeat(5,1fr);gap:25px}}
.summary-card {{background:rgba(255,255,255,0.05);border-radius:20px;padding:30px;text-align:center;backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1);transition:all 0.3s ease}}
.summary-card:hover {{transform:translateY(-3px);border-color:rgba(255,255,255,0.2)}}
.summary-title {{font-size:14px;color:#94a3b8;margin-bottom:12px;font-weight:500;text-transform:uppercase;letter-spacing:0.5px}}
.summary-value {{font-size:40px;font-weight:700;margin-bottom:8px}}
.summary-comparison {{margin-top:8px;height:16px}}
.summary-card.visitor .summary-value {{color:#60a5fa}}
.summary-card.stay .summary-value {{color:#fb923c}}
.summary-card.word-ministry .summary-value {{color:#f472b6}}
.summary-card.group .summary-value {{color:#facc15}}
.summary-card.serve .summary-value {{color:#4ade80}}
.footer-summary {{text-align:center;margin-top:50px;padding:25px;background:rgba(255,255,255,0.05);border-radius:20px;backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1)}}
.footer-summary h3 {{color:#ffffff;margin-bottom:15px;font-size:20px}}
.footer-summary p {{color:#94a3b8;font-size:15px;line-height:1.8}}
@media (max-width: 1200px) {{
    .dashboard {{ grid-template-columns: repeat(3, 1fr); }}
    .summary-cards {{ grid-template-columns: repeat(3, 1fr); }}
}}
</style></head><body>
<div class="header">
<div class="logo">📊</div>
<div class="title-container">
<h1 class="title">Church Health Analytics</h1>
<p class="subtitle">Generated {datetime.now().strftime("%B %d, %Y")} • Live Data Dashboard • Word-Based Ministry Edition</p>
</div></div>
<div class="stats-bar">
<div class="stat-item"><div class="stat-label">Serving Members</div><div class="stat-value">{metrics['serving_members']}</div></div>
<div class="stat-item"><div class="stat-label">Congregation</div><div class="stat-value">{metrics['congregation']}</div></div>
<div class="stat-item"><div class="stat-label">Professed This Year</div><div class="stat-value">{metrics['professing_members']}</div></div>
</div>
<div class="dashboard">
<div class="metric-card visitor">
<div class="gauge-container">
<div class="gauge">
<div class="gauge-ring" style="background:conic-gradient(from -90deg,#3b82f6 0deg,#3b82f6 {(metrics['visitor_ratio']/100)*360}deg,transparent {(metrics['visitor_ratio']/100)*360}deg)"></div>
<div class="gauge-center"><span class="gauge-percentage">{round(metrics['visitor_ratio'])}%</span></div>
</div></div>
<h3 class="metric-title">visitor ratio</h3>
<p class="metric-definition">{definitions['visitor']}</p>
<p class="metric-description">People Added in past year as % of current Members</p>
</div>
<div class="metric-card stay">
<div class="gauge-container">
<div class="gauge">
<div class="gauge-ring" style="background:conic-gradient(from -90deg,#f97316 0deg,#f97316 {(metrics['stay_ratio']/100)*360}deg,transparent {(metrics['stay_ratio']/100)*360}deg)"></div>
<div class="gauge-center"><span class="gauge-percentage">{round(metrics['stay_ratio'])}%</span></div>
</div></div>
<h3 class="metric-title">stay ratio</h3>
<p class="metric-definition">{definitions['stay']}</p>
<p class="metric-description">% of People Added in past year who are now Members</p>
</div>
<div class="metric-card word-ministry">
<div class="gauge-container">
<div class="gauge">
<div class="gauge-ring" style="background:conic-gradient(from -90deg,#ec4899 0deg,#ec4899 {(metrics['word_ministry_ratio']/100)*360}deg,transparent {(metrics['word_ministry_ratio']/100)*360}deg)"></div>
<div class="gauge-center"><span class="gauge-percentage">{round(metrics['word_ministry_ratio'])}%</span></div>
</div></div>
<h3 class="metric-title">word-based ministry</h3>
<p class="metric-definition">{definitions['word_ministry']}</p>
<p class="metric-description">% of Serving Members in word-based ministry roles</p>
</div>
<div class="metric-card group">
<div class="gauge-container">
<div class="gauge">
<div class="gauge-ring" style="background:conic-gradient(from -90deg,#eab308 0deg,#eab308 {(metrics['group_ratio']/100)*360}deg,transparent {(metrics['group_ratio']/100)*360}deg)"></div>
<div class="gauge-center"><span class="gauge-percentage">{round(metrics['group_ratio'])}%</span></div>
</div></div>
<h3 class="metric-title">group ratio</h3>
<p class="metric-definition">{definitions['group']}</p>
<p class="metric-description">% of Members who are in a Bible Study Group</p>
</div>
<div class="metric-card serve">
<div class="gauge-container">
<div class="gauge">
<div class="gauge-ring" style="background:conic-gradient(from -90deg,#22c55e 0deg,#22c55e {(metrics['serve_ratio']/100)*360}deg,transparent {(metrics['serve_ratio']/100)*360}deg)"></div>
<div class="gauge-center"><span class="gauge-percentage">{round(metrics['serve_ratio'])}%</span></div>
</div></div>
<h3 class="metric-title">serve ratio</h3>
<p class="metric-definition">{definitions['serve']}</p>
<p class="metric-description">% of Members who are Serving Members</p>
</div>
</div>
<div class="summary-cards">
<div class="summary-card visitor">
<div class="summary-title">People Added</div>
<div class="summary-value">{metrics['people_added']}</div>
<div class="summary-comparison">{format_change(metrics['people_added_change'], "year_over_year")}</div>
</div>
<div class="summary-card stay">
<div class="summary-title">New Members</div>
<div class="summary-value">{metrics['added_members']}</div>
<div class="summary-comparison">{format_change(metrics['new_members_change'], "year_over_year")}</div>
</div>
<div class="summary-card word-ministry">
<div class="summary-title">In Word-Based Ministry</div>
<div class="summary-value">{metrics['word_ministry_count']}</div>
<div class="summary-comparison">{format_change(metrics['word_ministry_change'], "calendar_year")}</div>
</div>
<div class="summary-card group">
<div class="summary-title">In Bible Study</div>
<div class="summary-value">{metrics['members_in_groups']}</div>
<div class="summary-comparison">{format_added(metrics['group_members_added'], 'group members')}</div>
</div>
<div class="summary-card serve">
<div class="summary-title">Serving Members</div>
<div class="summary-value">{metrics['current_serving']}</div>
<div class="summary-comparison">{format_added(metrics['new_serving'], 'serving members')}</div>
</div>
</div>
<div class="footer-summary">
<h3>Church Health Summary</h3>
<p>
Total People: {metrics['total_people']} | Members: {metrics['total_members']} | Professed: {metrics['professing_members']}<br>
Visitor: {round(metrics['visitor_ratio'])}% | Stay: {round(metrics['stay_ratio'])}% | Word Ministry: {round(metrics['word_ministry_ratio'])}% | Bible Study: {round(metrics['group_ratio'])}% | Serve: {round(metrics['serve_ratio'])}%<br>
<strong>Strategic Progress (WBM Method):</strong> {metrics.get('word_based_recruited_this_year', 0)} new word-based ministry recruits this year ({metrics.get('word_based_recruitment_progress', 0):.1f}% of annual target)
</p></div></body></html>'''

def main():
    try:
        people = fetch_all_people()
        categories = fetch_categories()
        groups = fetch_groups()

        if not people:
            print("Failed to fetch data. Check your API key.")
            return

        metrics = calculate_metrics(people, categories, groups)

        print("\n🎨 Generating dashboard...")
        dashboard_html = generate_dashboard_html(metrics)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"church_dashboard_word_ministry_fixed_{timestamp}.html"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(dashboard_html)

        print(f"✅ Dashboard saved: {filename}")
        print(f"📊 People: {metrics['total_people']}, Members: {metrics['total_members']}")
        print(f"🙏 Professed past year: {metrics['professing_members']}")
        print(f"📈 Visitor: {metrics['visitor_ratio']:.1f}% | Stay: {metrics['stay_ratio']:.1f}%")
        print(f"📖 Word Ministry: {metrics['word_ministry_ratio']:.1f}% | Groups: {metrics['group_ratio']:.1f}%")
        print(f"🤝 Serve: {metrics['serve_ratio']:.1f}%")
        print("\n📊 Year-over-year and baseline comparisons:")
        print(f"  - People added: {metrics['people_added_change']:+d} (12-month comparison)")
        print(f"  - New members: {metrics['new_members_change']:+d} (12-month comparison)")
        print(f"  - Word ministry: {metrics['word_ministry_change']:+d} this calendar year")
        print(f"  - Group members (people added to database in past year who are in groups): {metrics['group_members_added']}")
        print(f"  - New serving members (people added to database in past year who are serving members): {metrics['new_serving']}")
        print(f"\n🎯 Strategic Plan Progress (FIXED WITH WBM METHOD):")
        print(f"  - Word-based ministry recruits this year: {metrics['word_based_recruited_this_year']} (target: {metrics['word_based_needed_this_year']})")
        print(f"  - Progress toward annual target: {metrics['word_based_recruitment_progress']:.1f}%")
        
        # Print detailed recruitment list if available
        if metrics.get('word_based_recruitment_details'):
            print(f"\n📖 Detailed word-based ministry recruits this calendar year:")
            for detail in metrics['word_based_recruitment_details']:
                print(f"   • {detail['name']} (ID: {detail['member_id']}): {detail['reasons']} (from {detail['change_from']} on {detail['date_changed']})")

        file_path = os.path.abspath(filename)
        webbrowser.open(f"file://{file_path}")
        print("\n🌐 Dashboard opened!")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
