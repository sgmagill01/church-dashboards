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

    # bucket we’ll fill
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
        # numeric year fallback
        m = re.search(r'(20\d{2})', nm)
        if m:
            yr = int(m.group(1))
            if yr == yr_map['two_years_ago']:
                return 'two_years_ago'
            if yr == yr_map['last_year']:
                return 'last_year'
            if yr == yr_map['current']:
                return 'current'
        # default
        return 'current'

    # walk all groups once
    for g in groups:
        if not isinstance(g, dict):
            continue
        name = g.get('name', '')
        nm = norm(name)

        # Must be a "report" and "individual attendance"
        if 'report' not in nm or 'individual attendance' not in nm:
            continue

        is_service = 'service' in nm
        is_group   = 'group'   in nm

        if not (is_service or is_group):
            continue  # not one we care about

        period = detect_period(nm)
        bucket = 'service_reports' if is_service else 'group_reports'

        # Only assign if empty (first good match wins)
        if report_groups[bucket][period] is None:
            report_groups[bucket][period] = g
            nice = "service" if is_service else "group"
            if period == 'two_years_ago':
                print(f"✅ Found two years ago {nice} report: {g.get('name')}")
            elif period == 'last_year':
                print(f"✅ Found last year {nice} report: {g.get('name')}")
            else:
                print(f"✅ Found current year {nice} report: {g.get('name')}")

    # Quick visibility if anything is missing
    for kind in ['group_reports', 'service_reports']:
        for period in ['two_years_ago', 'last_year', 'current']:
            if report_groups[kind][period] is None:
                print(f"   ⚠️ Missing {period.replace('_',' ')} in {kind.replace('_',' ')}")

    return report_groups

def get_group_members_by_name(group_name):
    """Get all members of a specific group by name using the API (robust to list/dict variants)."""
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

    # --- COERCE to a list of person dicts, regardless of shape ---
    people_list = []
    if isinstance(people_node, dict):
        # typical: {'person': [...] } or {'person': {...}}
        persons = people_node.get('person', [])
        if isinstance(persons, list):
            people_list = persons
        elif persons:
            people_list = [persons]
    elif isinstance(people_node, list):
        # already a list of person dicts
        people_list = people_node
    else:
        people_list = []

    print(f"   👤 Roster size for '{group_name}': {len(people_list)}")

    # Build lookup: name variants -> person info
    name_lookup = {}
    for person in people_list:
        if not isinstance(person, dict):
            continue
        person_id = person.get('id') or person.get('person_id')  # be generous
        first_name = (person.get('firstname') or person.get('first_name') or '').strip()
        last_name  = (person.get('lastname')  or person.get('last_name')  or '').strip()
        if person_id and (first_name or last_name):
            full_name = f"{first_name} {last_name}".strip()
            key1 = full_name.lower()
            key2 = f"{last_name}, {first_name}".lower() if first_name and last_name else ''
            info = {'id': person_id, 'firstname': first_name, 'lastname': last_name, 'full_name': full_name}
            name_lookup[key1] = info
            if key2:
                name_lookup[key2] = info

    return name_lookup

def normalize_name_from_html(html_name):
    """Normalize a name from HTML attendance report"""
    if not html_name:
        return ""
    
    # Strip role suffixes
    role_suffixes = ['(leader)', '(assistant leader)', '(helper)', '(coordinator)']
    normalized = html_name.strip()
    
    for suffix in role_suffixes:
        if normalized.lower().endswith(suffix):
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

# === NAME/ID LOOKUPS (used across multiple sections) ===
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
    if not display_name:
        return None
    n = normalize_name_from_html(display_name)  # strips '(Leader)' etc., lowercases

    # direct hit
    if n in global_lookup:
        return global_lookup[n]

    # flip "last, first" <-> "first last"
    if ',' in n:
        last, first = [x.strip() for x in n.split(',', 1)]
        return global_lookup.get(f"{first} {last}".lower())
    else:
        parts = n.split()
        if len(parts) >= 2:
            first = parts[0]
            last  = ' '.join(parts[1:])
            return global_lookup.get(f"{last}, {first}".lower())
    return None

def fetch_people_categories():
    """Return dict: {category_id: category_name} for all People Categories."""
    resp = make_request('people/categories/getAll', {})
    cat_map = {}
    if not resp:
        print("   ❌ Failed to fetch people categories")
        return cat_map

    cats = resp.get('categories', {}).get('category', [])
    if not isinstance(cats, list):
        cats = [cats] if cats else []

    for c in cats:
        if isinstance(c, dict):
            cid = c.get('id')
            name = (c.get('name') or '').strip()
            if cid and name:
                cat_map[cid] = name
    print(f"   📋 Loaded {len(cat_map)} People Categories")
    return cat_map

def fetch_people_categories_map():
    """
    Returns {category_id: category_name} for all People Categories.
    """
    categories_by_id = {}
    page = 1
    while True:
        resp = make_request('people/categories/getAll', {
            'page': page,
            'page_size': 1000
        })
        if not resp:
            break
        cats = resp.get('categories', {}).get('category', [])
        if not isinstance(cats, list):
            cats = [cats] if cats else []
        for c in cats:
            cid = c.get('id')
            name = c.get('name', '')
            if cid and name is not None:
                categories_by_id[cid] = name
        if len(cats) < 1000:
            break
        page += 1
    return categories_by_id

def person_serves_via_people_category(person, categories_by_id):
    """
    Returns True if the person's People Category name indicates they are rostered.
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
    A person counts as serving if ANY of the following is true:
      • Their People Category (via category_id) name == 'RosteredMember_' (case/spacing-insensitive)
      • They belong to at least one Department
      • (Fallback) a demographic entry uses 'RosteredMember_' as value or name contains it
    Returns: set(person_id)
    """
    import re

    def _norm(s):
        return re.sub(r'[^a-z0-9]', '', str(s or '').lower())

    target_norm = _norm("RosteredMember_")
    rostered_ids = set()

    picked_by_category = []
    picked_by_department = []
    picked_by_demo = []

    for p in (people_data or []):
        pid = p.get('id')
        if not pid:
            continue

        # A) People Category via category_id -> name
        cat_id = p.get('category_id')
        if cat_id and cat_id in categories_by_id:
            cat_name = categories_by_id[cat_id]
            if _norm(cat_name) == target_norm:
                rostered_ids.add(pid)
                picked_by_category.append(f"{p.get('firstname','')} {p.get('lastname','')}".strip())
                continue  # no need to double count

        # B) Any Department assignment
        depts = p.get('departments')
        dept_hit = False
        if isinstance(depts, dict) and 'department' in depts:
            dlist = depts['department']
            if not isinstance(dlist, list):
                dlist = [dlist]
            dept_hit = len(dlist) > 0
        elif isinstance(depts, list):
            dept_hit = len(depts) > 0
        elif isinstance(depts, str):
            dept_hit = bool(depts.strip())

        if dept_hit:
            rostered_ids.add(pid)
            picked_by_department.append(f"{p.get('firstname','')} {p.get('lastname','')}".strip())
            continue

        # C) Fallback: demographics encoding
        demos = p.get('demographics') or {}
        demo_list = demos.get('demographic')
        if demo_list:
            if isinstance(demo_list, dict):
                demo_list = [demo_list]
            for d in demo_list:
                if not isinstance(d, dict):
                    continue
                name = (d.get('name') or '')
                val  = (d.get('value') or '')
                if _norm(val) == target_norm or target_norm in _norm(name):
                    rostered_ids.add(pid)
                    picked_by_demo.append(f"{p.get('firstname','')} {p.get('lastname','')}".strip())
                    break

    # Debug summary
    print(f"   🧾 Rostered by People Category: {len(picked_by_category)}")
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

def church_children_and_youth_from_services(service_year_data, global_lookup, kids_ids_all, youth_ids_all):
    """
    From one year's service data (output of parse_service_attendance),
    return (children_ids_in_church, youth_ids_in_church) where "in church"
    = attended a service ≥2 times in that year.

    Accepts:
      - 'people_overall' (preferred; already filtered to ≥2 in parse_service_attendance)
      - Fallback union of 'people_10_30' and 'people_6_30' if needed.
    """
    names = set()
    if isinstance(service_year_data, dict):
        if service_year_data.get('people_overall'):
            names |= set(service_year_data['people_overall'])
        else:
            if service_year_data.get('people_10_30'):
                names |= set(service_year_data['people_10_30'])
            if service_year_data.get('people_6_30'):
                names |= set(service_year_data['people_6_30'])

    attendee_ids = set()
    for n in names:
        pid = name_to_id(n, global_lookup)
        if pid:
            attendee_ids.add(pid)

    church_children_ids = attendee_ids & kids_ids_all
    church_youth_ids    = attendee_ids & youth_ids_all
    return church_children_ids, church_youth_ids

def kids_church_children_from_groups(group_year_data, global_lookup, kids_ids_all):
    """
    Children considered 'in church' for SERVING denominator:
    those who attended Kids Church (older Sunday program) ≥2 in this year.
    (JKC is NOT included here.)
    """
    ids = set()
    kc = group_year_data.get('kids_church')
    if kc:
        # prefer matched_people (already person IDs)
        if kc.get('matched_people'):
            ids |= set(kc['matched_people'].keys())
        else:
            # fallback: map names -> IDs
            for name, cnt in kc.get('people', {}).items():
                if cnt >= 2:
                    pid = name_to_id(name, global_lookup)
                    if pid:
                        ids.add(pid)

    # safety: only keep those who are actually in the Children demographic
    return ids & kids_ids_all

def extract_attendance_data_from_group(group, year_key):
    """Extract attendance data from 'Report of Group Individual Attendance' pages.

    Returns for each detected section (kids_club, youth_group, buzz, kids_church, junior_kids_church):
      {
        'people': { display_name -> count>=2 },
        'matched_people': { person_id: {'name': full_name, 'attendance_count': n} },
        'group_name': '<section title in report>',
        'weekly_counts': [int, ...],      # Y's per (valid) date column
        'date_labels':  [str, ...],       # matching headers (dd/mm[/yy])
      }
    """
    print(f"\n🔍 Extracting {year_key} group attendance data...")

    if not group:
        print(f"❌ No {year_key} group provided")
        return {}

    # Locate the public report URL on the group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break
    if not report_url:
        print(f"❌ No URL found in {year_key} group")
        return {}

    # Map year_key -> base year for dd/mm headers with no year
    now_year = datetime.now().year
    base_year = {
        'two_years_ago': now_year - 2,
        'last_year':     now_year - 1,
        'current':       now_year
    }.get(year_key, now_year)

    def classify_section(s: str):
        s = (s or '').strip().lower()
        if any(k in s for k in ['junior kids church', 'junior kids', 'jkc']):
            return 'junior_kids_church'
        if 'kids church (with youth)' in s or 'youth church' in s or 'kids church' in s:
            return 'kids_church'
        if 'buzz' in s:
            return 'buzz'
        if 'kids club' in s:
            return 'kids_club'
        if 'youth group' in s:
            return 'youth_group'
        return None

    YES_VALUES = {'Y', 'YES', '✓', '✔', '1', 'TRUE'}
    date_pat = re.compile(r'^\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?\s*$')

    try:
        html = requests.get(report_url, timeout=30).text
        soup = BeautifulSoup(html, 'html.parser')

        attendance_data = {}
        tables = soup.find_all('table')
        if not tables:
            print("   ⚠️ No tables found in report HTML")
            return {}

        for table in tables:
            rows = table.find_all('tr')
            if not rows:
                continue

            # --- header detection ---
            header_cells = rows[0].find_all(['th', 'td'])
            headers = [c.get_text(strip=True) for c in header_cells]

            name_idx = attended_idx = absent_idx = -1
            for i, h in enumerate(headers):
                hl = (h or '').strip().lower()
                if name_idx == -1 and 'name' in hl:
                    name_idx = i
                elif attended_idx == -1 and 'attended' in hl:
                    attended_idx = i
                elif absent_idx == -1 and 'absent' in hl:
                    absent_idx = i

            # Strictly collect only true dd/mm[/yy|yyyy] columns
            date_indices, date_labels, date_values = [], [], []
            for i, h in enumerate(headers):
                m = date_pat.match(h or '')
                if not m:
                    continue
                d = int(m.group(1)); mth = int(m.group(2))
                yr_token = m.group(3)
                if yr_token:
                    y = int(yr_token)
                    if y < 100:  # normalize 2-digit years
                        y += 2000 if y < 50 else 1900
                else:
                    y = base_year
                try:
                    dte = datetime(y, mth, d).date()
                except ValueError:
                    continue
                date_indices.append(i)
                date_labels.append(h)
                date_values.append(dte)

            # Fallback: if none matched, try the classic "after Attended/Absent" region
            if not date_indices:
                start = (absent_idx + 1) if absent_idx >= 0 else (
                        (attended_idx + 1) if attended_idx >= 0 else (
                        (name_idx + 1) if name_idx >= 0 else 1))
                for i in range(start, len(headers)):
                    h = headers[i]
                    m = date_pat.match(h or '')
                    if not m:
                        continue
                    d = int(m.group(1)); mth = int(m.group(2))
                    yr_token = m.group(3)
                    y = (base_year if not yr_token else (int(yr_token) + (2000 if int(yr_token) < 50 else 0)))
                    try:
                        dte = datetime(y, mth, d).date()
                    except Exception:
                        continue
                    date_indices.append(i)
                    date_labels.append(h)
                    date_values.append(dte)

            if name_idx < 0 or not date_indices:
                continue  # not a data table we can use

            # We will build per-section stores lazily, sized to the date_indices
            # and later trim future/tailing-zero columns.
            current_section_key = None
            for r in rows[1:]:
                cells = r.find_all(['td', 'th'])
                if not cells:
                    continue

                # Section header row (single cell)
                if len(cells) == 1:
                    section_title = cells[0].get_text(strip=True)
                    section_key = classify_section(section_title)
                    if section_key:
                        current_section_key = section_key
                        if section_key not in attendance_data:
                            attendance_data[section_key] = {
                                'people': {},
                                'group_name': section_title,
                                'weekly_counts': [0] * len(date_indices),
                                'date_labels': list(date_labels),
                                '_date_values': list(date_values),  # internal; dropped before return
                            }
                        print(f"   🎯 Found section: {section_title} → {section_key}")
                    else:
                        current_section_key = None
                    continue

                if not current_section_key:
                    continue

                # Data row
                raw_name = cells[name_idx].get_text(strip=True) if name_idx < len(cells) else ''
                if not raw_name:
                    continue
                if raw_name.lower() in ['name & position', 'name', 'first name']:
                    continue

                person_yes = 0
                # Only read true date columns
                wcounts = attendance_data[current_section_key]['weekly_counts']
                for j, col_idx in enumerate(date_indices):
                    if col_idx < len(cells):
                        v = cells[col_idx].get_text(strip=True).upper()
                        if v in YES_VALUES:
                            person_yes += 1
                            if j < len(wcounts):
                                wcounts[j] += 1

                if person_yes >= 2:
                    attendance_data[current_section_key]['people'][raw_name] = person_yes

        # --- Trim future dates and trailing zero-only columns ---
        today = datetime.now().date()
        for skey, sdata in attendance_data.items():
            counts = sdata.get('weekly_counts', [])
            labels = sdata.get('date_labels', [])
            dvals  = sdata.get('_date_values', [])

            # Drop columns with dates in the future
            keep_idx = [i for i, dt in enumerate(dvals) if (dt is None or dt <= today)]

            # If there is at least one non-zero, drop trailing zeros after the last non-zero
            last_nonzero = -1
            for i in keep_idx:
                if i < len(counts) and counts[i] > 0:
                    last_nonzero = i
            if last_nonzero >= 0:
                keep_idx = [i for i in keep_idx if i <= last_nonzero]
            else:
                keep_idx = []  # nothing recorded this year

            sdata['weekly_counts'] = [counts[i] for i in keep_idx]
            sdata['date_labels']   = [labels[i] for i in keep_idx]
            sdata.pop('_date_values', None)
            print(f"   🧹 Trimmed to {len(sdata['weekly_counts'])} valid weeks in {skey}")

        # Debug summary before roster matching
        for skey, sdata in attendance_data.items():
            print(f"   🧾 {skey}: raw people ≥2 = {len(sdata.get('people', {}))}, weeks={len(sdata.get('weekly_counts', []))}")

        # --- Map names -> IDs via roster (as before) ---
        for skey, sdata in attendance_data.items():
            matched_people = {}
            try:
                group_name = sdata.get('group_name', '')
                name_lookup = get_group_members_by_name(group_name)
                if name_lookup:
                    for html_name, cnt in sdata.get('people', {}).items():
                        mp = match_html_name_to_person(html_name, name_lookup)
                        if mp:
                            matched_people[mp['id']] = {
                                'name': mp['full_name'],
                                'attendance_count': cnt
                            }
            except Exception as e:
                print(f"   ⚠️ Skipping roster match for '{sdata.get('group_name','?')}': {e}")
            finally:
                sdata['matched_people'] = matched_people

            print(f"   ✅ Matched {len(matched_people)}/{len(sdata.get('people', {}))} people in {skey}")

        print(f"✅ Extracted data for {len(attendance_data)} programs")
        return attendance_data

    except Exception as e:
        print(f"❌ Error extracting {year_key} data: {e}")
        return {}

def kids_church_children_from_groups(group_year_data, global_lookup, kids_ids_all):
    """
    Return the set of person_ids for *older Sunday program* Kids Church (NOT Junior Kids Church)
    who attended ≥2 times in this year, AND who are currently in the 'Children' demographic.
    This is used as the denominator for the Kids Club participation rate so under-5s are excluded.

    Notes:
    - Prefers IDs from 'matched_people'.
    - Falls back to mapping raw names via the *actual Elvanto group roster* for the section
      using get_group_members_by_name(section_title) + match_html_name_to_person(...).
    - 'global_lookup' is ignored if section roster lookup is available (kept for drop-in compatibility).
    """
    ids = set()
    if not isinstance(group_year_data, dict):
        return ids

    kc = group_year_data.get('kids_church')  # older Sunday program only
    if not kc:
        return ids

    # Case 1: we already have person IDs from matched_people (attendance >=2 was applied upstream)
    matched = kc.get('matched_people')
    if isinstance(matched, dict) and matched:
        ids |= {pid for pid, info in matched.items()
                if isinstance(info, dict) and int(info.get('attendance_count', 0)) >= 2}

    # Case 2: fall back to raw 'people' names -> resolve via section roster
    if not ids:
        people_map = kc.get('people') or {}
        if isinstance(people_map, dict) and people_map:
            section_title = kc.get('group_name') or 'Kids Church'
            # Pull the live roster for this specific section
            section_lookup = get_group_members_by_name(section_title) or {}
            for raw_name, cnt in people_map.items():
                try:
                    cnt_int = int(cnt)
                except Exception:
                    cnt_int = 0
                if cnt_int < 2:
                    continue
                mp = match_html_name_to_person(raw_name, section_lookup)
                if mp and mp.get('id'):
                    ids.add(mp['id'])

    # Safety: restrict to those who are currently marked 'Children'
    return ids & (kids_ids_all or set())

def get_demographic_category(person):
    """Get the person's demographic category"""
    if not person.get('demographics'):
        return None
    
    demographics = person['demographics']
    if isinstance(demographics, dict) and 'demographic' in demographics:
        demo_list = demographics['demographic']
        if not isinstance(demo_list, list):
            demo_list = [demo_list] if demo_list else []
        
        for demo in demo_list:
            if isinstance(demo, dict) and demo.get('name', '').lower() == 'category':
                return demo.get('value', '')
    
    return None

def parse_service_attendance(group, year_key, year):
    """Parse service attendance data with ≥2 attendance rule and proper name matching"""
    print(f"\n🔍 Parsing {year_key} service attendance data...")

    if not group:
        print(f"❌ No {year_key} service group provided")
        return {}

    # Extract URL from group
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field])
            break

    if not report_url:
        print(f"❌ No URL found in {year_key} service group")
        return {}

    try:
        response = requests.get(report_url, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')
        tables = soup.find_all('table')

        def _h(s): return (s or '').strip().lower()

        for table in tables:
            header_row = table.find('tr')
            if not header_row:
                continue

            headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]

            # --- detect columns (First/Last or single Name) and service columns ---
            first_idx = -1
            last_idx  = -1
            name_idx  = -1
            service_columns = []

            for i, header in enumerate(headers):
                h = _h(header)
                if ('10:30' in h and 'am' in h) or ('6:30' in h and 'pm' in h):
                    service_columns.append({
                        'index': i,
                        'header': header,
                        'time': '10:30 AM' if '10:30' in h else '6:30 PM'
                    })
                    continue
                if first_idx == -1 and any(k in h for k in ['first name','firstname','given name']):
                    first_idx = i; continue
                if last_idx == -1 and any(k in h for k in ['last name','lastname','surname','family name']):
                    last_idx = i; continue
                if name_idx == -1 and any(k in h for k in ['name','full name']):
                    name_idx = i; continue

            # Need a name source and at least one service column
            if not service_columns or (first_idx < 0 and name_idx < 0):
                continue

            # --- nested helper to compose full names per row ---
            def compose_full_name(cells):
                f = cells[first_idx].get_text(strip=True) if (first_idx >= 0 and first_idx < len(cells)) else ''
                l = cells[last_idx].get_text(strip=True)  if (last_idx  >= 0 and last_idx  < len(cells)) else ''
                if f or l:
                    return f"{f} {l}".strip()
                if name_idx >= 0 and name_idx < len(cells):
                    return cells[name_idx].get_text(strip=True)
                return ''

            # --- scan rows, count Y's per person per service time ---
            people_attendance = {}
            for row in table.find_all('tr')[1:]:  # Skip header
                cells = row.find_all(['td', 'th'])
                full_name = compose_full_name(cells)
                if not full_name:
                    continue
                if _h(full_name) in ['first name','last name','surname','family name','name','full name']:
                    continue
                norm_name = normalize_name_from_html(full_name)
                if norm_name not in people_attendance:
                    people_attendance[norm_name] = {'10:30 AM': 0, '6:30 PM': 0}
                for service_col in service_columns:
                    col_idx = service_col['index']
                    if col_idx < len(cells):
                        val = cells[col_idx].get_text(strip=True).upper()
                        if val == 'Y':
                            people_attendance[norm_name][service_col['time']] += 1

            # --- apply ≥2 rule and build result ---
            ppl_1030 = set()
            ppl_630  = set()
            ppl_overall = set()

            for person_name, counts in people_attendance.items():
                if counts['10:30 AM'] >= 2:
                    ppl_1030.add(person_name)
                    ppl_overall.add(person_name)
                if counts['6:30 PM'] >= 2:
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
            # Optional peek:
            # print(f"   🧪 Sample attendees: {list(result['people_overall'])[:10]}")
            return result

        print(f"   ❌ No suitable service attendance table found")
        return {}

    except Exception as e:
        print(f"❌ Error parsing {year_key} service data: {e}")
        return {}

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
    
    return professed_candidates

def parse_date_robust(date_string):
    """Robust date parsing for various formats"""
    if not date_string:
        return None
    
    date_string = str(date_string).strip()
    
    # Common formats to try
    formats_to_try = [
        '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d',
        '%d-%m-%Y', '%m-%d-%Y', '%Y-%m-%d',
        '%d.%m.%Y', '%m.%d.%Y', '%Y.%m.%d',
        '%d %m %Y', '%m %d %Y', '%Y %m %d',
        '%Y-%m-%d', '%d/%m/%y', '%m/%d/%y'
    ]
    
    for fmt in formats_to_try:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    
    # Try regex parsing for more flexibility
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

def get_date_professed(person, professed_custom_fields=None):
    """Extract Date Professed from person's custom fields (Code GD approach)"""
    person_name = f"{person.get('firstname', '')} {person.get('lastname', '')}".strip()
    
    # Date Professed is a custom field - search custom fields first
    if professed_custom_fields:
        for custom_field in professed_custom_fields:
            custom_field_name = f"custom_{custom_field['id']}"
            custom_field_value = person.get(custom_field_name)
            
            if custom_field_value:
                # Handle datepicker fields specifically
                if custom_field['type'] == 'datepicker':
                    parsed_date = parse_datepicker_field(custom_field_value)
                else:
                    parsed_date = parse_date_robust(custom_field_value)
                
                if parsed_date:
                    return parsed_date.date() if hasattr(parsed_date, 'date') else parsed_date
    
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
                            parsed_date = parse_date_robust(date_value)
                            if parsed_date:
                                return parsed_date.date() if hasattr(parsed_date, 'date') else parsed_date
    
    return None

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
        if not name_set:
            return out
        for disp in name_set:
            if not disp:
                continue
            disp_norm = normalize_name_from_html(disp)
            pid_candidates = global_name_map.get(disp_norm, set())
            # light fuzzy if no exact
            if not pid_candidates:
                toks = [t for t in re.split(r'[\s,]+', disp_norm) if t]
                for key, ids in global_name_map.items():
                    if len(toks) >= 2 and all(tok in key for tok in toks[:2]):
                        pid_candidates |= ids
            for pid in pid_candidates:
                if restrict_ids is None or pid in restrict_ids:
                    out.add(pid)
        return out

    # ---- Serving roster (People Category + Departments + fallback demo) ----
    rostered_ids = build_rostered_ids(people_data, categories_by_id)

    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    year_keys = ['two_years_ago', 'last_year', 'current']

    metrics = {
        'church_kids_in_kids_club': {},
        'church_youth_in_youth_group': {},
        'kids_youth_serving': {},
        'conversions': conversions,
        'weekly_attendance': group_weekly_data,
        'service_attendance': service_attendance_data
    }

    for i, year in enumerate(years):
        yk = year_keys[i]
        print(f"\n   📈 Calculating metrics for {year} ({yk})")

        group_data   = (group_weekly_data or {}).get(yk, {}) or {}
        service_data = (service_attendance_data or {}).get(yk, {}) or {}

        # --- Kids Club Participation ---
        kc_denominator_ids = kids_church_children_from_groups(
            group_year_data=group_data,
            global_lookup={},    # kept for signature compatibility
            kids_ids_all=children_ids_all
        )
        kc_attendee_ids = _ids_from_section(group_data.get('kids_club'))

        if kc_denominator_ids:
            kids_club_church_kids = kc_attendee_ids & kc_denominator_ids
            pct = (len(kids_club_church_kids) / len(kc_denominator_ids)) * 100
            metrics['church_kids_in_kids_club'][year] = {
                'count': len(kids_club_church_kids),
                'total': len(kc_denominator_ids),
                'percentage': pct
            }
            print(f"      🎯 Kids Club: {len(kids_club_church_kids)}/{len(kc_denominator_ids)} = {pct:.1f}%")
        else:
            metrics['church_kids_in_kids_club'][year] = {'count': 0, 'total': 0, 'percentage': 0.0}
            print(f"      🎯 Kids Club: 0/0 (no Kids Church denominator found)")

        # --- Youth Group Participation (services ≥2) ---
        youth_names_min2 = (service_data or {}).get('people_overall_min2') \
                           or (service_data or {}).get('people_overall') \
                           or set()
        youth_service_ids_min2 = _ids_from_service_names(youth_names_min2, restrict_ids=youth_ids_all)

        yg_attendee_ids = _ids_from_section(group_data.get('youth_group'))

        if youth_service_ids_min2:
            yg_church_youth = yg_attendee_ids & youth_service_ids_min2
            pct = (len(yg_church_youth) / len(youth_service_ids_min2)) * 100
            metrics['church_youth_in_youth_group'][year] = {
                'count': len(yg_church_youth),
                'total': len(youth_service_ids_min2),
                'percentage': pct
            }
            print(f"      🎯 Youth Group: {len(yg_church_youth)}/{len(youth_service_ids_min2)} = {pct:.1f}%")
        else:
            metrics['church_youth_in_youth_group'][year] = {'count': 0, 'total': 0, 'percentage': 0.0}
            print(f"      🎯 Youth Group: 0/0 (no service-attending youth found)")

        # --- Serving % (eligible = KC children ≥2 ∪ Youth services ≥6 if available, else ≥2) ---
        youth_names_min6 = (service_data or {}).get('people_overall_min6') or set()
        youth_service_ids_min6 = _ids_from_service_names(youth_names_min6, restrict_ids=youth_ids_all)
        youth_for_serving = youth_service_ids_min6 if youth_service_ids_min6 else youth_service_ids_min2

        eligible_nextgen_ids = set(kc_denominator_ids) | set(youth_for_serving)
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

def create_dashboard(metrics):
    """
    Build the NextGen dashboard.

    Changes in this version:
      • Bottom-left & bottom-mid counts are combined into one grouped bar at (row=3, col=1)
        with thinner bars and a clearer title.
      • Bottom-mid (row=3, col=2) shows Program Weekly Averages computed over VALID WEEKS ONLY:
          - weeks up to the last non-zero week (YTD window)
          - zero weeks inside that window are excluded from the mean
      • Bottom-right (row=3, col=3) keeps the Key Metrics Summary.
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

    # Helper to pull a (x, y) weekly series from 'weekly_attendance'
    def series_for(year_key, section_key):
        sec = (metrics.get('weekly_attendance', {}).get(year_key, {}) or {}).get(section_key)
        if isinstance(sec, dict):
            y = sec.get('weekly_counts') or []
            x = sec.get('date_labels') or list(range(1, len(y) + 1))
            return (x, y) if y else ([], [])
        elif isinstance(sec, list):
            return (list(range(1, len(sec) + 1)), sec)
        return ([], [])

    # ---- Robust YTD averages over valid weeks only ----
    def yearly_avg(section_key, year_key):
        _, y = series_for(year_key, section_key)
        if not y:
            return 0.0

        # Trim to last non-zero week (YTD window)
        last_nz = -1
        for i, v in enumerate(y):
            try:
                if float(v) > 0:
                    last_nz = i
            except Exception:
                continue
        if last_nz == -1:
            return 0.0

        window = y[:last_nz + 1]
        # Exclude zero weeks inside the window
        valid = [float(v) for v in window if (isinstance(v, (int, float)) and v > 0)]
        if not valid:
            return 0.0
        return float(np.round(np.mean(valid), 1))

    # --- Figure & subplots (row 3 specs/titles adjusted) ---
    fig = make_subplots(
        rows=3, cols=3,
        subplot_titles=[
            "Kids Club Participation — Church Children (%)",
            "Youth Group Participation — Church Youth (%)",
            "Serving Rate — Eligible NextGen (%)",
            "NextGen Conversions per Year",
            "Weekly Attendance — Kids Club & Youth Group",
            "Weekly Attendance — Buzz Playgroup (Tue)",
            "Church-attending participants of Kids Club & Youth Group — counts per year",
            "Program Weekly Averages (including leaders)",
            "Key Metrics Summary"
        ],
        specs=[
            [{"type": "bar"}, {"type": "bar"}, {"type": "bar"}],
            [{"type": "bar"}, {"type": "scatter"}, {"type": "scatter"}],
            [{"type": "bar"}, {"type": "table"}, {"type": "table"}]
        ],
        vertical_spacing=0.12,
        horizontal_spacing=0.08
    )

    # Row 1: Percentages
    kids_club_values = pct_list('church_kids_in_kids_club')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=kids_club_values, name="Kids Club %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in kids_club_values], textposition='outside', cliponaxis=False
        ),
        row=1, col=1
    )

    youth_group_values = pct_list('church_youth_in_youth_group')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=youth_group_values, name="Youth Group %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in youth_group_values], textposition='outside', cliponaxis=False
        ),
        row=1, col=2
    )

    serving_values = pct_list('kids_youth_serving')
    fig.add_trace(
        go.Bar(
            x=year_labels, y=serving_values, name="Serving %",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=[f"{v:.1f}%" for v in serving_values], textposition='outside', cliponaxis=False
        ),
        row=1, col=3
    )

    # Row 2: Conversions and weekly series
    conversion_values = [conversions_count(y) for y in years]
    fig.add_trace(
        go.Bar(
            x=year_labels, y=conversion_values, name="Conversions",
            marker_color=[colors['two_years_ago'], colors['last_year'], colors['current_year']],
            text=conversion_values, textposition='outside', cliponaxis=False
        ),
        row=2, col=1
    )

    # Current weekly lines
    x1, y1 = series_for('current', 'kids_club')
    x2, y2 = series_for('current', 'youth_group')
    if y1:
        fig.add_trace(
            go.Scatter(x=x1, y=y1, mode='lines+markers',
                       name="Kids Club", line=dict(color=colors['primary'], width=3)),
            row=2, col=2
        )
    if y2:
        fig.add_trace(
            go.Scatter(x=x2, y=y2, mode='lines+markers',
                       name="Youth Group", line=dict(color=colors['secondary'], width=3)),
            row=2, col=2
        )
    xb, yb = series_for('current', 'buzz')
    if yb:
        fig.add_trace(
            go.Scatter(x=xb, y=yb, mode='lines+markers',
                       name="Buzz Playgroup", line=dict(color=colors['accent'], width=3)),
            row=2, col=3
        )

    # Row 3, Col 1: Combined grouped bars (thinner)
    kids_counts  = count_list('church_kids_in_kids_club')
    youth_counts = count_list('church_youth_in_youth_group')

    fig.add_trace(
        go.Bar(
            x=year_labels, y=kids_counts, name="Kids Count",
            marker_color=colors['primary'],
            width=0.28, offsetgroup="kids", legendgroup="combined"
        ),
        row=3, col=1
    )
    fig.add_trace(
        go.Bar(
            x=year_labels, y=youth_counts, name="Youth Count",
            marker_color=colors['secondary'],
            width=0.28, offsetgroup="youth", legendgroup="combined"
        ),
        row=3, col=1
    )

    # Row 3, Col 2: Program Weekly Averages (valid weeks only)
    programs = [
        ("Kids Club",   "kids_club"),
        ("Youth Group", "youth_group"),
        ("Buzz",        "buzz")
    ]
    prog_names = [p[0] for p in programs]
    avg_two   = [yearly_avg(p[1], 'two_years_ago') for p in programs]
    avg_last  = [yearly_avg(p[1], 'last_year')     for p in programs]
    avg_curr  = [yearly_avg(p[1], 'current')       for p in programs]

    fig.add_trace(
        go.Table(
            header=dict(
                values=["Program", "Two Years Ago", "Last Year", "Current Year"],
                fill_color=colors['primary'],
                font=dict(color='white', size=12)
            ),
            cells=dict(
                values=[prog_names, avg_two, avg_last, avg_curr],
                fill_color='lightgrey',
                font=dict(size=11),
                format=[None, ".1f", ".1f", ".1f"]
            )
        ),
        row=3, col=2
    )

    # Row 3, Col 3: Key Metrics Summary (unchanged)
    summary_rows = [
        ["Kids Club Participation — Church Children (%)",
         f"{kids_club_values[2]:.1f}%", f"{kids_club_values[1]:.1f}%",
         f"{kids_club_values[2] - kids_club_values[1]:+.1f}%"],
        ["Youth Group Participation — Church Youth (%)",
         f"{youth_group_values[2]:.1f}%", f"{youth_group_values[1]:.1f}%",
         f"{youth_group_values[2] - youth_group_values[1]:+.1f}%"],
        ["Serving Rate — Eligible NextGen (%)",
         f"{serving_values[2]:.1f}%", f"{serving_values[1]:.1f}%",
         f"{serving_values[2] - serving_values[1]:+.1f}%"],
        ["Conversions (NextGen)",
         f"{conversion_values[2]}", f"{conversion_values[1]}",
         f"{conversion_values[2] - conversion_values[1]:+d}"],
    ]
    fig.add_trace(
        go.Table(
            header=dict(
                values=["Metric", "Current Year", "Last Year", "Change"],
                fill_color=colors['primary'],
                font=dict(color='white', size=12)
            ),
            cells=dict(
                values=[
                    [r[0] for r in summary_rows],
                    [r[1] for r in summary_rows],
                    [r[2] for r in summary_rows],
                    [r[3] for r in summary_rows],
                ],
                fill_color='lightgrey',
                font=dict(size=11)
            )
        ),
        row=3, col=3
    )

    # Layout / styling
    fig.update_layout(
        title=dict(
            text=(
                f"<b style='font-size:28px; color:{colors['text']}'>🎯 NextGen Dashboard"
                f" — Kids, Youth & Conversions Analytics</b>"
                f"<br><span style='font-size:16px; color:{colors['text_light']}'>"
                f"St George's Magill Anglican Church</span>"
                f"<br><span style='font-size:12px; color:{colors['text_light']}'>"
                f"Generated {datetime.now().strftime('%B %d, %Y')}</span>"
            ),
            x=0.5, y=0.98,
            font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif")
        ),
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=12),
        plot_bgcolor=colors['background'],
        paper_bgcolor=colors['background'],
        height=1400, width=1600, showlegend=True,
        margin=dict(l=80, r=80, t=150, b=80),

        # Grouped bars look
        barmode="group",
        bargap=0.25,
        bargroupgap=0.18
    )

    for row in range(1, 4):
        for col in range(1, 4):
            # skip tables
            if (row == 3 and col in (2, 3)):
                continue
            fig.update_xaxes(
                showgrid=True, gridwidth=1, gridcolor=colors['grid'],
                tickfont=dict(size=11, color=colors['text']),
                linecolor=colors['grid'], row=row, col=col
            )
            fig.update_yaxes(
                showgrid=True, gridwidth=1, gridcolor=colors['grid'],
                tickfont=dict(size=11, color=colors['text']),
                linecolor=colors['grid'], rangemode='tozero',
                row=row, col=col
            )

    # Save / open (unchanged)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'NextGen_Dashboard_{timestamp}.html'
    try:
        fig.write_html(filename)
        print(f"✅ Dashboard saved as: {filename}")
        try:
            webbrowser.open(f'file://{os.path.abspath(filename)}')
            print("✅ Dashboard opened successfully in browser")
        except Exception as e:
            print(f"⚠️ Could not auto-open dashboard: {e}")
            print(f"Please manually open: {filename}")
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
