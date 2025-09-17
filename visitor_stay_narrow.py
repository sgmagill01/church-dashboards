#!/usr/bin/env python3
"""
Visitor and Stay Dashboard (Code VS)
Fixed version with proper data extraction, styling, charts, and PNG export
"""

import requests
import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import re
from collections import defaultdict
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio

# Import API key
try:
    from config import ELVANTO_API_KEY
    API_KEY = ELVANTO_API_KEY
except ImportError:
    print("❌ Error: config.py not found!")
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

def find_visitor_and_category_reports():
    """Find the 6 required reports using Elvanto API"""
    print("\n📋 Searching for visitor and category change reports using Elvanto API...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        print("❌ Failed to get groups from Elvanto API")
        return None, None
    
    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []
    
    print(f"   Found {len(groups)} total groups in Elvanto")
    
    # Initialize report containers
    category_reports = {
        'current': None,
        'last_year': None, 
        'two_years_ago': None
    }
    visitor_reports = {
        'current': None,
        'last_year': None,
        'two_years_ago': None
    }
    
    # Look for reports by exact name matching
    for group in groups:
        group_name = group.get('name', '')
        
        # People Category Change Reports
        if group_name == 'Report of People Category Change':
            category_reports['current'] = group
            print(f"✅ Found current category report: {group_name}")
        elif group_name == 'Report of Last Year People Category Change':
            category_reports['last_year'] = group
            print(f"✅ Found last year category report: {group_name}")
        elif group_name == 'Report of Two Years Ago People Category Change':
            category_reports['two_years_ago'] = group
            print(f"✅ Found two years ago category report: {group_name}")
            
        # New Visitors Reports
        elif group_name == 'Report of New Visitors':
            visitor_reports['current'] = group
            print(f"✅ Found current visitor report: {group_name}")
        elif group_name == 'Report of Last Year New Visitors':
            visitor_reports['last_year'] = group
            print(f"✅ Found last year visitor report: {group_name}")
        elif group_name == 'Report of Two Years Ago New Visitors':
            visitor_reports['two_years_ago'] = group
            print(f"✅ Found two years ago visitor report: {group_name}")
    
    # Debug: Show which reports we found
    print(f"\n📊 Report Discovery Summary:")
    for report_type, reports in [('Category Change', category_reports), ('New Visitors', visitor_reports)]:
        print(f"   {report_type}:")
        for year_key, report in reports.items():
            status = "✅ Found" if report else "❌ Missing"
            print(f"      {year_key}: {status}")
    
    return category_reports, visitor_reports

def visitors_for_ratio(year, svc, raw_visitors, counts_map, baseline_counts_by_service, *, today=None):
    """
    Return the visitor count to use in the *ratio*.
    - For non-current years: return raw_visitors (no scaling).
    - For the current year: scale by (baseline_counts / current_counts) if available,
      otherwise by a months fallback (12 / completed_months).
    """
    current_year = datetime.now().year
    if year != current_year:
        return raw_visitors  # absolutely no gross-up for past years

    # current year → apply scaling
    baseline = (baseline_counts_by_service or {}).get(svc, 0)
    current  = (counts_map or {}).get(svc, 0)
    if baseline and current:
        factor = baseline / current
    else:
        today = today or datetime.now()
        full_months = max(1, today.month - 1)  # completed months this year
        factor = 12.0 / full_months
    return raw_visitors * factor


def download_report_data(group, report_type):
    """Download report data from group URL using pattern from project knowledge"""
    if not group:
        print(f"   ❌ No {report_type} report group provided")
        return None
        
    group_name = group.get('name', 'Unknown')
    print(f"\n   📥 Downloading {report_type}: {group_name}")
    
    # Extract URL from group location fields (following Code GP pattern)
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if group.get(field) and 'http' in str(group[field]):
            report_url = str(group[field]).strip()
            print(f"      📍 Found report URL in {field}: {report_url[:80]}...")
            break
    
    if not report_url:
        print(f"      ❌ No download URL found in group fields")
        print(f"      🔍 Available fields: {list(group.keys())}")
        return None
    
    try:
        print(f"      📡 Downloading report data...")
        response = requests.get(report_url, timeout=60)
        if response.status_code == 200:
            print(f"      ✅ Downloaded {len(response.content)} bytes")
            return response.content
        else:
            print(f"      ❌ HTTP Error {response.status_code}")
            return None
    except Exception as e:
        print(f"      ❌ Download failed: {e}")
        return None

def parse_new_visitors_report(html_content, year_label):
    """Parse New Visitors report HTML to extract visitor data (faster, flexible)."""
    if not html_content:
        return []

    print(f"   📊 Parsing New Visitors data for {year_label}...")

    try:
        soup = BeautifulSoup(html_content, 'lxml')
    except Exception:
        soup = BeautifulSoup(html_content, 'html.parser')

    table = soup.find('table')
    if not table:
        print("      ❌ No table found in report")
        return []

    # One-pass text matrix
    rows = table.find_all('tr')
    if len(rows) < 2:
        print("      ❌ Table has insufficient rows")
        return []

    rows_text = []
    for r in rows:
        cells = r.find_all(['th', 'td'])
        rows_text.append([c.get_text(strip=True) for c in cells])

    # Find header row flexibly
    header_row_index = -1
    for i, r in enumerate(rows_text):
        lower = [t.lower() for t in r]
        has_member_id = any('member id' in t for t in lower)
        has_person = any(t in ('person', 'full name', 'name', 'first name') for t in lower)
        if has_member_id or (has_person and len(r) >= 4):
            header_row_index = i
            headers = r
            print(f"      ✅ Found headers at row {i}: {headers}")
            break

    if header_row_index == -1:
        print("      ❌ Could not find any suitable headers in New Visitors report")
        for i in range(min(3, len(rows_text))):
            print(f"         Row {i}: {rows_text[i]}")
        return []

    # Column mapping (case-insensitive)
    lower_headers = [h.lower() for h in headers]
    def find_col(*cands):
        for c in cands:
            if c in lower_headers:
                return lower_headers.index(c)
        return None

    member_id_col = find_col('member id')
    person_col = find_col('person', 'full name', 'name', 'first name')
    location_cols = [i for i, h in enumerate(lower_headers) if any(k in h for k in ('location', 'service', 'place', 'where'))]

    print(f"      🔍 Column mapping - Member ID: {member_id_col}, Person: {person_col}, Locations: {location_cols}")

    visitors = []
    for r in rows_text[header_row_index + 1:]:
        if len(r) < len(headers):
            continue
        member_id = r[member_id_col] if member_id_col is not None and member_id_col < len(r) else ''
        full_name = r[person_col] if person_col is not None and person_col < len(r) else ''
        locations = ' '.join(r[i] for i in location_cols if i < len(r)) if location_cols else ''

        # Try to discover UUID-like ID if missing
        if not member_id:
            for cell_text in r:
                if re.match(r'[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$', cell_text):
                    member_id = cell_text
                    break

        if full_name:
            visitors.append({
                'member_id': member_id or f"unknown_{len(visitors)+1}",
                'full_name': full_name,
                'locations': locations,
                'raw_data': dict(zip(headers, r[:len(headers)]))
            })

    print(f"      ✅ Extracted {len(visitors)} visitors for {year_label}")
    if visitors:
        print(f"      📋 Sample visitor: {visitors[0]['full_name']} | ID: {visitors[0]['member_id']} | Loc: '{visitors[0]['locations']}'")
    return visitors

def parse_category_change_report(html_content, year_label):
    """Parse People Category Change report and return rows where a visitor became Congregation_ or RosteredMember_."""
    if not html_content:
        return []

    print(f"   📊 Parsing Category Change data for {year_label}...")

    try:
        soup = BeautifulSoup(html_content, 'lxml')
    except Exception:
        soup = BeautifulSoup(html_content, 'html.parser')

    table = soup.find('table')
    if not table:
        print("      ❌ No table found in report")
        return []

    rows = table.find_all('tr')
    if len(rows) < 2:
        print("      ❌ Table has insufficient rows")
        return []

    # --- locate headers (person + change from/to + date). Member ID is optional.
    headers = []
    header_row_index = -1
    for i, row in enumerate(rows):
        cells = row.find_all(['th', 'td'])
        cell_texts = [cell.get_text(strip=True) for cell in cells]
        lower = [t.lower() for t in cell_texts]
        has_person      = any(t in ('person','full name','name','first name') for t in lower)
        has_change_from = any('change from' in t for t in lower)
        has_change_to   = any('change to'   in t for t in lower)
        has_date        = any(t == 'date' or ('date' in t and 'date added' not in t) for t in lower)
        if has_person and has_change_from and has_change_to and has_date:
            headers = cell_texts
            header_row_index = i
            print(f"      📋 Found headers at row {i}: {headers[:5]}...")
            break

    if not headers:
        print("      ❌ Could not find headers in Category Change report")
        return []

    lower_headers = [h.lower() for h in headers]

    def idx_of(*cand):
        for c in cand:
            if c in lower_headers:
                return lower_headers.index(c)
        return None

    person_col      = idx_of('person','full name','name','first name')
    change_from_col = idx_of('change from')
    change_to_col   = idx_of('change to')
    date_col        = idx_of('date')
    member_id_col   = idx_of('member id','person id','id')  # optional

    def norm(s: str) -> str:
        # lower, remove non-alphanum to single spaces, collapse slashes/spaces
        import re as _re
        return _re.sub(r'\s+', ' ', _re.sub(r'[^a-z0-9]+', ' ', (s or '').lower())).strip()

    stayed_people = []

    # --- parse rows
    for row in rows[header_row_index + 1:]:
        cells = row.find_all(['td', 'th'])
        if len(cells) < len(headers):
            continue

        def text_at(col):
            return cells[col].get_text(strip=True) if col is not None and col < len(cells) else ''

        person_name = text_at(person_col)
        change_from = text_at(change_from_col)
        change_to   = text_at(change_to_col)
        date_field  = text_at(date_col)
        member_id   = text_at(member_id_col) if member_id_col is not None else ''

        cf = norm(change_from)
        ct = norm(change_to)

        # Treat these as "visitor-ish" sources (ignore "former visitor")
        from_is_visitor = (
            ('visitor' in cf or 'newcomer' in cf) and 'former' not in cf
        )
        # destination is congregation or rostered member
        to_is_member = ('congregation' in ct) or ('rosteredmember' in ct) or ('rostered member' in ct)

        if from_is_visitor and to_is_member:
            stayed_people.append({
                'member_id': member_id,
                'person_name': person_name,
                'change_from': change_from,
                'change_to': change_to,
                'date': date_field,
                'raw_data': {
                    'Person': person_name,
                    'Member ID': member_id,
                    'Change From': change_from,
                    'Change To': change_to,
                    'Date': date_field
                }
            })

    # --- Debug preview
    print(f"      🔍 Category change analysis for {year_label}:")
    sample_changes = set()
    sample_people = []
    change_count = 0
    preview_rows = rows[header_row_index + 1: header_row_index + 21]
    for row in preview_rows:
        cells = row.find_all(['td','th'])
        def t(col): return cells[col].get_text(strip=True) if col is not None and col < len(cells) else ''
        pn, cf_txt, ct_txt = t(person_col), t(change_from_col), t(change_to_col)
        if cf_txt or ct_txt:
            change_count += 1
            sample_changes.add(f"'{cf_txt}' → '{ct_txt}'")
            if len(sample_people) < 5 and pn:
                sample_people.append(f"{pn}: '{cf_txt}' → '{ct_txt}'")

    print(f"         Total category changes found: {change_count}")
    if sample_changes:
        print("         Unique category transitions:")
        for ch in sorted(sample_changes):
            print(f"         {ch}")
    if sample_people:
        print("      🔍 Sample people with changes:")
        for ex in sample_people:
            print(f"         {ex}")

    print("      🎯 Looking for visitor→(Congregation_|RosteredMember_) after normalisation")
    print(f"      ✅ Found {len(stayed_people)} people who stayed for {year_label}")
    return stayed_people

def gross_up_factor_for_current_year(baseline_counts_by_service, counts_map, svc, today=None):
    """
    Returns a gross-up factor for current-year visitor ratios.
    Priority:
      1) If we have per-service counts for last year vs current, use baseline/current.
      2) Otherwise, fall back to a months-based factor: 12 / full_months_elapsed.
         (full months = months fully completed this year; e.g., on Sept 2 → 8)
    """
    from datetime import datetime
    today = today or datetime.now()
    # Try per-service counts first
    baseline = (baseline_counts_by_service or {}).get(svc, 0)
    current  = (counts_map or {}).get(svc, 0)
    if baseline and current:
        return baseline / current

    # Months fallback
    # Use completed months so far (Jan..Aug = 8 when early September)
    full_months = max(1, today.month - 1)
    return 12.0 / full_months


def classify_visitor_location(locations):
    """Classify visitor location into service categories - ENHANCED WITH DEBUGGING"""
    if not locations:
        return 'empty'
    
    locations_lower = locations.lower()
    
    # DEBUG: Show what we're classifying
    # print(f"      🔍 Classifying location: '{locations}' -> '{locations_lower}'")
    
    # Check for special events first
    if any(term in locations_lower for term in ['good friday', 'easter sunday', 'easter']):
        return 'Easter'
    
    if any(term in locations_lower for term in ['christmas eve', 'christmas day', '5pm christmas', '10:30pm christmas', 'christmas']):
        return 'Christmas'
    
    # Check for regular services - prioritize largest categories (10:30 > 8:30 > 6:30 > Mid-week)
    services_found = []
    
    # More flexible matching for service times
    if any(term in locations_lower for term in ['10:30', '10.30', 'ten thirty', '10 30']):
        services_found.append('10:30AM')
    if any(term in locations_lower for term in ['8:30', '8.30', 'eight thirty', '8 30']):
        services_found.append('8:30AM')
    if any(term in locations_lower for term in ['6:30', '6.30', 'six thirty', '6 30', '6:00', '6 00', 'evening']):
        services_found.append('6:30PM')
    if any(term in locations_lower for term in ['mid-week', 'midweek', 'wednesday', 'mid week', 'bible study']):
        services_found.append('Mid-week')
    
    # Also check for service names or descriptions
    if any(term in locations_lower for term in ['morning prayer', 'communion', 'morning service']):
        # If no specific time, default to 10:30AM (largest service)
        if not services_found:
            services_found.append('10:30AM')
    
    if any(term in locations_lower for term in ['evening prayer', 'evening service', 'evensong']):
        if not services_found:
            services_found.append('6:30PM')
    
    # If multiple services, return the largest (10:30 > 8:30 > 6:30 > Mid-week)
    if services_found:
        priority_order = ['10:30AM', '8:30AM', '6:30PM', 'Mid-week']
        for service in priority_order:
            if service in services_found:
                return service
    
    return 'empty'

def apply_pro_rata_estimation(visitors_by_service, year_label):
    """Apply pro-rata estimation to empty location visitors"""
    total_visitors = sum(visitors_by_service.values())
    empty_visitors = visitors_by_service.get('empty', 0)
    
    if empty_visitors == 0:
        return visitors_by_service
    
    # Calculate ratios from regular services only (exclude Christmas/Easter/empty)
    regular_services = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
    regular_totals = {service: visitors_by_service.get(service, 0) for service in regular_services}
    regular_sum = sum(regular_totals.values())
    
    if regular_sum == 0:
        print(f"      ⚠️ No regular service visitors found for {year_label}, cannot apply pro-rata")
        return visitors_by_service
    
    print(f"      📊 Applying pro-rata estimation to {empty_visitors} empty location visitors:")
    
    # Distribute empty visitors proportionally
    result = visitors_by_service.copy()
    for service in regular_services:
        if regular_sum > 0:
            proportion = regular_totals[service] / regular_sum
            additional = int(empty_visitors * proportion)
            result[service] = result.get(service, 0) + additional
            print(f"         {service}: +{additional} (ratio: {proportion:.2f})")
    
    # Remove empty category
    result.pop('empty', None)
    return result

def match_visitors_to_stayed(visitors_by_year, stayed_by_year):
    """
    Match visitors to stayed using:
      • Canonical person data from Elvanto people/getAll (one-time index)
      • UUID normalisation + 2-year lookback across visitor reports
      • Name/token fallbacks
      • Proportional allocation for any still-unmatched
    """

    import re
    from collections import defaultdict

    print("\n🔗 Matching visitors to those who stayed (API-backed people index)...")

    # ---------- helpers (scoped inside so this is a single drop-in) ----------
    uuid_re = re.compile(r'[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}', re.I)

    def normalize_uuid(s: str) -> str:
        if not s:
            return ''
        m = uuid_re.search(str(s))
        return m.group(0).lower() if m else ''

    def norm_name(s: str) -> str:
        s = (s or "").lower()
        s = re.sub(r'[^a-z0-9]+', ' ', s)
        return re.sub(r'\s+', ' ', s).strip()

    def name_tokens(s: str):
        return [t for t in norm_name(s).split() if t]

    def canonical_key(s: str) -> str:
        toks = name_tokens(s)
        return ' '.join(sorted(toks))

    def _dedupe_stayed_rows(rows):
        """
        Collapse multiple 'visitor -> member' transitions for the same person in the same year.
        Prefer the earliest-dated transition we saw (but date may be texty; if missing,
        we just keep the first encountered).
        """
        from datetime import datetime as _dt

        def _parse_date(s):
            # very forgiving: try dd/mm/yyyy, yyyy-mm-dd, mm/dd/yyyy
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%y", "%d-%m-%Y"):
                try:
                    return _dt.strptime(s, fmt).date()
                except Exception:
                    pass
            return None

        bucket = {}  # key -> (date_or_None, row)
        for sp in rows:
            mid = normalize_uuid(sp.get('member_id', ''))
            key = ('id', mid) if mid else ('name', canonical_key(sp.get('person_name', '')))
            d = _parse_date(sp.get('date', ''))  # may be None
            if key not in bucket:
                bucket[key] = (d, sp)
            else:
                old_d, _ = bucket[key]
                # keep the earliest transition if both dates exist; otherwise keep the first we saw
                if d and old_d and d < old_d:
                    bucket[key] = (d, sp)
        return [row for (_, row) in bucket.values()]

    def parse_last_first(person_name: str):
        """Return (first,last) best guess from 'Last, First' or 'First Last'."""
        if not person_name:
            return ('','')
        if ',' in person_name:
            last, first = [x.strip() for x in person_name.split(',', 1)]
        else:
            parts = person_name.strip().split()
            if len(parts) >= 2:
                first, last = parts[0], parts[-1]
            else:
                first, last = parts[0], ''
        return (first, last)

    def build_year_lookups(rows):
        """Per-year visitor lookups by id / canonical name / token index."""
        by_id = {}
        by_canon = {}
        token_index = defaultdict(list)
        for v in rows:
            mid = normalize_uuid(v.get('member_id', ''))
            if mid:
                by_id[mid] = v
            nm = v.get('full_name', '')
            if nm:
                ck = canonical_key(nm)
                if ck and ck not in by_canon:
                    by_canon[ck] = v
                for t in set(name_tokens(nm)):
                    token_index[t].append(v)
        return by_id, by_canon, token_index

    def fetch_people_index():
        """
        One-time pull of all people via API to map:
          id -> {first,last,preferred,full_variants}
          variant_name_key -> set(ids)
        Uses your existing make_request().
        """
        print("   🗂️ Building people index from Elvanto (people/getAll)…")
        page, page_size = 1, 1000
        id_to_person = {}
        name_to_ids = defaultdict(set)

        while True:
            resp = make_request('people/getAll', {
                'page': page,
                'page_size': page_size,
                # Pull only core fields – names are returned by default per API docs
                # You can add extra fields via 'fields': [...]
            })
            if not resp or resp.get('status') != 'ok':
                break

            people_obj = resp.get('people', {})
            persons = people_obj.get('person', [])
            if not isinstance(persons, list):
                persons = [persons] if persons else []

            for p in persons:
                pid = normalize_uuid(p.get('id', ''))
                if not pid:
                    continue
                first = p.get('firstname', '') or ''
                last = p.get('lastname', '') or ''
                pref = p.get('preferred_name', '') or ''

                # Store
                id_to_person[pid] = {'first': first, 'last': last, 'preferred': pref}

                # Build a few robust variants
                variants = set()
                first_last = f"{first} {last}".strip()
                last_first = f"{last}, {first}".strip(', ').strip()
                pref_last = f"{pref} {last}".strip()
                variants.update([first_last, last_first, pref_last, first, last, pref])

                for v in variants:
                    ck = canonical_key(v)
                    if ck:
                        name_to_ids[ck].add(pid)

            on_this_page = int(people_obj.get('on_this_page', len(persons) or 0))
            total = int(people_obj.get('total', 0) or 0)
            page += 1
            if on_this_page == 0 or (page - 1) * page_size >= total:
                break

        print(f"      ✅ People index ready: {len(id_to_person)} IDs; {len(name_to_ids)} name keys")
        return id_to_person, name_to_ids

    # ---------- build the people index once ----------
    id_to_person, name_to_ids = fetch_people_index()

    REGULAR = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
    matched_stayed = {}

    for year in stayed_by_year:
        print(f"\n   📅 Processing {year}...")
        stayed_people = stayed_by_year.get(year, [])
        before = len(stayed_people)
        stayed_people = _dedupe_stayed_rows(stayed_people)
        print(f"      🔁 De-duplicated stayed rows: {before} → {len(stayed_people)} unique people")

        # Lookback: year, year-1, year-2
        look_years = [year, year - 1, year - 2]
        lookup_chain = []
        for y in look_years:
            rows = visitors_by_year.get(y, [])
            id_map, canon_map, toks = build_year_lookups(rows)
            lookup_chain.append((y, id_map, canon_map, toks))
            print(f"      🔍 Year {y} visitors: {len(rows)} (IDs:{len(id_map)}, canonical:{len(canon_map)})")

        stayed_by_service = defaultdict(int)
        unmatched = []
        matched_ct = {'this': 0, 'minus1': 0, 'minus2': 0}

        for sp in stayed_people:
            person_name = sp.get('person_name', '')
            member_id = normalize_uuid(sp.get('member_id', ''))

            # If stayed row didn't parse an ID, try to recover it from the people index
            if not member_id and person_name:
                # Try 'Last, First' and 'First Last' variants via our index
                first, last = parse_last_first(person_name)
                candidates = []
                for cand in [
                    f"{first} {last}".strip(),
                    f"{last}, {first}".strip(', ').strip(),
                    first, last
                ]:
                    ck = canonical_key(cand)
                    if ck and ck in name_to_ids and len(name_to_ids[ck]) == 1:
                        member_id = next(iter(name_to_ids[ck]))
                        print(f"      🔎 Resolved ID for '{person_name}' → {member_id[:8]}… via people index")
                        break
                    elif ck and ck in name_to_ids and len(name_to_ids[ck]) > 1:
                        # keep as potential; we may still match by tokens below
                        candidates.extend(list(name_to_ids[ck]))

            found = None  # (yearMatched, visitor_record, method)

            # 1) Try by ID across lookback years
            if member_id:
                for (y, id_map, _, _) in lookup_chain:
                    if member_id in id_map:
                        found = (y, id_map[member_id], "Member ID")
                        break

            # 2) Canonical name, then token-overlap across lookback years
            if not found and person_name:
                ck = canonical_key(person_name)
                for (y, _, canon_map, _) in lookup_chain:
                    if ck and ck in canon_map:
                        found = (y, canon_map[ck], "Canonical name")
                        break
                if not found:
                    stoks = set(name_tokens(person_name))
                    best = None
                    for (y, _, _, toks) in lookup_chain:
                        candidates = []
                        for t in stoks:
                            for v in toks.get(t, []):
                                overlap = len(stoks & set(name_tokens(v.get('full_name', ''))))
                                if overlap > 0:
                                    candidates.append((overlap, y, v))
                        if candidates:
                            candidates.sort(key=lambda x: (-x[0], x[2].get('full_name', '')))
                            best = candidates[0]
                            break
                    if best:
                        found = (best[1], best[2], f"Token overlap {best[0]}")

            if found:
                ymatch, visitor_record, method = found
                svc = classify_visitor_location(visitor_record.get('locations', ''))
                stayed_by_service[svc] += 1
                bucket = 'this' if ymatch == year else ('minus1' if ymatch == year - 1 else 'minus2')
                matched_ct[bucket] += 1
                tag = {'this': 'this-year', 'minus1': 'prior-year', 'minus2': 'two-years-prior'}[bucket]
                if svc in REGULAR:
                    print(f"      ✅ Stayed match ({tag}): {person_name} → {visitor_record.get('full_name','?')} | {svc} (by {method})")
                elif svc == 'empty':
                    print(f"      ⚠️ Stayed match with empty location: {person_name} → {visitor_record.get('full_name','?')} (by {method}, {tag})")
                else:
                    print(f"      🎄🐣 Stayed match to special event: {person_name} → {visitor_record.get('full_name','?')} | {svc} (by {method}, {tag})")
            else:
                if member_id:
                    print(f"      ❌ No visitor row for {person_name} with ID={member_id[:8]}… in years {look_years}")
                else:
                    print(f"      ❌ No visitor row for {person_name} (no usable ID; names differ)")
                unmatched.append(sp)

        print(f"      📊 Matched breakdown — same-year:{matched_ct['this']} prior-1y:{matched_ct['minus1']} prior-2y:{matched_ct['minus2']} / total:{len(stayed_people)}")

        # Allocate still-unmatched using SAME-YEAR visitor distribution
        if unmatched:
            dist = {s: 0 for s in REGULAR}
            for v in visitors_by_year.get(year, []):
                s = classify_visitor_location(v.get('locations', ''))
                if s in REGULAR:
                    dist[s] += 1
            total = sum(dist.values())
            n = len(unmatched)
            print(f"      ➕ Allocating {n} unmatched to services based on {year} visitor mix: {dist}")
            if total == 0:
                per = n // len(REGULAR)
                rem = n % len(REGULAR)
                for i, s in enumerate(REGULAR):
                    add = per + (1 if i < rem else 0)
                    stayed_by_service[s] += add
                    print(f"         • {s}: +{add} (even split)")
            else:
                targets = {s: n * (dist[s] / total) for s in REGULAR}
                adds = {s: int(targets[s]) for s in REGULAR}
                rem = n - sum(adds.values())
                # largest fractional remainders get the extras (BUGFIX: correct unpacking)
                order = sorted(REGULAR, key=lambda s: (targets[s] - adds[s]), reverse=True)
                for s in order[:rem]:
                    adds[s] += 1
                for s in REGULAR:
                    stayed_by_service[s] += adds[s]
                    print(f"         • {s}: +{adds[s]} (share {dist[s]}/{total})")

        if stayed_by_service.get('empty', 0) > 0:
            stayed_by_service = apply_pro_rata_to_stayed(stayed_by_service, year)

        print(f"      📈 {year} final stayed by service: {dict(stayed_by_service)}")
        matched_stayed[year] = dict(stayed_by_service)

    return matched_stayed

def apply_pro_rata_to_stayed(stayed_by_service, year):
    """Apply pro-rata estimation to stayed people with empty locations"""
    empty_stayed = stayed_by_service.get('empty', 0)
    if empty_stayed == 0:
        return stayed_by_service
    
    print(f"      📊 Applying pro-rata to {empty_stayed} stayed people with empty locations:")
    
    regular_services = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
    regular_totals = {service: stayed_by_service.get(service, 0) for service in regular_services}
    regular_sum = sum(regular_totals.values())
    
    if regular_sum == 0:
        # If no regular service data, distribute evenly
        for service in regular_services:
            stayed_by_service[service] = stayed_by_service.get(service, 0) + (empty_stayed // len(regular_services))
    else:
        # Distribute proportionally
        for service in regular_services:
            if regular_sum > 0:
                proportion = regular_totals[service] / regular_sum
                additional = int(empty_stayed * proportion)
                stayed_by_service[service] = stayed_by_service.get(service, 0) + additional
                print(f"         {service}: +{additional} stayed (ratio: {proportion:.2f})")
    
    # Remove empty category
    stayed_by_service.pop('empty', None)
    return stayed_by_service

def get_congregation_averages_att_methodology():
    """
    Build {year: {'averages': {...}, 'service_counts': {...}}} for the last 3 years
    using the 'Service Individual Attendance' reports.

    Accepts BOTH of these parser return shapes:
      • {'8:30AM': 45.2, '10:30AM': ...}                         # bare averages (old)
      • {'averages': {...}, 'service_counts': {...}}             # new preferred

    If parsing or download fails for a year, fall back to get_default_single_year_stats(year).
    """
    print("\n📊 Calculating congregation averages using ATT methodology...")

    # 1) Discover the 3 service-attendance reports
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        print("   ❌ Failed to get groups from API")
        return get_default_congregation_stats()

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    service_reports = {}
    for g in groups:
        name = g.get('name', '')
        if name == 'Report of Service Individual Attendance':
            service_reports['current'] = g
            print(f"✅ Found current service report: {name}")
        elif name == 'Report of Last Year Service Individual Attendance':
            service_reports['last_year'] = g
            print(f"✅ Found last year service report: {name}")
        elif name == 'Report of Two Years Ago Service Individual Attendance':
            service_reports['two_years_ago'] = g
            print(f"✅ Found two years ago service report: {name}")

    # 2) For each of the 3 years, parse and normalize output
    stats_by_year = {}
    current_year = datetime.now().year
    year_plan = [
        ('current', current_year),
        ('last_year', current_year - 1),
        ('two_years_ago', current_year - 2),
    ]

    for key, yr in year_plan:
        if key not in service_reports:
            print(f"   ⚠️ {yr}: No service report group found, using defaults")
            stats_by_year[yr] = get_default_single_year_stats(yr)
            continue

        print(f"\n   🔄 Processing {yr} service attendance...")
        report_data = download_report_data(service_reports[key], f"service attendance {yr}")
        if not report_data:
            print(f"   ⚠️ {yr}: Download failed, using defaults")
            stats_by_year[yr] = get_default_single_year_stats(yr)
            continue

        out = parse_service_attendance_for_averages_att_style(report_data, yr)

        # ---- Normalize shapes ----
        if out and isinstance(out, dict):
            if 'averages' in out:
                # Preferred shape already
                stats_by_year[yr] = out
            else:
                # Old shape: a bare map of averages
                stats_by_year[yr] = {'averages': out, 'service_counts': {}}
        else:
            print(f"   ⚠️ {yr}: Parsing returned nothing usable, using defaults")
            stats_by_year[yr] = get_default_single_year_stats(yr)
            continue

        # Ensure both keys exist
        stats_by_year[yr].setdefault('averages', {})
        stats_by_year[yr].setdefault('service_counts', {})

        # Log what we ended up with
        avgs = stats_by_year[yr]['averages']
        counts = stats_by_year[yr]['service_counts']
        print(f"   ✅ {yr}: Using parsed averages"
              f"{' + service counts' if counts else ''}")
        for svc in ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']:
            if svc in avgs:
                print(f"      {svc}: avg={avgs.get(svc, 0):.1f}"
                      f"{(' • cols=' + str(counts.get(svc))) if svc in counts else ''}")

    return stats_by_year


def get_default_congregation_stats():
    """Defaults for three years when API fetch fails."""
    y = datetime.now().year
    return {
        y-2: get_default_single_year_stats(y-2),
        y-1: get_default_single_year_stats(y-1),
        y:   get_default_single_year_stats(y),
    }


def get_default_single_year_stats(year):
    """Defaults for one year: averages + rough service counts."""
    return {
        'averages': {
            '8:30AM': 45,
            '10:30AM': 85,
            '6:30PM': 25,
            'Mid-week': 15
        },
        # rough typical counts (Sundays only)
        'service_counts': {
            '8:30AM': 48,
            '10:30AM': 48,
            '6:30PM': 48
        }
    }

def parse_service_attendance_for_averages_att_style(html_content, year):
    """
    Parse 'Service Individual Attendance' by TIME IN HEADER (no column collapsing).

    Fixes:
      • Handles non-breaking space before AM/PM (e.g., '6:30\u00A0PM').
      • Accepts ':' or '.' in times; AM/PM any case; also '9:30 AM' → 10:30 bucket.
      • Keeps only headers whose DATE is in `year`.
      • Requires Sunday for 8:30/10:30/6:30; keeps Wednesday as 'Mid-week'.
      • Skips obvious non-service items.
      • Applies ATT filter (final January point only) if available.
    """
    if not html_content:
        return None

    from bs4 import BeautifulSoup
    from collections import defaultdict
    from datetime import datetime as dt
    import re

    YES = {'Y', 'YES', '✓', '✔', '1', 'TRUE'}

    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    if not tables:
        print("         ❌ No tables found in service report")
        return None

    # pick the widest grid-like table
    def width(t):
        r = t.find('tr')
        return len(r.find_all(['th','td'])) if r else 0
    table = max(tables, key=width)

    rows = table.find_all('tr')
    if len(rows) < 2:
        print("         ❌ Table has insufficient rows")
        return None

    header_cells = rows[0].find_all(['th','td'])
    headers = [(c.get_text(strip=True) or c.get('title') or c.get('abbr') or '') for c in header_cells]

    # ---------- helpers ----------
    num_date = re.compile(r'(\d{1,2})\s*[/\.\-]\s*(\d{1,2})(?:\s*[/\.\-]\s*(\d{2,4}))?')
    mon_date = re.compile(r'(\d{1,2})\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s*(\d{2,4})?', re.I)
    month_map = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'sept':9,'oct':10,'nov':11,'dec':12}

    # NOTE: allow normal spaces AND NBSP (\u00A0) between time and am/pm
    time_rx = re.compile(r'(\d{1,2})\s*[:\.]\s*(\d{2})(?:[\s\u00A0])*([AaPp][Mm])\b')

    NOISE = (
        'prayer meeting', 'weekly prayer', 'buzz', 'playgroup',
        'kids club', 'youth group', 'alpha', 'course', 'practice', 'rehearsal',
        'meeting', 'training'
    )
    def is_noise(h):
        h = (h or '').lower()
        return any(k in h for k in NOISE)

    def parse_header_date(h):
        h = (h or '').strip()
        m = num_date.search(h)
        if m:
            d = int(m.group(1)); mo = int(m.group(2)); yr = m.group(3)
        else:
            v = mon_date.search(h)
            if not v:
                return None
            d = int(v.group(1)); mo = month_map[v.group(2).lower()]; yr = v.group(3)

        yr = int(yr) if yr else int(year)
        if yr < 100: yr = 2000 + yr
        try:
            return dt(yr, mo, d).date()
        except Exception:
            return None

    def bucket_from_time_text(h):
        """
        Return '8:30AM'|'10:30AM'|'6:30PM'|None and a flag if it's explicitly evening.
        Matches '6:30 PM', '6.30pm', '10:30AM', '9:30 am' (→ 10:30 bucket), including NBSP before AM/PM.
        """
        m = time_rx.search(h or '')
        if not m:
            return (None, False)
        hh = int(m.group(1)); mm = int(m.group(2)); ampm = m.group(3).lower()
        if ampm == 'am':
            if hh == 8 and mm == 30:
                return ('8:30AM', False)
            if (hh == 10 and mm == 30) or (hh == 9 and mm == 30):
                return ('10:30AM', False)
            return (None, False)
        # PM
        if 18 <= hh <= 19:   # 6:00–7:00 PM -> evening bucket
            return ('6:30PM', True)
        return (None, True)

    def looks_evening_by_words(h):
        h = (h or '').lower()
        return ('evening' in h) or ('evensong' in h) or ('gathering' in h and 'evening' in h)

    # ---------- scan headers ----------
    total_cols = len(headers)
    kept_cols = 0
    dropped_noise = dropped_outyr = dropped_wrong_dow = dropped_no_time = 0
    eve_by_time = eve_by_word = 0

    per_service_records = defaultdict(list)  # svc -> list of {count,date,header}

    for idx, h in enumerate(headers):
        if not h or is_noise(h):
            if is_noise(h): dropped_noise += 1
            continue

        d = parse_header_date(h)
        if not d or d.year != int(year):
            dropped_outyr += 1
            continue

        bucket, explicit_eve = bucket_from_time_text(h)

        # If no time match but header clearly says 'Evening ...', treat as 6:30PM
        if bucket is None and looks_evening_by_words(h):
            bucket = '6:30PM'
            explicit_eve = True
            eve_by_word += 1

        if bucket in ('8:30AM', '10:30AM', '6:30PM'):
            # Require Sunday for these buckets
            if d.weekday() != 6:  # Monday=0 … Sunday=6
                dropped_wrong_dow += 1
                continue
            if bucket == '6:30PM':
                if explicit_eve:
                    eve_by_time += 1  # count those matched via time; word-matched already tallied
        else:
            # Maybe a Mid-week on Wednesday
            if d.weekday() == 2:  # Wednesday
                bucket = 'Mid-week'
            else:
                dropped_no_time += 1
                continue

        # count 'Y' in this column
        y_count = 0
        for r in rows[1:]:
            cells = r.find_all(['td','th'])
            if idx < len(cells):
                v = (cells[idx].get_text(strip=True) or '').upper()
                if v in YES:
                    y_count += 1

        per_service_records[bucket].append({
            'count': y_count,
            'date': f"{d.day}/{d.month}/{d.year}",
            'header': h
        })
        kept_cols += 1

    print(f"         ⏱ Scanned {total_cols} columns → kept {kept_cols} in {year}")
    if dropped_noise:      print(f"           • Ignored non-service columns: {dropped_noise}")
    if dropped_outyr:      print(f"           • Dropped out-of-year or undated columns: {dropped_outyr}")
    if dropped_wrong_dow:  print(f"           • Dropped non-Sunday columns for 8:30/10:30/6:30: {dropped_wrong_dow}")
    if dropped_no_time:    print(f"           • Skipped columns without a usable time (not Wed Mid-week): {dropped_no_time}")
    if eve_by_time or eve_by_word:
        print(f"           • Evening recognised by time: {eve_by_time}; by 'evening' keyword: {eve_by_word}")

    # ---------- averages (apply ATT filter) ----------
    averages = {}
    for svc in ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']:
        recs = per_service_records.get(svc, [])
        if not recs:
            averages[svc] = 0.0
            print(f"         ⚠️ {svc}: 0 services")
            continue

        filtered = filter_for_average_calculation_att_style(recs) if 'filter_for_average_calculation_att_style' in globals() else recs
        if not filtered:  # safety
            filtered = recs

        avg = sum(r['count'] for r in filtered) / len(filtered)
        averages[svc] = avg
        print(f"         ✅ {svc}: {avg:.1f} (from {len(filtered)} services)")

    # After computing `averages` …
    service_counts = {
        '8:30AM':  len(per_service_records.get('8:30AM',  [])),
        '10:30AM': len(per_service_records.get('10:30AM', [])),
        '6:30PM':  len(per_service_records.get('6:30PM',  [])),
        'Mid-week':len(per_service_records.get('Mid-week',[])),
    }
    return {'averages': averages, 'service_counts': service_counts}

def parse_service_column_header_att_style(header, year):
    """
    Parse a service column header into:
      {
        'normalized_time': '8:30AM' | '10:30AM' | '6:30PM',
        'time': 'HH:MM AM/PM',
        'day': int,
        'month': int,
        'year': int,
        'date_key': 'YYYY-MM-DD'
      }
    Accepts '8:30' or '8.30', any case 'am/pm', and multiple date styles.
    """
    import re
    h = str(header)

    # ---- time: allow ':' or '.' ----
    tm = re.search(r'(\d{1,2})\s*[:\.]\s*(\d{2})\s*(am|pm)\b', h, re.I)
    if not tm:
        return None
    hh = int(tm.group(1))
    mm = int(tm.group(2))
    ampm = tm.group(3).upper()
    time_norm = f"{hh}:{mm:02d} {ampm}"

    # normalize to one of our buckets
    bucket = None
    tkey = f"{hh:02d}:{mm:02d}"
    if tkey == "08:30" and ampm == "AM":
        bucket = "8:30AM"
    elif tkey in ("10:30", "09:30") and ampm == "AM":
        bucket = "10:30AM"  # treat 9:30 combined as 10:30 bucket
    elif (tkey in ("06:30", "06:00") and ampm == "PM") or (hh in (18,) and ampm == "PM"):
        bucket = "6:30PM"
    else:
        return None

    # ---- date (handle 01/09, 1-9, 10 Sep 2024, etc.) ----
    month_map = {
        'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
        'jul':7,'aug':8,'sep':9,'sept':9,'oct':10,'nov':11,'dec':12
    }

    d = m = None

    # numeric styles: 14/01, 14-01, 14.01 (optionally with year at the end)
    dm = re.search(r'(\d{1,2})\s*[/\.\-]\s*(\d{1,2})(?:\s*[/\.\-]\s*(\d{2,4}))?', h)
    if dm:
        d = int(dm.group(1)); m = int(dm.group(2))
    else:
        # e.g., '14 Sep', '14 September 2024'
        vw = re.search(r'(\d{1,2})\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s*(\d{2,4})?', h, re.I)
        if vw:
            d = int(vw.group(1)); m = month_map[vw.group(2).lower()]

    if not d or not m:
        # default (rare) – we still return a record but with 1/1
        d, m = 1, 1

    y = year
    date_key = f"{y:04d}-{m:02d}-{d:02d}"
    return {
        'normalized_time': bucket,
        'time': time_norm,
        'day': d,
        'month': m,
        'year': y,
        'date_key': date_key
    }

def filter_for_average_calculation_att_style(records):
    """ATT: keep final January datum (if any) + all non-January dates. Supports parsed month/day."""
    if not records:
        return []

    def m_of(r):
        if isinstance(r.get('month'), int):
            return r['month']
        try:
            parts = r.get('date', '').split('/')
            if len(parts) == 3:
                return int(parts[1])
        except Exception:
            pass
        return None

    jan = [r for r in records if m_of(r) == 1]
    other = [r for r in records if m_of(r) != 1]

    out = []
    if jan:
        # choose latest January by 'day' if available
        jan_sorted = sorted(jan, key=lambda r: (r.get('day') if isinstance(r.get('day'), int) else -1))
        out.append(jan_sorted[-1])
    out.extend(other)
    return out


def get_default_congregation_averages():
    """Get default congregation sizes when service reports are unavailable"""
    current_year = datetime.now().year
    
    default_averages = {}
    for year in [current_year - 2, current_year - 1, current_year]:
        default_averages[year] = get_default_single_year(year)
    
    print("   ⚠️ Using default congregation averages (ATT methodology failed)")
    for year, averages in default_averages.items():
        print(f"      {year}: {averages}")
    
    return default_averages

def get_default_single_year(year):
    """Get default congregation sizes for a single year"""
    return {
        '8:30AM': 45,
        '10:30AM': 85,
        '6:30PM': 25,
        'Mid-week': 15
    }

def create_plotly_charts(visitor_data, stayed_data, congregation_averages):
    """
    Stacked portrait layout with a single compact 'Changes' table:
      • Four charts stacked vertically (unchanged metrics/series).
      • One Changes table showing:
          - Δ (Last Year − Two Years Ago)
          - Δ (This Year − Last Year)
    No changes to calculations, only layout and the table contents.
    """
    from datetime import datetime
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go

    print("\n📊 Creating Plotly charts in stacked (portrait) layout + Changes table…")

    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    y_23, y_24, y_25 = years  # two_years_ago, last_year, this_year

    regular_services = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
    services = ['Overall'] + regular_services  # Overall first for headline

    # Helpers to read both shapes from congregation_averages
    def get_avg_map(y):
        yval = congregation_averages.get(y, {})
        if isinstance(yval, dict) and 'averages' in yval:
            return yval.get('averages', {})
        return yval if isinstance(yval, dict) else {}

    def get_counts_map(y):
        yval = congregation_averages.get(y, {})
        if isinstance(yval, dict) and 'service_counts' in yval:
            return yval.get('service_counts', {})
        return {}

    last_year = y_24
    two_years_ago = y_23
    baseline_counts_by_service = get_counts_map(last_year) or get_counts_map(two_years_ago) or get_counts_map(current_year)

    # Compute metrics (uses visitors_for_ratio() defined earlier)
    visitor_counts = {s: [] for s in services}
    stayed_counts  = {s: [] for s in services}
    visitor_ratios = {s: [] for s in services}
    stay_ratios    = {s: [] for s in services}
    metrics = {s: {y: {'visitors': 0, 'vratio': 0.0, 'stayed': 0, 'sratio': 0.0} for y in years} for s in services}

    # Build per-year metrics
    for year in years:
        year_visitors = visitor_data.get(year, {})
        year_stayed   = stayed_data.get(year, {})
        avg_map       = get_avg_map(year)
        counts_map    = get_counts_map(year)

        total_visitors_raw = 0
        total_stayed = 0
        total_avg_congregation = 0
        total_visitors_for_ratio = 0.0

        for svc in regular_services:
            v = year_visitors.get(svc, 0)
            s = year_stayed.get(svc, 0)
            avg_cong = avg_map.get(svc, 50)

            v_for_ratio = visitors_for_ratio(year, svc, v, counts_map, baseline_counts_by_service)
            vr = (v_for_ratio / avg_cong * 100) if avg_cong > 0 else 0.0
            sr = (s / v * 100) if v > 0 else 0.0

            # stash per-service arrays for charts (append in order y_23,y_24,y_25)
            if svc not in visitor_counts:  # because we prefixed 'Overall'
                visitor_counts[svc] = []
                stayed_counts[svc] = []
                visitor_ratios[svc] = []
                stay_ratios[svc] = []

            visitor_counts[svc].append(v)
            stayed_counts[svc].append(s)
            visitor_ratios[svc].append(vr)
            stay_ratios[svc].append(sr)

            metrics.setdefault(svc, {})[year] = {'visitors': v, 'vratio': vr, 'stayed': s, 'sratio': sr}

            total_visitors_raw += v
            total_stayed += s
            total_avg_congregation += avg_cong
            total_visitors_for_ratio += v_for_ratio

        # Overall row
        ov_vis = total_visitors_raw
        ov_sty = total_stayed
        ov_vr  = (total_visitors_for_ratio / total_avg_congregation * 100) if total_avg_congregation > 0 else 0.0
        ov_sr  = (ov_sty / ov_vis * 100) if ov_vis > 0 else 0.0

        visitor_counts['Overall'].append(ov_vis)
        stayed_counts['Overall'].append(ov_sty)
        visitor_ratios['Overall'].append(ov_vr)
        stay_ratios['Overall'].append(ov_sr)

        metrics['Overall'][year] = {'visitors': ov_vis, 'vratio': ov_vr, 'stayed': ov_sty, 'sratio': ov_sr}

    # ---- Layout: 5 rows (4 charts + 1 changes table) x 1 col; A4-ish portrait width ----
    fig = make_subplots(
        rows=5, cols=1,
        specs=[
            [{"type": "xy"}],      # 1: Visitor Numbers
            [{"type": "xy"}],      # 2: Visitor Ratios
            [{"type": "xy"}],      # 3: Stayed Numbers
            [{"type": "xy"}],      # 4: Stay Ratios
            [{"type": "domain"}],  # 5: Changes table only
        ],
        row_heights=[0.19, 0.19, 0.19, 0.19, 0.24],
        vertical_spacing=0.05,
        subplot_titles=[
            'Visitor Numbers by Service',
            'Visitor Ratios (% of Congregation)',
            'Stay Numbers by Service',
            'Stay Ratios (%)',
            'Year-over-Year Changes'
        ]
    )

    service_colors = {
        '8:30AM':   '#dc2626',
        '10:30AM':  '#2563eb',
        '6:30PM':   '#059669',
        'Mid-week': '#ea580c',
        'Overall':  '#f59e0b'
    }
    year_labels = [str(y) for y in years]

    # Charts: rows 1–4
    for svc in services:
        fig.add_trace(go.Bar(x=year_labels, y=visitor_counts[svc], name=svc,
                             marker_color=service_colors[svc], legendgroup='services', showlegend=True),
                      row=1, col=1)
    for svc in services:
        fig.add_trace(go.Bar(x=year_labels, y=visitor_ratios[svc], name=svc,
                             marker_color=service_colors[svc], legendgroup='services', showlegend=False),
                      row=2, col=1)
    for svc in services:
        fig.add_trace(go.Bar(x=year_labels, y=stayed_counts[svc], name=svc,
                             marker_color=service_colors[svc], legendgroup='services', showlegend=False),
                      row=3, col=1)
    for svc in services:
        fig.add_trace(go.Bar(x=year_labels, y=stay_ratios[svc], name=svc,
                             marker_color=service_colors[svc], legendgroup='services', showlegend=False),
                      row=4, col=1)

    # --- CHANGES TABLE: include both intervals (LY−2Y) and (TY−LY) ---
    def fmt_i(n): return int(round(n))
    def fmt_p(x): return f"{x:.1f}"

    table_header = [
        "Service",
        # Visitors block
        f"Visitors {y_23}", f"Visitors {y_24}", f"Δ {y_24}-{y_23}",
        f"Visitors {y_25}", f"Δ {y_25}-{y_24}",
        # Visitor Ratio block
        f"VR% {y_23}", f"VR% {y_24}", f"Δ pp {y_24}-{y_23}",
        f"VR% {y_25}", f"Δ pp {y_25}-{y_24}",
        # Stayed block
        f"Stayed {y_23}", f"Stayed {y_24}", f"Δ {y_24}-{y_23}",
        f"Stayed {y_25}", f"Δ {y_25}-{y_24}",
        # Stay Ratio block
        f"SR% {y_23}", f"SR% {y_24}", f"Δ pp {y_24}-{y_23}",
        f"SR% {y_25}", f"Δ pp {y_25}-{y_24}",
    ]

    table_rows = []
    for svc in services:  # Overall first, then services
        m23 = metrics[svc][y_23]; m24 = metrics[svc][y_24]; m25 = metrics[svc][y_25]
        table_rows.append([
            svc,
            # Visitors
            fmt_i(m23['visitors']), fmt_i(m24['visitors']), fmt_i(m24['visitors'] - m23['visitors']),
            fmt_i(m25['visitors']), fmt_i(m25['visitors'] - m24['visitors']),
            # Visitor Ratio (pp)
            fmt_p(m23['vratio']), fmt_p(m24['vratio']), f"{(m24['vratio'] - m23['vratio']):.1f}",
            fmt_p(m25['vratio']), f"{(m25['vratio'] - m24['vratio']):.1f}",
            # Stayed
            fmt_i(m23['stayed']), fmt_i(m24['stayed']), fmt_i(m24['stayed'] - m23['stayed']),
            fmt_i(m25['stayed']), fmt_i(m25['stayed'] - m24['stayed']),
            # Stay Ratio (pp)
            fmt_p(m23['sratio']), fmt_p(m24['sratio']), f"{(m24['sratio'] - m23['sratio']):.1f}",
            fmt_p(m25['sratio']), f"{(m25['sratio'] - m24['sratio']):.1f}",
        ])

    fig.add_trace(
        go.Table(
            header=dict(
                values=table_header,
                fill_color='#e2e8f0',
                font=dict(color='#0f172a', size=11),
                align='center', height=26
            ),
            cells=dict(
                values=list(map(list, zip(*table_rows))),
                fill_color='#ffffff',
                align='right',
                height=22,
                font=dict(size=10, color='#334155')
            ),
            columnwidth=[
                94,  # Service
                # Visitors block (5)
                86, 86, 86, 86, 86,
                # VR% block (5)
                80, 80, 86, 80, 86,
                # Stayed block (5)
                86, 86, 86, 86, 86,
                # SR% block (5)
                80, 80, 86, 80, 86
            ]
        ),
        row=5, col=1
    )

    # A4-ish portrait size (moderately wide)
    fig.update_layout(
        title=dict(
            text="<b>Visitor and Stay Dashboard</b><br><span style='font-size:16px; color:#64748b'>Three-Year Analysis Across Four Congregations + Overall Church</span>",
            x=0.5, y=0.988,
            font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=24, color='#1e293b')
        ),
        font=dict(family="Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif", size=12),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=1980, width=1400,   # a touch wider to accommodate the single wide table
        barmode='group',
        legend=dict(
            orientation="v",
            yanchor="top", y=1.0,
            xanchor="right", x=0.99,
            font=dict(size=12, color='#374151'),
            bgcolor='rgba(255,255,255,0.9)'
        ),
        margin=dict(l=50, r=50, t=160, b=80)
    )

    fig.update_xaxes(showgrid=False, showline=True, linewidth=1, linecolor='#e2e8f0',
                     tickfont=dict(size=11, color='#374151'))
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#f1f5f9',
                     showline=True, linewidth=1, linecolor='#e2e8f0',
                     tickfont=dict(size=11, color='#374151'))
    fig.update_yaxes(title_text="Number of Visitors", row=1, col=1)
    fig.update_yaxes(title_text="Visitor Ratio (%)", row=2, col=1)
    fig.update_yaxes(title_text="Number Who Stayed", row=3, col=1)
    fig.update_yaxes(title_text="Stay Ratio (%)", row=4, col=1)

    return fig

def create_visitor_stay_dashboard(visitor_data, stayed_data, congregation_averages):
    """Create the complete Visitor and Stay Dashboard with charts (portrait PNG)."""
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    
    print(f"\n🎨 Creating Visitor and Stay Dashboard with charts (portrait export, A4-ish)…")
    
    fig = create_plotly_charts(visitor_data, stayed_data, congregation_averages)
    html_content = generate_dashboard_html_with_chart(fig, visitor_data, stayed_data, congregation_averages, years)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    filename = f"visitor_stay_dashboard_{timestamp}.html"
    
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"✅ Dashboard saved as: {filename}")
        
        try:
            import webbrowser, os
            png_filename = filename.replace('.html', '.png')
            # Match the figure’s new size
            fig.write_image(png_filename, width=1400, height=1980, scale=2)
            print(f"✅ Chart saved as PNG: {png_filename}")
            webbrowser.open(f"file://{os.path.abspath(png_filename)}")
            print(f"🚀 Auto-opened PNG: {png_filename}")
        except Exception as e:
            print(f"⚠️ PNG export/open failed (install kaleido and pillow): {e}")
            try:
                png_filename = filename.replace('.html', '.png')
                fig.write_image(png_filename, width=1400, height=1980, scale=2)
                print(f"✅ Chart saved as PNG: {png_filename}")
            except:
                print("❌ PNG export completely failed")
    except Exception as e:
        print(f"❌ Failed to save dashboard: {e}")
        return None

    return filename

def generate_dashboard_html_with_chart(fig, visitor_data, stayed_data, congregation_averages, years):
    """
    Generate HTML dashboard with embedded Plotly chart + numeric tables under the charts.
    Tables include:
      • Data Table: Visitors, Visitor Ratio (% of congregation), Stayed, Stay Ratio (%) by Year × Service.
      • YoY Change Table: Current vs Last Year deltas (counts and percentage-points).

    Uses the same gross-up logic as create_plotly_charts() for Visitor Ratio (current year only):
      1) Prefer per-service service_counts ratio (baseline last year ÷ current year).
      2) Fallback to months-based factor: 12 ÷ full_months_elapsed.
    """
    import sys
    from datetime import datetime
    import plotly.io as pio
    from bs4 import BeautifulSoup

    # --- helper (scoped here so this function is self-contained) ---
    def _gross_up_factor_for_current_year(baseline_counts_by_service, counts_map, svc, today=None):
        """
        Returns a gross-up factor for current-year visitor ratios.
        Priority:
          1) If we have per-service counts for last year vs current, use baseline/current.
          2) Otherwise, fall back to a months-based factor: 12 / full_months_elapsed.
             (full months = months fully completed this year; e.g., on Sept 2 → 8)
        """
        today = today or datetime.now()
        baseline = (baseline_counts_by_service or {}).get(svc, 0)
        current  = (counts_map or {}).get(svc, 0)
        if baseline and current:
            return baseline / current
        full_months = max(1, today.month - 1)   # completed months this year
        return 12.0 / full_months

    # ---- Convert Plotly figure to an embeddable div ----
    chart_html = pio.to_html(fig, include_plotlyjs='cdn', div_id='plotly-div')
    soup = BeautifulSoup(chart_html, 'lxml') if 'lxml' in sys.modules or True else BeautifulSoup(chart_html, 'html.parser')
    chart_div = soup.find('div', {'id': 'plotly-div'})

    # ---- Helpers to read old/new congregation_averages shapes ----
    def get_avg_map(y):
        yval = congregation_averages.get(y, {})
        if isinstance(yval, dict) and 'averages' in yval:
            return yval.get('averages', {})
        return yval if isinstance(yval, dict) else {}

    def get_counts_map(y):
        yval = congregation_averages.get(y, {})
        if isinstance(yval, dict) and 'service_counts' in yval:
            return yval.get('service_counts', {})
        return {}

    current_year = years[-1]
    last_year = years[-2] if len(years) >= 2 else None
    two_years_ago = years[-3] if len(years) >= 3 else None

    # Baseline for gross-up is LAST YEAR's service counts; fallback to two years ago, then current
    baseline_counts_by_service = get_counts_map(last_year) if last_year else {}
    if not baseline_counts_by_service and two_years_ago is not None:
        baseline_counts_by_service = get_counts_map(two_years_ago)
    if not baseline_counts_by_service:
        baseline_counts_by_service = get_counts_map(current_year)

    regular_services = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
    all_services = regular_services + ['Overall']

    # ---- Compute the numbers that back the charts (with gross-up for current year visitor ratio) ----
    metrics = {y: {svc: {'visitors': 0, 'vratio': 0.0, 'stayed': 0, 'sratio': 0.0} for svc in all_services} for y in years}

    for y in years:
        avg_map = get_avg_map(y)
        counts_map = get_counts_map(y)

        total_visitors_raw = 0
        total_visitors_for_ratio = 0.0  # after gross-up
        total_stayed = 0
        total_avg = 0

        for svc in regular_services:
            v = visitor_data.get(y, {}).get(svc, 0)
            s = stayed_data.get(y, {}).get(svc, 0)
            avg_cong = avg_map.get(svc, 50)

            # Gross-up ONLY for current year
            v_for_ratio = visitors_for_ratio(
                y, svc, v, counts_map, baseline_counts_by_service
            )
            
            vratio = (v_for_ratio / avg_cong * 100) if avg_cong > 0 else 0.0
            sratio = (s / v * 100) if v > 0 else 0.0

            metrics[y][svc] = {'visitors': v, 'vratio': vratio, 'stayed': s, 'sratio': sratio}

            total_visitors_raw += v
            total_visitors_for_ratio += v_for_ratio
            total_stayed += s
            total_avg += avg_cong

        # Overall row
        metrics[y]['Overall']['visitors'] = total_visitors_raw
        metrics[y]['Overall']['stayed'] = total_stayed
        metrics[y]['Overall']['vratio'] = (total_visitors_for_ratio / total_avg * 100) if total_avg > 0 else 0.0
        metrics[y]['Overall']['sratio'] = (total_stayed / total_visitors_raw * 100) if total_visitors_raw > 0 else 0.0

    # ---- Build the Data Table HTML ----
    def fmt_int(n): return f"{int(round(n))}"
    def fmt_pct(x, decimals=1): return f"{x:.{decimals}f}%"
    def fmt_pp(x, decimals=1): return f"{x:.{decimals}f} pp"

    table_rows_html = ""
    for svc in all_services:
        for y in years:
            m = metrics[y][svc]
            table_rows_html += f"""
                <tr>
                    <td class="t-service">{svc}</td>
                    <td>{y}</td>
                    <td class="num">{fmt_int(m['visitors'])}</td>
                    <td class="num">{fmt_pct(m['vratio'])}</td>
                    <td class="num">{fmt_int(m['stayed'])}</td>
                    <td class="num">{fmt_pct(m['sratio'])}</td>
                </tr>
            """

    # YoY change table (Current vs Last Year)
    yoy_rows_html = ""
    if last_year is not None:
        for svc in all_services:
            this_m = metrics[current_year][svc]
            last_m = metrics[last_year][svc]
            d_vis = this_m['visitors'] - last_m['visitors']
            d_vr  = this_m['vratio'] - last_m['vratio']   # pp
            d_sty = this_m['stayed'] - last_m['stayed']
            d_sr  = this_m['sratio'] - last_m['sratio']   # pp

            def delta_cell(val, is_pp=False):
                cls = "pos" if val > 0 else ("neg" if val < 0 else "zero")
                txt = fmt_pp(val) if is_pp else fmt_int(val)
                return f'<td class="num delta {cls}">{txt}</td>'

            yoy_rows_html += f"""
                <tr>
                    <td class="t-service">{svc}</td>
                    <td class="num">{fmt_int(last_m['visitors'])}</td>
                    <td class="num">{fmt_int(this_m['visitors'])}</td>
                    {delta_cell(d_vis)}
                    <td class="num">{fmt_pct(last_m['vratio'])}</td>
                    <td class="num">{fmt_pct(this_m['vratio'])}</td>
                    {delta_cell(d_vr, is_pp=True)}
                    <td class="num">{fmt_int(last_m['stayed'])}</td>
                    <td class="num">{fmt_int(this_m['stayed'])}</td>
                    {delta_cell(d_sty)}
                    <td class="num">{fmt_pct(last_m['sratio'])}</td>
                    <td class="num">{fmt_pct(this_m['sratio'])}</td>
                    {delta_cell(d_sr, is_pp=True)}
                </tr>
            """

    # ---- Full HTML (existing layout + NEW tables section right under the chart) ----
    html = f'''<!DOCTYPE html>
<html>
<head>
    <title>Visitor and Stay Dashboard</title>
    <meta charset="UTF-8">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {{
            font-family: Inter, -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
            margin: 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            color: #1e293b;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            overflow: hidden;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            margin-top: 20px;
            margin-bottom: 20px;
        }}
        .header {{
            background: linear-gradient(135deg, #dc2626, #b91c1c);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        .header h1 {{ font-size: 2.5rem; font-weight: bold; margin: 0 0 15px 0; }}
        .header p {{ font-size: 1.1rem; margin: 5px 0; opacity: 0.9; }}

        .chart-container {{ padding: 20px; background: white; }}

        /* NEW: Tables under the charts */
        .tables-section {{
            background: #ffffff;
            padding: 8px 24px 32px;
            border-top: 3px solid #e2e8f0;
        }}
        .table-title {{
            color: #1e293b;
            font-size: 1.35rem;
            font-weight: 700;
            margin: 10px 0 12px;
        }}
        table.data {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            overflow: hidden;
            border-radius: 12px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
            margin-bottom: 18px;
        }}
        table.data thead th {{
            background: #f1f5f9;
            color: #0f172a;
            font-weight: 600;
            text-align: center;
            padding: 10px 8px;
            border-bottom: 1px solid #e2e8f0;
            position: sticky;
            top: 0;
            z-index: 1;
        }}
        table.data tbody td {{
            padding: 10px 8px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }}
        table.data tbody tr:nth-child(even) {{ background: #fafafa; }}
        td.t-service {{ font-weight: 600; color: #0f172a; }}
        td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
        .delta.pos {{ color: #059669; font-weight: 700; }}
        .delta.neg {{ color: #dc2626; font-weight: 700; }}
        .delta.zero {{ color: #475569; font-weight: 700; }}

        .special-events {{ background: linear-gradient(135deg, #fbbf24, #f59e0b); color: white; padding: 25px; border-radius: 12px; margin: 20px; }}
        .special-events h3 {{ margin-top: 0; font-size: 1.4rem; }}
        .special-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-top: 15px; }}
        .special-card {{ background: rgba(255,255,255,0.2); padding: 15px; border-radius: 8px; text-align: center; }}
        .special-number {{ font-size: 2rem; font-weight: bold; margin-bottom: 5px; }}
        .special-label {{ font-size: 0.9rem; opacity: 0.9; }}

        .summary-section {{
            background: #f8fafc;
            padding: 30px;
            border-top: 3px solid #e2e8f0;
        }}
        .summary-title {{
            color: #1e293b;
            font-size: 1.8rem;
            font-weight: bold;
            margin-bottom: 25px;
            text-align: center;
            border-bottom: 2px solid #3498db;
            padding-bottom: 15px;
        }}
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .year-card {{
            background: white;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-left: 5px solid;
        }}
        .year-title {{ font-size: 1.4rem; font-weight: bold; color: #1e293b; margin-bottom: 20px; }}
        .metric-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; padding: 8px 0; border-bottom: 1px solid #f1f5f9; }}
        .metric-label {{ font-weight: 500; color: #475569; }}
        .metric-value {{ font-weight: bold; color: #1e293b; font-size: 1.1rem; }}

        .methodology {{ background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 12px; padding: 25px; margin: 20px; }}
        .methodology h3 {{ color: #1e293b; margin-top: 0; font-size: 1.3rem; }}
        .methodology ul {{ color: #475569; line-height: 1.7; }}
        .methodology li {{ margin-bottom: 8px; }}

        .footer {{ background: #1e293b; color: #94a3b8; text-align: center; padding: 20px; font-size: 0.9rem; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚪 Visitor and Stay Dashboard</h1>
            <p>Comprehensive Analysis of Visitor Numbers and Stay Rates</p>
            <p>Three-Year Comparison Across Four Congregations</p>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>

        <div class="chart-container">
            {chart_div}
        </div>

        <!-- NEW: Numeric tables under the charts -->
        <div class="tables-section">
            <div class="table-title">📋 Data Table — values behind the charts</div>
            <table class="data">
                <thead>
                    <tr>
                        <th>Service</th>
                        <th>Year</th>
                        <th>Visitors</th>
                        <th>Visitor Ratio (%)</th>
                        <th>Stayed</th>
                        <th>Stay Ratio (%)</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows_html}
                </tbody>
            </table>
            <div class="table-title" style="margin-top:26px;">↕️ Year-over-Year Change — Current vs Last Year</div>
            <table class="data">
                <thead>
                    <tr>
                        <th rowspan="2">Service</th>
                        <th colspan="3">Visitors</th>
                        <th colspan="3">Visitor Ratio (pp)</th>
                        <th colspan="3">Stayed</th>
                        <th colspan="3">Stay Ratio (pp)</th>
                    </tr>
                    <tr>
                        <th>{last_year if last_year is not None else ''}</th>
                        <th>{current_year}</th>
                        <th>Δ</th>
                        <th>{last_year if last_year is not None else ''}</th>
                        <th>{current_year}</th>
                        <th>Δ</th>
                        <th>{last_year if last_year is not None else ''}</th>
                        <th>{current_year}</th>
                        <th>Δ</th>
                        <th>{last_year if last_year is not None else ''}</th>
                        <th>{current_year}</th>
                        <th>Δ</th>
                    </tr>
                </thead>
                <tbody>
                    {yoy_rows_html if last_year is not None else ''}
                </tbody>
            </table>
            <p style="color:#64748b; font-size:0.9rem;">Visitor Ratio uses a grossed-up numerator for the current year so partial-year data is comparable with last year.</p>
        </div>

        <div class="special-events">
            <h3>🎄🐣 Special Events Visitors</h3>
            <div class="special-grid">'''

    # Special events cards
    for y in years:
        visitors = visitor_data.get(y, {})
        christmas = visitors.get('Christmas', 0)
        easter = visitors.get('Easter', 0)
        html += f'''
                <div class="special-card">
                    <div class="special-number">{christmas}</div>
                    <div class="special-label">Christmas {y}</div>
                </div>
                <div class="special-card">
                    <div class="special-number">{easter}</div>
                    <div class="special-label">Easter {y}</div>
                </div>'''

    html += '''
            </div>
        </div>

        <div class="summary-section">
            <div class="summary-title">📊 Three-Year Summary</div>
            <div class="metrics-grid">'''

    # Year cards
    for y in years:
        visitors = visitor_data.get(y, {})
        stayed = stayed_data.get(y, {})
        averages = get_avg_map(y)
        total_visitors = sum(visitors.get(s, 0) for s in regular_services)
        total_stayed = sum(stayed.get(s, 0) for s in regular_services)
        total_avg_congregation = sum(averages.get(s, 0) for s in regular_services)
        overall_visitor_ratio = (total_visitors / total_avg_congregation * 100) if total_avg_congregation > 0 else 0
        overall_stay_ratio = (total_stayed / total_visitors * 100) if total_visitors > 0 else 0

        html += f'''
                <div class="year-card year-{y}">
                    <div class="year-title">{y}</div>
                    <div class="metric-row">
                        <span class="metric-label"><strong>OVERALL CHURCH:</strong></span>
                        <span class="metric-value">{total_visitors}v/{total_stayed}s ({overall_visitor_ratio:.1f}%/{overall_stay_ratio:.1f}%)</span>
                    </div>'''

        for s in regular_services:
            v_count = visitors.get(s, 0)
            s_count = stayed.get(s, 0)
            avg_cong = averages.get(s, 0)
            v_ratio = (v_count / avg_cong * 100) if avg_cong > 0 else 0
            s_ratio = (s_count / v_count * 100) if v_count > 0 else 0
            html += f'''
                    <div class="metric-row">
                        <span class="metric-label">{s}:</span>
                        <span class="metric-value">{v_count}v/{s_count}s ({v_ratio:.0f}%/{s_ratio:.0f}%)</span>
                    </div>'''
        html += '''</div>'''

    html += f'''
            </div>
        </div>

        <div class="methodology">
            <h3>📋 Methodology & Notes</h3>
            <ul>
                <li><strong>Data Sources:</strong> Elvanto "People Category Change" and "New Visitors" reports for three years</li>
                <li><strong>Visitor Classification:</strong> Locations assigned to largest service (10:30AM &gt; 8:30AM &gt; 6:30PM &gt; Mid-week)</li>
                <li><strong>Stay Definition:</strong> Visitor categories moving to "Congregation_" or "RosteredMember_"</li>
                <li><strong>Name Matching:</strong> Member ID then exact/partial name</li>
                <li><strong>Empty Locations:</strong> Distributed proportionally (pro-rata)</li>
                <li><strong>Special Events:</strong> Good Friday &amp; Easter → Easter; Christmas Eve/Day → Christmas</li>
                <li><strong>Visitor Ratios:</strong> For the current year only, numerators are grossed up by a per-service factor. If service-counts data is missing, a months-based factor (12 ÷ completed months) is used.</li>
            </ul>
        </div>

        <div class="footer">
            Dashboard generated from Elvanto API data • {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
        </div>
    </div>
</body>
</html>'''

    return html

def main():
    """Main execution function"""
    print("🚀 Starting Visitor and Stay Dashboard Analysis...")
    print("="*70)
    
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    
    print(f"📅 Analyzing data for years: {years}")
    
    # Step 1: Find all required reports using Elvanto API
    category_reports, visitor_reports = find_visitor_and_category_reports()
    
    if not category_reports or not visitor_reports:
        print("❌ Could not find required reports in Elvanto")
        return
    
    # Step 2: Download and parse New Visitors reports
    print("\n" + "="*60)
    print("📥 DOWNLOADING AND PARSING NEW VISITORS REPORTS")
    print("="*60)
    
    visitors_by_year = {}
    for year_key, year_value in [('two_years_ago', years[0]), ('last_year', years[1]), ('current', years[2])]:
        if visitor_reports.get(year_key):
            print(f"\n🔄 Processing New Visitors for {year_value}...")
            report_data = download_report_data(visitor_reports[year_key], f"New Visitors {year_value}")
            if report_data:
                visitors_raw = parse_new_visitors_report(report_data, str(year_value))
                visitors_by_year[year_value] = visitors_raw
            else:
                visitors_by_year[year_value] = []
                print(f"      ❌ Failed to download New Visitors report for {year_value}")
        else:
            visitors_by_year[year_value] = []
            print(f"      ❌ No New Visitors report found for {year_value}")
    
    # Step 3: Download and parse People Category Change reports  
    print("\n" + "="*60)
    print("📥 DOWNLOADING AND PARSING PEOPLE CATEGORY CHANGE REPORTS")
    print("="*60)
    
    stayed_by_year = {}
    for year_key, year_value in [('two_years_ago', years[0]), ('last_year', years[1]), ('current', years[2])]:
        if category_reports.get(year_key):
            print(f"\n🔄 Processing Category Changes for {year_value}...")
            report_data = download_report_data(category_reports[year_key], f"Category Change {year_value}")
            if report_data:
                stayed_raw = parse_category_change_report(report_data, str(year_value))
                stayed_by_year[year_value] = stayed_raw
            else:
                stayed_by_year[year_value] = []
                print(f"      ❌ Failed to download Category Change report for {year_value}")
        else:
            stayed_by_year[year_value] = []
            print(f"      ❌ No Category Change report found for {year_value}")
    
    # Step 4: Process visitor data by service
    print("\n" + "="*50)
    print("📊 PROCESSING VISITOR DATA BY SERVICE")
    print("="*50)
    
    visitors_by_service_year = {}
    for year in years:
        visitors = visitors_by_year.get(year, [])
        visitors_by_service = defaultdict(int)
        
        print(f"\n📅 Processing {year} visitors ({len(visitors)} total):")
        
        # DEBUG: Show sample locations to understand the data
        if visitors:
            print(f"      🔍 Sample locations from first 5 visitors:")
            for i, visitor in enumerate(visitors[:5]):
                locations = visitor.get('locations', '')
                print(f"         {i+1}. {visitor['full_name']}: '{locations}'")
        
        location_debug = defaultdict(list)
        for visitor in visitors:
            locations = visitor.get('locations', '')
            service = classify_visitor_location(locations)
            visitors_by_service[service] += 1
            location_debug[service].append(f"{visitor['full_name']}: '{locations}'")
        
        # Show classification results
        print(f"      📊 Initial classification results:")
        for service, count in visitors_by_service.items():
            print(f"         {service}: {count} visitors")
            if service == 'empty' and count > 0:
                print(f"            Sample empty locations: {location_debug[service][:3]}")
        
        # Apply pro-rata estimation to empty locations
        visitors_by_service = apply_pro_rata_estimation(dict(visitors_by_service), str(year))
        visitors_by_service_year[year] = visitors_by_service
        
        print(f"   📊 Final visitor distribution: {visitors_by_service}")
    
    # Step 5: Match stayed people to their congregations
    print("\n" + "="*50)
    print("🔗 MATCHING STAYED PEOPLE TO CONGREGATIONS")
    print("="*50)
    
    stayed_by_service_year = match_visitors_to_stayed(visitors_by_year, stayed_by_year)
    
    # Step 6: Get congregation averages using ATT methodology
    print("\n" + "="*50)
    print("📊 CALCULATING CONGREGATION AVERAGES (ATT METHOD)")
    print("="*50)
    
    congregation_averages = get_congregation_averages_att_methodology()
    
    # Step 7: Create dashboard with charts
    print("\n" + "="*30)
    print("🎨 CREATING DASHBOARD WITH CHARTS")
    print("="*30)
    
    dashboard_file = create_visitor_stay_dashboard(
        visitors_by_service_year, 
        stayed_by_service_year, 
        congregation_averages
    )
    
    # Summary
    print("\n" + "="*50)
    print("📋 FINAL DASHBOARD SUMMARY")
    print("="*50)
    
    for year in years:
        visitors = visitors_by_service_year.get(year, {})
        stayed = stayed_by_service_year.get(year, {})
        
        services = ['8:30AM', '10:30AM', '6:30PM', 'Mid-week']
        total_visitors = sum(visitors.get(s, 0) for s in services)
        total_stayed = sum(stayed.get(s, 0) for s in services)
        
        stay_rate = (total_stayed / total_visitors * 100) if total_visitors > 0 else 0
        
        print(f"\n🗓️ {year}:")
        print(f"   👥 Regular Service Visitors: {total_visitors}")
        print(f"   🎄 Christmas Visitors: {visitors.get('Christmas', 0)}")
        print(f"   🐣 Easter Visitors: {visitors.get('Easter', 0)}")
        print(f"   🏠 People Who Stayed: {total_stayed}")
        print(f"   📈 Overall Stay Rate: {stay_rate:.1f}%")
        
        # Service breakdown
        for service in services:
            v_count = visitors.get(service, 0)
            s_count = stayed.get(service, 0)
            s_rate = (s_count / v_count * 100) if v_count > 0 else 0
            print(f"      {service}: {v_count} visitors, {s_count} stayed ({s_rate:.1f}%)")
    
    print(f"\n✅ Dashboard Analysis Complete!")
    print(f"📄 Dashboard saved as: {dashboard_file}")
    print(f"📊 Professional charts with PNG export included")
    print("\n" + "="*70)

if __name__ == "__main__":
    main()
