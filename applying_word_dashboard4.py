#!/usr/bin/env python3
"""
Applying the Word Dashboard (Code AW)
Analyzes Bible Study Group participation and Church Day Away attendance
- Bible Study Group Attendance as % of congregation average attendance  
- Number of Bible Study Groups over time
- Church Day Away attendance as % of combined 10:30+6:30 congregation
"""

import subprocess
import sys

# Auto-install required packages
def install_packages():
    packages = ['beautifulsoup4', 'plotly', 'requests', 'pandas', 'kaleido']
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

import requests
import json
from datetime import datetime, timedelta
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import re
from collections import defaultdict
from bs4 import BeautifulSoup
import numpy as np
import webbrowser
import os

print("📖 ST GEORGE'S MAGILL - APPLYING THE WORD DASHBOARD")
print("="*60)
print("📊 Bible Study Group Attendance & Church Day Away Analysis")
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
# --- NEW: helpers for 'regulars' logic ---

REGULAR_CATEGORY_NAMES = {"Congregation_", "RosteredMember_"}

def is_regular_category(label: str) -> bool:
    """
    UI strings from the change report can have punctuation/spacing variations.
    Normalize and accept 'Congregation' and 'Rostered Member' equivalents.
    """
    import re
    if not isinstance(label, str):
        return False
    norm = re.sub(r'[^a-z]', '', label.lower())
    return norm in {'congregation', 'rosteredmember'}

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

def _parse_date_any(s):
    """Parse a date in several common Elvanto formats; return datetime or None."""
    if not s:
        return None
    s = str(s).strip()
    from datetime import datetime
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None

def _load_people_meta_cache(path="people_meta_cache.json"):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            # Convert ISO strings back to datetimes
            out = {}
            for pid, data in raw.items():
                dd = _parse_date_any(data.get("date_deceased"))
                da = _parse_date_any(data.get("date_archived"))
                out[pid] = {"date_deceased": dd, "date_archived": da}
            return out
    except Exception as e:
        print(f"   ⚠️ Could not load cache {path}: {e}")
    return {}

def _save_people_meta_cache(cache_dict, path="people_meta_cache.json"):
    try:
        serializable = {}
        for pid, data in cache_dict.items():
            dd = data.get("date_deceased")
            da = data.get("date_archived")
            serializable[pid] = {
                "date_deceased": dd.isoformat() if dd else None,
                "date_archived": da.isoformat() if da else None,
            }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(serializable, f)
    except Exception as e:
        print(f"   ⚠️ Could not save cache {path}: {e}")


def enrich_people_meta_for_snapshot_pruning(person_ids, people_meta, cache_path="people_meta_cache.json"):
    """
    Enrich ONLY the people who actually need lifecycle dates for pruning:
      - If currently deceased -> fetch Date Deceased (once, then cached)
      - If currently archived -> fetch Date Archived (if your tenant exposes it)
    Uses a JSON cache on disk to avoid repeat API calls across runs.

    Updates people_meta[pid] with:
      'date_deceased': datetime|None
      'date_archived': datetime|None
    """
    print(f"🧩 Enriching lifecycle meta for snapshot pruning…")

    # Load cache and prefill anything we already know
    cache = _load_people_meta_cache(cache_path)
    prefilled = 0
    for pid in person_ids:
        if pid in cache:
            pm = people_meta.get(pid, {})
            if "date_deceased" not in pm:
                pm["date_deceased"] = cache[pid].get("date_deceased")
            if "date_archived" not in pm:
                pm["date_archived"] = cache[pid].get("date_archived")
            people_meta[pid] = pm
            prefilled += 1
    if prefilled:
        print(f"   • Prefilled from cache: {prefilled}")

    # Build the set of IDs that truly need fetching
    to_fetch = []
    for pid in person_ids:
        meta = people_meta.get(pid, {})
        # We only fetch if the person is currently deceased/archived
        # and we don't already have the corresponding date.
        need_dd = meta.get("deceased", False) and ("date_deceased" not in meta or meta["date_deceased"] is None)
        need_da = meta.get("archived", False) and ("date_archived" not in meta or meta["date_archived"] is None)
        if need_dd or need_da:
            to_fetch.append(pid)

    if not to_fetch:
        print("   • No API fetch needed (all lifecycle dates present or not applicable).")
        return

    # Helpers to normalize shapes
    def _coerce_person(person_obj):
        if isinstance(person_obj, dict):
            return person_obj
        if isinstance(person_obj, list):
            for item in person_obj:
                if isinstance(item, dict):
                    return item
        return {}

    def _iter_demographics(person_dict):
        demo = person_dict.get('demographics')
        if not demo:
            return []
        if isinstance(demo, dict) and 'demographic' in demo:
            demo_list = demo['demographic']
        else:
            demo_list = demo
        if not isinstance(demo_list, list):
            demo_list = [demo_list] if demo_list else []
        return [d for d in demo_list if isinstance(d, dict)]

    def _first_date_deceased(person_dict):
        for k in ('date_deceased', 'deceased_date', 'date_of_death', 'dateOfDeath'):
            dt = _parse_date_any(person_dict.get(k))
            if dt:
                return dt
        for d in _iter_demographics(person_dict):
            name = (d.get('name') or '').strip().lower()
            if any(k in name for k in ('deceased', 'date deceased', 'date of death', 'death')):
                dt = _parse_date_any(d.get('value'))
                if dt:
                    return dt
        return None

    def _first_date_archived(person_dict):
        for k in ('date_archived', 'archived_date', 'dateArchived'):
            dt = _parse_date_any(person_dict.get(k))
            if dt:
                return dt
        for d in _iter_demographics(person_dict):
            name = (d.get('name') or '').strip().lower()
            if 'archived' in name:
                dt = _parse_date_any(d.get('value'))
                if dt:
                    return dt
        return None

    api_calls = 0
    found_dd = 0
    found_da = 0

    # Fetch just what we need
    for pid in to_fetch:
        info = make_request('people/getInfo', {'id': pid, 'fields': ['demographics']})
        api_calls += 1
        if not info or 'person' not in info:
            continue
        person = _coerce_person(info['person'])
        if not person:
            continue

        pm = people_meta.get(pid, {})

        if pm.get("deceased", False) and pm.get("date_deceased") is None:
            dd = _first_date_deceased(person)
            if dd:
                pm["date_deceased"] = dd
                found_dd += 1

        if pm.get("archived", False) and pm.get("date_archived") is None:
            da = _first_date_archived(person)
            if da:
                pm["date_archived"] = da
                found_da += 1

        people_meta[pid] = pm
        # Update cache as we go
        cached = cache.get(pid, {})
        cached["date_deceased"] = pm.get("date_deceased")
        cached["date_archived"] = pm.get("date_archived")
        cache[pid] = cached

    _save_people_meta_cache(cache, cache_path)
    print(f"   ✅ Enrichment complete. API calls: {api_calls} | Date Deceased found: {found_dd} | Date Archived found: {found_da}")

def _prune_by_lifecycle_at_snapshot(person_ids, people_meta, snapshot_dt, label):
    """
    Remove anyone who (a) has a Date Deceased <= snapshot, or
    (b) has a Date Archived <= snapshot.
    Prints removals for visibility.
    """
    if not person_ids:
        print(f"   • {label}: nothing to prune by lifecycle")
        return person_ids

    removed_deceased = 0
    removed_archived = 0
    kept = set()

    for pid in person_ids:
        meta = people_meta.get(pid, {})
        dd = meta.get('date_deceased')
        da = meta.get('date_archived')

        # If we positively know they were deceased by snapshot → exclude
        if dd and dd <= snapshot_dt:
            removed_deceased += 1
            continue

        # If we positively know they were archived by snapshot → exclude
        if da and da <= snapshot_dt:
            removed_archived += 1
            continue

        kept.add(pid)

    if removed_deceased or removed_archived:
        print(f"   • {label}: lifecycle pruning removed "
              f"{removed_deceased} deceased + {removed_archived} archived (<= {snapshot_dt.date()})")
    else:
        print(f"   • {label}: no lifecycle removals (good)")

    return kept

def fetch_category_lookup():
    """
    Return {category_id: category_name} using people/categories/getAll.
    """
    print("📋 Fetching people categories (for id→name map)...")
    resp = make_request('people/categories/getAll')
    lookup = {}
    if resp and resp.get('categories'):
        cats = resp['categories'].get('category', [])
        if not isinstance(cats, list):
            cats = [cats] if cats else []
        for c in cats:
            cid = c.get('id')
            name = c.get('name')
            if cid and name:
                lookup[cid] = name
    print(f"   ✅ Loaded {len(lookup)} categories")
    return lookup

def find_category_change_reports():
    """
    Find the three People Category Change report groups.
    Returns a dict like:
      {
        'current':  <group_obj for 'Report of People Category Change'>,
        'last_year': <group_obj for 'Report of Last Year People Category Change'>,
        'two_years_ago': <group_obj for 'Report of Two Years Ago People Category Change'>
      }
    """
    print("\n📋 Searching for People Category Change reports...")

    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return {}

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    out = {}
    for g in groups:
        name = g.get('name', '')
        if name == 'Report of People Category Change':
            out['current'] = g
            print("✅ Found: Report of People Category Change")
        elif name == 'Report of Last Year People Category Change':
            out['last_year'] = g
            print("✅ Found: Report of Last Year People Category Change")
        elif name == 'Report of Two Years Ago People Category Change':
            out['two_years_ago'] = g
            print("✅ Found: Report of Two Years Ago People Category Change")

    # Light warning if any missing
    for key in ['current', 'last_year', 'two_years_ago']:
        if key not in out:
            print(f"⚠️ Missing People Category Change group: {key}")

    return out

def parse_category_change_report(data_content, year_label: str):
    """
    Parse the People Category Change HTML into a list of events:
      {'ts': datetime, 'member_id': str, 'change_from': str, 'change_to': str}

    We normalize the 'from'/'to' labels so UI strings like "Congregation," or "Rostered Member*"
    are recognized as regulars.
    """
    events = []
    if not data_content:
        return events

    print(f"🧾 Parsing People Category Change report for {year_label}...")

    try:
        html = data_content.decode('utf-8', errors='ignore')
        soup = BeautifulSoup(html, 'html.parser')

        table = soup.find('table')
        if not table:
            print("   ❌ No table found in change report")
            return events

        rows = table.find_all('tr')
        if not rows:
            print("   ❌ No rows in change report table")
            return events

        headers = [th.get_text(strip=True).lower() for th in rows[0].find_all(['th', 'td'])]
        col = {}
        for i, h in enumerate(headers):
            if h == 'date' or h.startswith('date '):
                col['date'] = i
            elif 'change from' in h:
                col['from'] = i
            elif 'change to' in h:
                col['to'] = i
            elif 'member id' in h:
                col['member'] = i

        needed = {'date', 'from', 'to', 'member'}
        if not needed.issubset(col):
            print(f"   ❌ Missing expected columns. Headers: {headers}")
            return events

        def parse_ts(s):
            for fmt in [
                "%d %B, %Y %I:%M %p", "%d %B %Y %I:%M %p",
                "%d %b, %Y %I:%M %p",  "%d %b %Y %I:%M %p",
                "%d %B, %Y", "%d %B %Y", "%d %b, %Y", "%d %b %Y",
            ]:
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            return None

        for r in rows[1:]:
            cells = r.find_all(['td', 'th'])
            if len(cells) <= max(col.values()):
                continue

            raw_date = cells[col['date']].get_text(strip=True)
            change_from = cells[col['from']].get_text(strip=True) or ""
            change_to   = cells[col['to']].get_text(strip=True) or ""
            member_id   = cells[col['member']].get_text(strip=True)

            if not member_id:
                continue

            ts = parse_ts(raw_date)
            if ts is None:
                continue

            events.append({
                'ts': ts,
                'member_id': member_id,
                'change_from': change_from,  # keep raw (we normalize when applying)
                'change_to': change_to,
            })

        events.sort(key=lambda e: e['ts'])
        print(f"   ✅ Parsed {len(events)} change events for {year_label}")
        return events

    except Exception as e:
        print(f"   ❌ Error parsing change report ({year_label}): {e}")
        return events

def get_current_regular_ids(exclude_archived=True, exclude_deceased=True):
    """
    Reliable 'regulars today' via category_id→name mapping with robust pagination.
    Excludes deceased and (by default) archived using robust truthy checks.

    Regulars = category_name exactly 'Congregation_' or 'RosteredMember_'.
    """
    def _truthy(v):
        # Handles 1, "1", True, "true", "yes", "y", "on" (case-insensitive)
        return str(v).strip().lower() in {"1", "true", "yes", "y", "on"}

    # 1) Build category_id -> name lookup
    cat_lookup = fetch_category_lookup()
    if not cat_lookup:
        print("   ⚠️ No categories returned; cannot determine current regulars.")
        return set()

    page = 1
    page_size = 1000
    regular_ids = set()
    counts_by_cat = {"Congregation_": 0, "RosteredMember_": 0}
    total_scanned = 0
    pages_reported_by_api = None

    # Debug counters
    excluded_archived = 0
    excluded_deceased = 0
    excluded_status_deceased = 0

    print("\n🧑‍🤝‍🧑 Scanning people for current regulars (via category_id)...")

    while True:
        # IMPORTANT: do NOT pass unsupported 'fields' keys here.
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
            # Robust exclusions (match overall dashboard intent)
            if exclude_deceased and _truthy(p.get('deceased', 0)):
                excluded_deceased += 1
                continue
            status = (p.get('status') or "").strip().lower()
            if exclude_deceased and status == 'deceased':
                excluded_status_deceased += 1
                continue
            if exclude_archived and _truthy(p.get('archived', 0)):
                excluded_archived += 1
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

        # Robust pagination:
        paging = resp.get('paging') or {}
        if 'pages' in paging:
            try:
                pages_reported_by_api = int(paging.get('pages'))
            except Exception:
                pages_reported_by_api = None

        if pages_reported_by_api is not None:
            if page >= pages_reported_by_api:
                break
            page += 1
        else:
            if batch_count < page_size:
                break
            page += 1

    print(f"   ✅ Current regulars: {len(regular_ids)} (scanned {total_scanned} people across {page} page(s))")
    print(f"      • Congregation_:   {counts_by_cat['Congregation_']}")
    print(f"      • RosteredMember_: {counts_by_cat['RosteredMember_']}")
    if exclude_archived or exclude_deceased:
        print("   🚫 Exclusions applied:")
        if exclude_archived:
            print(f"      • Archived excluded: {excluded_archived}")
        if exclude_deceased:
            print(f"      • Deceased (flag) excluded: {excluded_deceased}")
            print(f"      • Deceased (status) excluded: {excluded_status_deceased}")

    # Guardrail: highlight if de-dup changed totals
    sum_by_cat = counts_by_cat['Congregation_'] + counts_by_cat['RosteredMember_']
    if len(regular_ids) != sum_by_cat:
        print(f"      • Note: unique ID count {len(regular_ids)} vs sum-by-category {sum_by_cat} (duplicates de-duped)")

    return regular_ids

def fetch_people_meta():
    """
    Fetch minimal person metadata we need for accurate snapshots.
    Returns: { person_id: { 'date_added': datetime|None,
                            'archived': bool,
                            'deceased': bool,
                            'status': str } }
    """
    from datetime import datetime
    import re

    def _truthy(v):
        return str(v).strip().lower() in {"1", "true", "yes", "y", "on"}

    meta = {}
    page = 1
    page_size = 1000
    total = 0

    print("📇 Fetching minimal person metadata (id, date_added, archived/deceased/status)...")
    while True:
        resp = make_request('people/getAll', {'page': page, 'page_size': page_size})
        if not resp or not resp.get('people'):
            break

        people = resp['people'].get('person', [])
        if not isinstance(people, list):
            people = [people] if people else []

        batch = len(people)
        total += batch

        for p in people:
            pid = p.get('id')
            if not pid:
                continue

            # Parse date_added (Elvanto usually returns 'YYYY-MM-DD HH:MM:SS')
            raw = p.get('date_added') or ""
            dt = None
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    dt = datetime.strptime(raw, fmt) if raw else None
                    break
                except Exception:
                    pass

            meta[pid] = {
                'date_added': dt,
                'archived': _truthy(p.get('archived', 0)),
                'deceased': _truthy(p.get('deceased', 0)),
                'status': (p.get('status') or "").strip().lower(),
            }

        # pagination fallback
        paging = resp.get('paging') or {}
        pages = paging.get('pages')
        if pages is not None:
            try:
                pages = int(pages)
            except Exception:
                pages = None

        if pages is not None:
            if page >= pages:
                break
            page += 1
        else:
            if batch < page_size:
                break
            page += 1

    print(f"   ✅ Loaded meta for {len(meta)} people (scanned {total})")
    return meta

def _prune_by_existence_at_snapshot(person_ids, people_meta, snapshot_dt, label):
    """
    Remove anyone whose date_added is AFTER the snapshot date (they didn't exist yet).
    Returns the pruned set. Prints how many were removed.
    """
    if not person_ids:
        print(f"   • {label}: nothing to prune by existence")
        return person_ids

    removed = 0
    kept = set()
    for pid in person_ids:
        info = people_meta.get(pid)
        if info and info.get('date_added') and info['date_added'] > snapshot_dt:
            removed += 1
            continue
        kept.add(pid)

    if removed:
        print(f"   • {label}: removed {removed} people who were created after {snapshot_dt.date()}")
    else:
        print(f"   • {label}: no removals by date_added (good)")

    return kept


def reconstruct_regular_sets(today_regular_ids, events_by_year, people_meta):
    """
    Reverse-replay to get regulars at the start of each year, then prune by:
      1) Existence at snapshot (date_added > snapshot → remove)
      2) Lifecycle at snapshot (date_deceased/date_archived <= snapshot → remove)
    Now enriches lifecycle dates ONLY for people who could need them
    (currently deceased/archived), and caches results.
    """
    from datetime import datetime

    current_year = datetime.now().year
    jan1_this = datetime(current_year, 1, 1)
    jan1_last = datetime(current_year - 1, 1, 1)
    jan1_2yrs = datetime(current_year - 2, 1, 1)

    def apply_reverse(events_list, working_set, label_for_print):
        stats = {'events_total': 0, 'into_regular': 0, 'out_of_regular': 0,
                 'regular_to_regular': 0, 'non_to_non': 0}
        before_size = len(working_set)

        for ev in reversed(events_list or []):
            stats['events_total'] += 1
            frm = (ev.get('change_from') or "")
            to  = (ev.get('change_to') or "")
            pid = ev.get('member_id')
            if not pid:
                continue

            from_is_regular = is_regular_category(frm)
            to_is_regular   = is_regular_category(to)

            if (not from_is_regular) and to_is_regular:
                stats['into_regular'] += 1
                if pid in working_set:
                    working_set.remove(pid)
            elif from_is_regular and (not to_is_regular):
                stats['out_of_regular'] += 1
                working_set.add(pid)
            elif from_is_regular and to_is_regular:
                stats['regular_to_regular'] += 1
            else:
                stats['non_to_non'] += 1

        after_size = len(working_set)
        print(f"\n🔎 Reverse-replay summary for {label_for_print}")
        print(f"   • Events processed: {stats['events_total']}")
        print(f"   • Forward INTO regular: {stats['into_regular']}  (reverse → removals)")
        print(f"   • Forward OUT OF regular: {stats['out_of_regular']}  (reverse → additions)")
        print(f"   • Regular ↔ Regular (ignored): {stats['regular_to_regular']}")
        print(f"   • Non-regular ↔ Non-regular (ignored): {stats['non_to_non']}")
        print(f"   • Set size before: {before_size}")
        print(f"   • Set size after : {after_size}")
        return working_set

    print("\n🧮 RECONSTRUCTING REGULARS AT YEAR STARTS")
    print(f"   Current regulars (today): {len(today_regular_ids)}")

    # Reverse replay
    start_this_year = apply_reverse(events_by_year.get(current_year, []), set(today_regular_ids), f"{current_year} (→ start {current_year})")
    start_last_year = apply_reverse(events_by_year.get(current_year - 1, []), set(start_this_year), f"{current_year - 1} (→ start {current_year - 1})")
    start_two_years = apply_reverse(events_by_year.get(current_year - 2, []), set(start_last_year), f"{current_year - 2} (→ start {current_year - 2})")

    # Prune by existence at snapshot
    print("\n🧼 Pruning reconstructed sets by existence at snapshot...")
    start_this_year = _prune_by_existence_at_snapshot(start_this_year, people_meta, jan1_this, f"Start {current_year}")
    start_last_year = _prune_by_existence_at_snapshot(start_last_year, people_meta, jan1_last, f"Start {current_year - 1}")
    start_two_years = _prune_by_existence_at_snapshot(start_two_years, people_meta, jan1_2yrs, f"Start {current_year - 2}")

    # Enrich lifecycle dates ONLY for people who are currently deceased/archived
    candidates = set().union(start_this_year, start_last_year, start_two_years)
    candidates = {pid for pid in candidates if people_meta.get(pid, {}).get('deceased') or people_meta.get(pid, {}).get('archived')}
    if candidates:
        enrich_people_meta_for_snapshot_pruning(candidates, people_meta)
    else:
        print("🧩 No lifecycle enrichment needed (no candidates currently deceased/archived).")

    # Prune by lifecycle at snapshot (uses dates if we have them)
    print("\n🧼 Pruning reconstructed sets by lifecycle at snapshot (deceased/archived dates)…")
    start_this_year = _prune_by_lifecycle_at_snapshot(start_this_year, people_meta, jan1_this, f"Start {current_year}")
    start_last_year = _prune_by_lifecycle_at_snapshot(start_last_year, people_meta, jan1_last, f"Start {current_year - 1}")
    start_two_years = _prune_by_lifecycle_at_snapshot(start_two_years, people_meta, jan1_2yrs, f"Start {current_year - 2}")

    print("\n✅ DENOMINATORS (regulars at 1 Jan after pruning by existence + lifecycle)")
    print(f"   {current_year - 2}: {len(start_two_years)}")
    print(f"   {current_year - 1}: {len(start_last_year)}")
    print(f"   {current_year}: {len(start_this_year)}")

    return {
        'start_this_year': start_this_year,
        'start_last_year': start_last_year,
        'start_two_years_ago': start_two_years,
    }

def compute_biblestudy_percentages_using_regulars(attendance_data, regular_sets):
    """
    Uses reconstructed denominators (regulars @ 1 Jan of each year).
    Prints the numerator and denominator used for each bar.
    """
    current_year = datetime.now().year
    yrs = [current_year - 2, current_year - 1, current_year]

    denominators = {
        current_year:     len(regular_sets['start_this_year']),
        current_year - 1: len(regular_sets['start_last_year']),
        current_year - 2: len(regular_sets['start_two_years_ago']),
    }

    print("\n📏 TOP-CHART INPUTS (Bible Study % of regulars @ year start)")
    for y in yrs:
        print(f"   Year {y}: denominator (regulars @ 1 Jan) = {denominators.get(y, 0)}")

    out = {}
    for y in yrs:
        numerator = attendance_data.get(y, {}).get('_unique_people_count', 0)
        denom = denominators.get(y, 0)
        pct = (numerator / denom * 100.0) if denom > 0 else 0.0
        print(f"   Year {y}: numerator (unique attendees ≥2) = {numerator}  →  {numerator}/{denom} = {pct:.1f}%")
        out[y] = pct

    return out

def find_attendance_reports():
    """Find service and group attendance reports"""
    print("\n📋 Searching for attendance reports...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return {}, {}

    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []

    service_reports = {}
    group_reports = {}

    print(f"   Scanning {len(groups)} groups for attendance reports...")
    
    for group in groups:
        group_name = group.get('name', '')
        
        # Service attendance reports (for congregation size calculation)
        if group_name == 'Report of Service Individual Attendance':
            service_reports['current'] = group
            print(f"✅ Found current service report: {group_name}")
        elif group_name == 'Report of Last Year Service Individual Attendance':
            service_reports['last_year'] = group
            print(f"✅ Found last year service report: {group_name}")
        elif group_name == 'Report of Two Years Ago Service Individual Attendance':
            service_reports['two_years_ago'] = group
            print(f"✅ Found two years ago service report: {group_name}")
        
        # Group attendance reports (for Bible Study analysis)
        elif group_name == 'Report of Group Individual Attendance':
            group_reports['current'] = group
            print(f"✅ Found current group report: {group_name}")
        elif group_name == 'Report of Last Year Group Individual Attendance':
            group_reports['last_year'] = group
            print(f"✅ Found last year group report: {group_name}")
        elif group_name == 'Report of Two Years Ago Group Individual Attendance':
            group_reports['two_years_ago'] = group
            print(f"✅ Found two years ago group report: {group_name}")
    
    return service_reports, group_reports

def download_report_data(report_group, report_type):
    """Download report data from group URL (Code GP pattern)"""
    if not report_group:
        print(f"   ❌ No {report_type} report group provided")
        return None
        
    group_name = report_group.get('name', 'Unknown')
    print(f"   📥 Downloading {report_type} data: {group_name}")
    
    # Extract URL from group location fields (Code GP pattern)
    report_url = None
    for field in ['meeting_address', 'location', 'website']:
        if report_group.get(field) and 'http' in str(report_group[field]):
            report_url = str(report_group[field]).strip()
            print(f"   📍 Found URL in {field}: {report_url[:80]}...")
            break
    
    if not report_url:
        print(f"   ❌ No download URL found in group fields")
        print(f"   🔍 Available fields: {list(report_group.keys())}")
        return None
    
    try:
        print(f"   📡 Downloading report data...")
        response = requests.get(report_url, timeout=60)
        if response.status_code == 200:
            print(f"   ✅ Downloaded {len(response.content)} bytes")
            return response.content
        else:
            print(f"   ❌ HTTP Error {response.status_code}")
            return None
    except Exception as e:
        print(f"   ❌ Download failed: {e}")
        return None

def parse_column_header(header):
    """Parse column headers like '10:30 AM Morning Prayer 14/01' or '8:30 AM 02/06/2024' (Code AX pattern)"""
    
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
        service_name = re.sub(r'\d{1,2}/\d{1,2}/\d{4}\s*\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
        service_name = service_name.strip()
        
        return {
            'time': time_str,
            'day': day,
            'month': month,
            'year': year,
            'service_name': service_name,
            'original_header': header,
            'date_format': date_format,
            'normalized_time': time_str.replace(' ', '').replace('AM', '').replace('PM', '')
        }
    
    # Try DD/MM format (last year data)
    date_match = re.search(r'(\d{1,2})/(\d{1,2})(?:\s|$)', header)
    if date_match:
        day = int(date_match.group(1))
        month = int(date_match.group(2))
        date_format = 'short'
        
        # Extract service name (everything between time and date)
        service_name = header
        service_name = re.sub(r'\d{1,2}:\d{2}\s*(AM|PM)', '', service_name, flags=re.IGNORECASE)
        service_name = re.sub(r'\d{1,2}/\d{1,2}(?:\s|$)', '', service_name)
        service_name = service_name.strip()
        
        return {
            'time': time_str,
            'day': day,
            'month': month,
            'year': None,  # Year will be assigned later
            'service_name': service_name,
            'original_header': header,
            'date_format': date_format,
            'normalized_time': time_str.replace(' ', '').replace('AM', '').replace('PM', '')
        }
    
    return None

def parse_service_attendance_data(data_content, year_label, target_year):
    """Parse service attendance data using Code AX approach - count weekly attendance"""
    if not data_content:
        return {}
        
    print(f"📊 Parsing {year_label} service attendance using Code AX method...")
    
    try:
        html_content = data_content.decode('utf-8')
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find table with attendance data
        table = soup.find('table')
        if not table:
            print(f"   ❌ No table found in {year_label} service data")
            return {}
        
        rows = table.find_all('tr')
        if not rows:
            print(f"   ❌ No rows found in {year_label} service table")
            return {}
        
        # Parse header to find service columns with dates
        header_row = rows[0]
        headers = [cell.get_text(strip=True) for cell in header_row.find_all(['th', 'td'])]
        
        # Find name columns
        first_name_col = None
        last_name_col = None
        for i, header in enumerate(headers):
            if 'first name' in header.lower():
                first_name_col = i
            elif 'last name' in header.lower():
                last_name_col = i
        
        if first_name_col is None or last_name_col is None:
            print(f"   ❌ Could not find name columns in {year_label} data")
            return {}
        
        # Parse service columns
        service_columns = []
        for i, header in enumerate(headers):
            if header.lower() in ['first name', 'last name', 'category', 'email', 'phone']:
                continue
                
            parsed = parse_column_header(header)
            if not parsed:
                continue
            
            # Determine year
            if parsed.get('date_format') == 'full':
                service_year = parsed['year']
            else:
                service_year = target_year
            
            try:
                service_date = datetime(service_year, parsed['month'], parsed['day'])
                
                # Only include Sundays
                if service_date.weekday() != 6:  # Sunday = 6
                    continue
                
                # Only include services from the target year
                if service_date.year != target_year:
                    continue
                
                service_columns.append({
                    'header': header,
                    'column_index': i,
                    'date': service_date,
                    'time': parsed['time'],
                    'normalized_time': parsed['normalized_time']
                })
            except ValueError:
                continue
        
        print(f"   ✅ Found {len(service_columns)} valid Sunday services for {target_year}")
        
        if not service_columns:
            print(f"   ❌ No valid Sunday services found for {target_year}")
            return {}
        
        # Group services by date (Sunday) to handle multiple services per day
        sundays_data = {}
        
        for service in service_columns:
            date = service['date']
            time = service['normalized_time']
            column_index = service['column_index']
            
            if date not in sundays_data:
                sundays_data[date] = {
                    '8:30': set(),
                    '10:30': set(), 
                    '6:30': set(),
                    'all_services': set()
                }
            
            # Get all attendees for this specific service (using full names as identifiers)
            attendees = set()
            for row in rows[1:]:  # Skip header
                cells = row.find_all(['td', 'th'])
                if len(cells) > max(column_index, first_name_col, last_name_col):
                    cell_value = cells[column_index].get_text(strip=True)
                    if cell_value.upper() == 'Y':
                        # Get person's full name
                        first_name = cells[first_name_col].get_text(strip=True)
                        last_name = cells[last_name_col].get_text(strip=True)
                        full_name = f"{first_name} {last_name}".strip()
                        if full_name:
                            attendees.add(full_name)
            
            # Store attendees by service time
            if '8:30' in time or '8.30' in time:
                sundays_data[date]['8:30'].update(attendees)
            elif '10:30' in time or '10.30' in time:
                sundays_data[date]['10:30'].update(attendees)
            elif '6:30' in time or '6.30' in time:
                sundays_data[date]['6:30'].update(attendees)
            
            # Track all unique attendees for this Sunday
            sundays_data[date]['all_services'].update(attendees)
        
        # Convert to list format for Code AX-style processing
        attendance_data = []
        for date, data in sundays_data.items():
            # Calculate combined 10:30 + 6:30 (avoiding double counting)
            combined_1030_630 = data['10:30'].union(data['6:30'])
            
            sunday_record = {
                'date': date,
                '8:30 AM': len(data['8:30']),
                '10:30 AM': len(data['10:30']),
                '6:30 PM': len(data['6:30']),
                'combined_1030_630': len(combined_1030_630),
                'overall': len(data['all_services'])
            }
            attendance_data.append(sunday_record)
        
        # Sort by date
        attendance_data.sort(key=lambda x: x['date'])
        
        print(f"   📊 {year_label} - Found {len(attendance_data)} Sundays")
        
        # Calculate averages following Code AX approach
        if not attendance_data:
            return {}
        
        # Extract values for each service (only non-zero values)
        values_830 = [record['8:30 AM'] for record in attendance_data if record['8:30 AM'] > 0]
        values_1030 = [record['10:30 AM'] for record in attendance_data if record['10:30 AM'] > 0]
        values_630 = [record['6:30 PM'] for record in attendance_data if record['6:30 PM'] > 0]
        values_combined_1030_630 = [record['combined_1030_630'] for record in attendance_data if record['combined_1030_630'] > 0]
        values_overall = [record['overall'] for record in attendance_data if record['overall'] > 0]
        
        # Calculate averages
        avg_830 = np.mean(values_830) if values_830 else 0
        avg_1030 = np.mean(values_1030) if values_1030 else 0
        avg_630 = np.mean(values_630) if values_630 else 0
        avg_combined_1030_630 = np.mean(values_combined_1030_630) if values_combined_1030_630 else 0
        avg_all_services = np.mean(values_overall) if values_overall else 0
        
        print(f"   📊 {year_label} - 8:30 average: {avg_830:.1f}")
        print(f"   📊 {year_label} - 10:30 average: {avg_1030:.1f}")
        print(f"   📊 {year_label} - 6:30 average: {avg_630:.1f}")
        print(f"   📊 {year_label} - Combined 10:30 + 6:30 average: {avg_combined_1030_630:.1f}")
        print(f"   📊 {year_label} - All services combined average: {avg_all_services:.1f}")
        
        return {
            'combined_annual_average': avg_combined_1030_630,
            'all_services_average': avg_all_services,
            '8:30_average': avg_830,
            '10:30_average': avg_1030,
            '6:30_average': avg_630
        }
        
    except Exception as e:
        print(f"   ❌ Error parsing {year_label} service data: {e}")
        return {}

def parse_group_attendance_data(data_content, year_label, year):
    """Parse group attendance data for Bible Study groups"""
    if not data_content:
        return {}
        
    print(f"📚 Parsing {year_label} group attendance for Bible Study groups...")
    
    try:
        html_content = data_content.decode('utf-8')
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find all table rows
        rows = soup.find_all('tr')
        
        bible_study_groups = {}
        current_group = None
        bible_study_keywords = ['bible study', 'youth group', 'kids club', 'international food & friends', 'iff']
        
        for row in rows:
            # Check if this is a group header (special formatting)
            style = row.get('style', '')
            cells = row.find_all(['td', 'th'])
            
            if not cells:
                continue
                
            cell_text = cells[0].get_text(strip=True) if cells else ""
            
            # Group header detection (white text on dark background)
            is_group_header = False
            if 'background' in style and ('black' in style or 'dark' in style or '#' in style):
                is_group_header = True
            elif 'color' in style and 'white' in style:
                is_group_header = True
            elif len(cells) == 1 and cell_text and not any(char.isdigit() for char in cell_text):
                is_group_header = True
            
            if is_group_header and cell_text:
                # Check if this is a Bible Study related group
                if any(keyword in cell_text.lower() for keyword in bible_study_keywords):
                    current_group = cell_text
                    bible_study_groups[current_group] = {}
                    print(f"   📚 Found Bible Study group: {current_group}")
                else:
                    current_group = None
            
            # Person attendance data
            elif current_group and len(cells) >= 2:
                person_name = cells[0].get_text(strip=True)
                if person_name and not person_name.lower() in ['name', 'person', '']:
                    # Count attendance (Y marks)
                    attendance_count = 0
                    for cell in cells[1:]:  # Skip name column
                        if cell.get_text(strip=True).upper() == 'Y':
                            attendance_count += 1
                    
                    # Only include if attended at least twice
                    if attendance_count >= 2:
                        bible_study_groups[current_group][person_name] = attendance_count
        
        # Calculate unique people count across all groups (avoid double counting)
        all_people = set()
        for group_name, members in bible_study_groups.items():
            all_people.update(members.keys())
        
        bible_study_groups['_unique_people_count'] = len(all_people)
        bible_study_groups['_total_groups'] = len([g for g in bible_study_groups.keys() if not g.startswith('_')])
        
        print(f"   📊 {year_label} - Found {len(all_people)} unique people across {bible_study_groups['_total_groups']} groups")
        
        return bible_study_groups
        
    except Exception as e:
        print(f"   ❌ Error parsing {year_label} group data: {e}")
        return {}

def count_bible_study_groups_from_attendance_data(attendance_data):
    """Count Bible Study groups from attendance data (excluding Youth Group and Kids Club)"""
    print("\n📚 Counting Bible Study groups...")
    
    exclude_keywords = ['youth group', 'kids club']
    current_year = datetime.now().year
    
    group_counts = {}
    
    for year, year_data in attendance_data.items():
        total_groups = 0
        included_groups = []
        excluded_groups = []
        
        for group_name, members in year_data.items():
            # Skip metadata fields
            if group_name.startswith('_'):
                continue
                
            group_name_lower = group_name.lower()
            
            # Exclude Youth Group and Kids Club from count, but include IFF
            if any(exclude in group_name_lower for exclude in exclude_keywords):
                excluded_groups.append(group_name)
            else:
                total_groups += 1
                included_groups.append(group_name)
        
        group_counts[year] = total_groups
        
        print(f"   📊 {year}: {total_groups} groups")
        for group in included_groups:
            print(f"      ✅ {group}")
        for group in excluded_groups:
            print(f"      ❌ {group} (excluded from count)")
    
    return group_counts

def find_church_day_away_groups():
    """Find Church Day Away groups for each year"""
    print("\n🏕️ Searching for Church Day Away groups...")
    
    response = make_request('groups/getAll', {'page_size': 1000})
    if not response:
        return {}
    
    groups = response['groups'].get('group', [])
    if not isinstance(groups, list):
        groups = [groups] if groups else []
    
    day_away_groups = {}
    current_year = datetime.now().year
    
    for year in [current_year - 2, current_year - 1, current_year]:
        target_name = f"Church Day Away {year}"
        
        for group in groups:
            if group.get('name', '') == target_name:
                day_away_groups[year] = group
                print(f"   ✅ Found: {target_name}")
                break
        
        if year not in day_away_groups:
            print(f"   ❌ Not found: {target_name}")
    
    return day_away_groups

def get_church_day_away_attendance(day_away_groups):
    """Get attendance numbers for Church Day Away events"""
    attendance_data = {}
    
    for year, group in day_away_groups.items():
        group_id = group.get('id')
        
        # Get group members with fields parameter
        response = make_request('groups/getInfo', {
            'id': group_id,
            'fields': ['people']
        })
        
        if response:
            group_info = response.get('group', {})
            # Handle case where API returns a list instead of single object
            if isinstance(group_info, list):
                group_info = group_info[0] if group_info else {}
                
            if group_info.get('people') and group_info['people'].get('person'):
                people = group_info['people']['person']
                if not isinstance(people, list):
                    people = [people] if people else []
                
                attendance_data[year] = len(people)
                print(f"   📊 {year}: {len(people)} attendees")
            else:
                attendance_data[year] = 0
                print(f"   📊 {year}: 0 attendees (no people data)")
        else:
            attendance_data[year] = 0
            print(f"   📊 {year}: 0 attendees (API error)")
    
    return attendance_data

def get_user_targets():
    """Get user input for targets with defaults"""
    current_year = datetime.now().year
    next_year = current_year + 1
    four_years_time = current_year + 4
    
    print(f"\n🎯 TARGET SETTINGS")
    print("="*50)
    
    # Bible Study Group Attendance targets
    print(f"📚 Bible Study Group Attendance Targets:")
    print(f"   Default {next_year} target: Add 5 percentage points to {current_year}")
    print(f"   Default {four_years_time} target: Add 10 percentage points to {current_year}")
    
    while True:
        choice = input(f"\nUse default Bible Study attendance targets? (y/n): ").strip().lower()
        if choice in ['y', 'yes']:
            bible_study_next = 5
            bible_study_four = 10
            break
        elif choice in ['n', 'no']:
            try:
                bible_study_next = float(input(f"Enter percentage points to add for {next_year}: "))
                bible_study_four = float(input(f"Enter percentage points to add for {four_years_time}: "))
                break
            except ValueError:
                print("Please enter valid numbers.")
        else:
            print("Please enter 'y' or 'n'.")
    
    # Bible Study Groups count targets
    print(f"\n📚 Bible Study Groups Count Targets:")
    print(f"   Default {next_year} target: 1 more group than current")
    print(f"   Default {four_years_time} target: 4 more groups than current")
    
    while True:
        choice = input(f"\nUse default Bible Study group count targets? (y/n): ").strip().lower()
        if choice in ['y', 'yes']:
            groups_next = 1
            groups_four = 4
            break
        elif choice in ['n', 'no']:
            try:
                groups_next = int(input(f"Enter additional groups for {next_year}: "))
                groups_four = int(input(f"Enter additional groups for {four_years_time}: "))
                break
            except ValueError:
                print("Please enter valid numbers.")
        else:
            print("Please enter 'y' or 'n'.")
    
    # Church Day Away targets
    print(f"\n🏕️ Church Day Away Attendance Targets:")
    print(f"   Default {next_year} target: Add 5 percentage points to {current_year}")
    print(f"   Default {four_years_time} target: Add 10 percentage points to {current_year}")
    
    while True:
        choice = input(f"\nUse default Church Day Away targets? (y/n): ").strip().lower()
        if choice in ['y', 'yes']:
            day_away_next = 5
            day_away_four = 10
            break
        elif choice in ['n', 'no']:
            try:
                day_away_next = float(input(f"Enter percentage points to add for {next_year}: "))
                day_away_four = float(input(f"Enter percentage points to add for {four_years_time}: "))
                break
            except ValueError:
                print("Please enter valid numbers.")
        else:
            print("Please enter 'y' or 'n'.")
    
    return {
        'bible_study_attendance': {'next': bible_study_next, 'four': bible_study_four},
        'bible_study_groups': {'next': groups_next, 'four': groups_four},
        'day_away': {'next': day_away_next, 'four': day_away_four}
    }

def create_dashboard(attendance_percentages, group_counts, day_away_percentages, targets):
    """Create the dashboard with three sections (ONLY the top chart semantics changed)"""
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    year_labels = [str(year) for year in years]
    
    # Create subplots
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=(
            "Bible Study Group Attendance (% of Church Regulars at Year Start)",
            "Number of Bible Study Groups",
            "Church Day Away Attendance (% of Annual Average Combined 10:30 + 6:30 Attendance)"
        ),
        vertical_spacing=0.15
    )
    
    # Colors
    historical_color = '#1f77b4'  # Blue
    target_color = '#ff7f0e'      # Orange
    
    # Section 1: Bible Study Group Attendance (now % of Regulars at Year Start)
    historical_attendance = [attendance_percentages.get(year, 0) for year in years]
    current_attendance = historical_attendance[-1] if historical_attendance else 0
    
    # Targets: still relative to current percentage (as before)
    next_attendance = current_attendance + targets['bible_study_attendance']['next']
    four_attendance = current_attendance + targets['bible_study_attendance']['four']
    
    fig.add_trace(
        go.Bar(
            x=year_labels,
            y=historical_attendance,
            name='Actual',
            marker_color=historical_color,
            text=[f"{round(val)}%" for val in historical_attendance],
            textposition='outside',
            showlegend=True
        ),
        row=1, col=1
    )
    fig.add_trace(
        go.Bar(
            x=[str(current_year + 1), str(current_year + 4)],
            y=[next_attendance, four_attendance],
            name='Target',
            marker_color=target_color,
            text=[f"{round(next_attendance)}%", f"{round(four_attendance)}%"],
            textposition='outside',
            showlegend=True
        ),
        row=1, col=1
    )
    
    # Section 2: unchanged
    historical_groups = [group_counts.get(year, 0) for year in years]
    current_groups = historical_groups[-1] if historical_groups else 0
    next_groups = current_groups + targets['bible_study_groups']['next']
    four_groups = current_groups + targets['bible_study_groups']['four']
    fig.add_trace(
        go.Bar(
            x=year_labels, y=historical_groups, name='Actual',
            marker_color=historical_color, text=[f"{val}" for val in historical_groups],
            textposition='outside', showlegend=False
        ),
        row=2, col=1
    )
    fig.add_trace(
        go.Bar(
            x=[str(current_year + 1), str(current_year + 4)], y=[next_groups, four_groups],
            name='Target', marker_color=target_color,
            text=[f"{next_groups}", f"{four_groups}"], textposition='outside',
            showlegend=False
        ),
        row=2, col=1
    )
    
    # Section 3: unchanged
    historical_day_away = [day_away_percentages.get(year, 0) for year in years]
    current_day_away = historical_day_away[-1] if historical_day_away else 0
    next_day_away = current_day_away + targets['day_away']['next']
    four_day_away = current_day_away + targets['day_away']['four']
    fig.add_trace(
        go.Bar(
            x=year_labels, y=historical_day_away, name='Actual',
            marker_color=historical_color,
            text=[f"{round(val)}%" for val in historical_day_away],
            textposition='outside', showlegend=False
        ),
        row=3, col=1
    )
    fig.add_trace(
        go.Bar(
            x=[str(current_year + 1), str(current_year + 4)], y=[next_day_away, four_day_away],
            name='Target', marker_color=target_color,
            text=[f"{round(next_day_away)}%", f"{round(four_day_away)}%"],
            textposition='outside', showlegend=False
        ),
        row=3, col=1
    )
    
    fig.update_layout(
        title={'text': "📖 Applying the Word Dashboard", 'x': 0.5, 'xanchor': 'center', 'font': {'size': 24, 'color': 'black'}},
        height=1200, paper_bgcolor='white', plot_bgcolor='white',
        font=dict(family="Arial, sans-serif", size=12, color='black'),
        margin=dict(l=80, r=80, t=120, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=12))
    )
    fig.update_xaxes(title_text="Year", row=3, col=1, color='black')
    fig.update_yaxes(title_text="Percentage (% of Church Regulars at Year Start)", row=1, col=1, color='black')
    fig.update_yaxes(title_text="Number of Groups", row=2, col=1, color='black')
    fig.update_yaxes(title_text="Percentage (%)", row=3, col=1, color='black')

    max_attendance = max(historical_attendance + [next_attendance, four_attendance]) if historical_attendance else 100
    max_groups = max(historical_groups + [next_groups, four_groups]) if historical_groups else 10
    max_day_away = max(historical_day_away + [next_day_away, four_day_away]) if historical_day_away else 100
    fig.update_yaxes(range=[0, max_attendance * 1.15], row=1, col=1)
    fig.update_yaxes(range=[0, max_groups * 1.15], row=2, col=1)
    fig.update_yaxes(range=[0, max_day_away * 1.15], row=3, col=1)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
    fig.update_xaxes(showgrid=False, color='black')
    return fig

def main():
    print("🚀 Starting Applying the Word Dashboard analysis...")

    current_year = datetime.now().year
    years_to_process = [
        (current_year, 'current'),
        (current_year - 1, 'last_year'),
        (current_year - 2, 'two_years_ago')
    ]

    # ------------------------------------------------------------
    # 1) Attendance reports (service + group)
    # ------------------------------------------------------------
    service_reports, group_reports = find_attendance_reports()

    # ------------------------------------------------------------
    # 2) Group attendance parsing (Bible Study unique attendees, per year)
    # ------------------------------------------------------------
    attendance_data = {}
    for year, key in years_to_process:
        if key in group_reports:
            data = download_report_data(group_reports[key], "group attendance")
            if data:
                attendance_data[year] = parse_group_attendance_data(data, str(year), year)
            else:
                attendance_data[year] = {}
        else:
            attendance_data[year] = {}
            print(f"   ⚠️ No {key} group report available")

    # ------------------------------------------------------------
    # 3) Service attendance parsing (used for Day Away % denominators only)
    # ------------------------------------------------------------
    congregation_attendance = {}
    for year, key in years_to_process:
        if key in service_reports:
            data = download_report_data(service_reports[key], "service attendance")
            if data:
                parsed = parse_service_attendance_data(data, str(year), year)
                if parsed and parsed.get('all_services_average', 0) > 0:
                    congregation_attendance[year] = parsed
                    print(f"   ✅ {year} service data: {parsed.get('all_services_average', 0):.1f} average attendance")
                else:
                    congregation_attendance[year] = {'all_services_average': 85, 'combined_annual_average': 85}
                    print(f"   ⚠️ {year} service parsing failed, using estimate of 85")
            else:
                congregation_attendance[year] = {'all_services_average': 85, 'combined_annual_average': 85}
                print(f"   ⚠️ {year} service download failed, using estimate of 85")
        else:
            congregation_attendance[year] = {'all_services_average': 85, 'combined_annual_average': 85}
            print(f"   ⚠️ No {key} service report available, using estimate of 85")

    # ------------------------------------------------------------
    # 4) People Category Change reports → reverse replay inputs
    # ------------------------------------------------------------
    print("\n🔄 Reconstructing 'regulars at year start' from People Category Change reports...")
    change_groups = find_category_change_reports()  # assumes you added this helper
    events_by_year = {}
    for year, key in years_to_process:
        if key in change_groups:
            raw = download_report_data(change_groups[key], "people category change")
            events_by_year[year] = parse_category_change_report(raw, f"{year}") if raw else []
        else:
            events_by_year[year] = []
            print(f"   ⚠️ No {key} category change report available")

    # ------------------------------------------------------------
    # 5) Robust "regulars today" (Congregation_ + RosteredMember_), paginated & with exclusions
    # ------------------------------------------------------------
    today_regulars = get_current_regular_ids()  # you replaced this with the robust, paginated, truthy-filters version

    # ------------------------------------------------------------
    # 6) Fetch minimal person meta for snapshot pruning (date_added, archived/deceased/status)
    # ------------------------------------------------------------
    people_meta = fetch_people_meta()  # you added this

    # ------------------------------------------------------------
    # 7) Reverse-replay + prune by existence (@ 1 Jan each year)
    # ------------------------------------------------------------
    regular_sets = reconstruct_regular_sets(today_regulars, events_by_year, people_meta)  # you replaced this

    # Quick denominators recap
    print("\n📌 DENOMINATOR SUMMARY (regulars @ 1 Jan, after pruning)")
    print(f"   {current_year - 2}: {len(regular_sets['start_two_years_ago'])}  |  "
          f"{current_year - 1}: {len(regular_sets['start_last_year'])}  |  "
          f"{current_year}: {len(regular_sets['start_this_year'])}")

    # ------------------------------------------------------------
    # 8) Top-chart: Bible Study % using those denominators
    # ------------------------------------------------------------
    attendance_percentages = compute_biblestudy_percentages_using_regulars(attendance_data, regular_sets)

    # ------------------------------------------------------------
    # 9) Middle chart: Bible Study group counts (excludes Youth/Kids)
    # ------------------------------------------------------------
    group_counts = count_bible_study_groups_from_attendance_data(attendance_data)

    # ------------------------------------------------------------
    # 10) Church Day Away counts & percentages (still based on combined 10:30+6:30 average)
    # ------------------------------------------------------------
    day_away_groups = find_church_day_away_groups()
    day_away_attendance = get_church_day_away_attendance(day_away_groups)

    day_away_percentages = {}
    for year in [current_year - 2, current_year - 1, current_year]:
        day_away_count = day_away_attendance.get(year, 0)
        congregation_data = congregation_attendance.get(year, {})
        congregation_avg = congregation_data.get('combined_annual_average', 85)
        percentage = (day_away_count / congregation_avg) * 100 if congregation_avg > 0 else 0
        day_away_percentages[year] = percentage
        print(f"🏕️ {year}: {day_away_count}/{congregation_avg:.0f} = {percentage:.1f}%")

    # ------------------------------------------------------------
    # 11) Targets (unchanged prompts)
    # ------------------------------------------------------------
    targets = get_user_targets()

    # ------------------------------------------------------------
    # 12) Build dashboard figure
    # ------------------------------------------------------------
    print("\n📊 Creating dashboard...")
    fig = create_dashboard(attendance_percentages, group_counts, day_away_percentages, targets)

    # Tweak labels to reflect the new top-chart denominator
    try:
        fig.update_yaxes(title_text="Percentage (% of Regulars at Year Start)", row=1, col=1)
        # Update first subplot title text if present
        if hasattr(fig.layout, "annotations") and len(fig.layout.annotations) > 0:
            fig.layout.annotations[0].text = "Bible Study Group Attendance (% of Regulars at Year Start)"
    except Exception as _e:
        print(f"   ⚠️ Could not retitle axes/subplot: {_e}")

    # ------------------------------------------------------------
    # 13) Save and open
    # ------------------------------------------------------------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"applying_the_word_dashboard_{timestamp}.png"

    print(f"💾 Saving dashboard as {filename}...")
    fig.write_image(filename, width=1400, height=1200, scale=2)
    print(f"✅ Dashboard saved successfully!")
    print(f"📁 File: {filename}")

    print(f"🚀 Opening dashboard...")
    try:
        abs_path = os.path.abspath(filename)
        webbrowser.open(f'file://{abs_path}')
        print(f"✅ Dashboard opened in default viewer")
    except Exception as e:
        print(f"⚠️ Could not auto-open dashboard: {e}")
        print(f"   Please manually open: {filename}")

    # ------------------------------------------------------------
    # 14) Summary (explicitly note the new denominator)
    # ------------------------------------------------------------
    print("\n📋 DASHBOARD SUMMARY")
    print("="*50)
    print(f"Bible Study Attendance (as % of regulars @ year start): {attendance_percentages.get(current_year, 0):.1f}%")
    print(f"Bible Study Groups: {group_counts.get(current_year, 0)} groups")
    print(f"Church Day Away Attendance: {day_away_percentages.get(current_year, 0):.1f}% of combined 10:30+6:30 congregation")

if __name__ == "__main__":
    main()
