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
    """Create the dashboard with three sections"""
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    year_labels = [str(year) for year in years]
    
    # Create subplots
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=(
            "Bible Study Group Attendance (% of Annual Average Congregation Attendance)",
            "Number of Bible Study Groups",
            "Church Day Away Attendance (% of Annual Average Combined 10:30 + 6:30 Attendance)"
        ),
        vertical_spacing=0.15
    )
    
    # Colors
    historical_color = '#1f77b4'  # Blue
    target_color = '#ff7f0e'      # Orange
    
    # Section 1: Bible Study Group Attendance Percentages
    historical_attendance = [attendance_percentages.get(year, 0) for year in years]
    current_attendance = historical_attendance[-1] if historical_attendance else 0
    
    # Calculate targets (add to current percentage)
    next_attendance = current_attendance + targets['bible_study_attendance']['next']
    four_attendance = current_attendance + targets['bible_study_attendance']['four']
    
    # Historical bars
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
    
    # Target bars
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
    
    # Section 2: Number of Bible Study Groups
    historical_groups = [group_counts.get(year, 0) for year in years]
    current_groups = historical_groups[-1] if historical_groups else 0
    
    # Calculate targets (add to current count)
    next_groups = current_groups + targets['bible_study_groups']['next']
    four_groups = current_groups + targets['bible_study_groups']['four']
    
    # Historical bars
    fig.add_trace(
        go.Bar(
            x=year_labels,
            y=historical_groups,
            name='Actual',
            marker_color=historical_color,
            text=[f"{val}" for val in historical_groups],
            textposition='outside',
            showlegend=False
        ),
        row=2, col=1
    )
    
    # Target bars
    fig.add_trace(
        go.Bar(
            x=[str(current_year + 1), str(current_year + 4)],
            y=[next_groups, four_groups],
            name='Target',
            marker_color=target_color,
            text=[f"{next_groups}", f"{four_groups}"],
            textposition='outside',
            showlegend=False
        ),
        row=2, col=1
    )
    
    # Section 3: Church Day Away Attendance Percentages
    historical_day_away = [day_away_percentages.get(year, 0) for year in years]
    current_day_away = historical_day_away[-1] if historical_day_away else 0
    
    # Calculate targets (add to current percentage)
    next_day_away = current_day_away + targets['day_away']['next']
    four_day_away = current_day_away + targets['day_away']['four']
    
    # Historical bars
    fig.add_trace(
        go.Bar(
            x=year_labels,
            y=historical_day_away,
            name='Actual',
            marker_color=historical_color,
            text=[f"{round(val)}%" for val in historical_day_away],
            textposition='outside',
            showlegend=False
        ),
        row=3, col=1
    )
    
    # Target bars
    fig.add_trace(
        go.Bar(
            x=[str(current_year + 1), str(current_year + 4)],
            y=[next_day_away, four_day_away],
            name='Target',
            marker_color=target_color,
            text=[f"{round(next_day_away)}%", f"{round(four_day_away)}%"],
            textposition='outside',
            showlegend=False
        ),
        row=3, col=1
    )
    
    # Update layout
    fig.update_layout(
        title={
            'text': "📖 Applying the Word Dashboard",
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 24, 'color': 'black'}
        },
        height=1200,
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(family="Arial, sans-serif", size=12, color='black'),
        margin=dict(l=80, r=80, t=120, b=80),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=12)
        )
    )
    
    # Update axes with extra room at top for labels
    fig.update_xaxes(title_text="Year", row=3, col=1, color='black')
    fig.update_yaxes(title_text="Percentage (% of Annual Average Congregation Attendance)", row=1, col=1, color='black')
    fig.update_yaxes(title_text="Number of Groups", row=2, col=1, color='black')
    fig.update_yaxes(title_text="Percentage (%)", row=3, col=1, color='black')
    
    # Add extra room at top for text labels in all sections
    # Get max values to set appropriate y-axis ranges
    max_attendance = max(historical_attendance + [next_attendance, four_attendance]) if historical_attendance else 100
    max_groups = max(historical_groups + [next_groups, four_groups]) if historical_groups else 10
    max_day_away = max(historical_day_away + [next_day_away, four_day_away]) if historical_day_away else 100
    
    fig.update_yaxes(range=[0, max_attendance * 1.15], row=1, col=1)  # 15% extra room at top
    fig.update_yaxes(range=[0, max_groups * 1.15], row=2, col=1)      # 15% extra room at top  
    fig.update_yaxes(range=[0, max_day_away * 1.15], row=3, col=1)    # 15% extra room at top
    
    # Add grid
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
    fig.update_xaxes(showgrid=False, color='black')
    
    return fig

def main():
    print("🚀 Starting Applying the Word Dashboard analysis...")
    
    # Step 1: Find all attendance reports
    service_reports, group_reports = find_attendance_reports()
    
    # Step 2: Download and parse group attendance data for Bible Study analysis
    attendance_data = {}
    current_year = datetime.now().year
    
    years_to_process = [
        (current_year, 'current'),
        (current_year - 1, 'last_year'),
        (current_year - 2, 'two_years_ago')
    ]
    
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
    
    # Step 3: Download and parse service attendance data for congregation size
    congregation_attendance = {}
    for year, key in years_to_process:
        if key in service_reports:
            data = download_report_data(service_reports[key], "service attendance")
            if data:
                parsed_data = parse_service_attendance_data(data, str(year), year)
                if parsed_data and parsed_data.get('all_services_average', 0) > 0:
                    congregation_attendance[year] = parsed_data
                    print(f"   ✅ {year} service data: {parsed_data.get('all_services_average', 0):.1f} average attendance")
                else:
                    congregation_attendance[year] = {'all_services_average': 85}
                    print(f"   ⚠️ {year} service parsing failed, using estimate of 85")
            else:
                congregation_attendance[year] = {'all_services_average': 85}
                print(f"   ⚠️ {year} service download failed, using estimate of 85")
        else:
            congregation_attendance[year] = {'all_services_average': 85}
            print(f"   ⚠️ No {key} service report available, using estimate of 85")

    # Step 4: Calculate Bible Study attendance percentages
    attendance_percentages = {}
    for year, year_data in attendance_data.items():
        unique_attendees = year_data.get('_unique_people_count', 0)
        year_service_data = congregation_attendance.get(year, {})
        all_services_avg = year_service_data.get('all_services_average', 0)
        
        if all_services_avg > 0:
            percentage = (unique_attendees / all_services_avg) * 100
        else:
            percentage = 0
        
        attendance_percentages[year] = percentage
        print(f"📊 {year}: {unique_attendees} unique people / {all_services_avg:.1f} avg attendance = {percentage:.1f}%")
    
    # Step 5: Count Bible Study groups from attendance data
    group_counts = count_bible_study_groups_from_attendance_data(attendance_data)
    
    # Step 6: Get Church Day Away data
    day_away_groups = find_church_day_away_groups()
    day_away_attendance = get_church_day_away_attendance(day_away_groups)
    
    # Step 7: Calculate Day Away percentages
    day_away_percentages = {}
    for year in [current_year - 2, current_year - 1, current_year]:
        day_away_count = day_away_attendance.get(year, 0)
        congregation_data = congregation_attendance.get(year, {})
        congregation_avg = congregation_data.get('combined_annual_average', 85)
        percentage = (day_away_count / congregation_avg) * 100 if congregation_avg > 0 else 0
        day_away_percentages[year] = percentage
        print(f"🏕️ {year}: {day_away_count}/{congregation_avg:.0f} = {percentage:.1f}%")
    
    # Step 8: Get user targets
    targets = get_user_targets()
    
    # Step 9: Create dashboard
    print("\n📊 Creating dashboard...")
    fig = create_dashboard(attendance_percentages, group_counts, day_away_percentages, targets)
    
    # Step 10: Save as PNG
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"applying_the_word_dashboard_{timestamp}.png"
    
    print(f"💾 Saving dashboard as {filename}...")
    fig.write_image(filename, width=1400, height=1200, scale=2)
    
    print(f"✅ Dashboard saved successfully!")
    print(f"📁 File: {filename}")
    
    # Step 11: Auto-open the PNG file
    print(f"🚀 Opening dashboard...")
    try:
        abs_path = os.path.abspath(filename)
        webbrowser.open(f'file://{abs_path}')
        print(f"✅ Dashboard opened in default viewer")
    except Exception as e:
        print(f"⚠️ Could not auto-open dashboard: {e}")
        print(f"   Please manually open: {filename}")
    
    # Summary
    print(f"\n📋 DASHBOARD SUMMARY")
    print("="*50)
    print(f"Bible Study Attendance: {attendance_percentages.get(current_year, 0):.1f}% of annual congregation attendance")
    print(f"Bible Study Groups: {group_counts.get(current_year, 0)} groups")
    print(f"Church Day Away Attendance: {day_away_percentages.get(current_year, 0):.1f}% of combined 10:30+6:30 congregation")

if __name__ == "__main__":
    main()
