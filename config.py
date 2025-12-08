
import os
from dotenv import load_dotenv

load_dotenv()
ELVANTO_API_KEY = os.getenv('ELVANTO_API_KEY')

# ST GEORGE'S MAGILL ANGLICAN CHURCH
# Strategic Plan 2025-2029 - Complete Data Reference
# Source: Draft_Strategic_Plan_20252029_v2.pdf

# ============================================
# OVERVIEW
# ============================================
# Vision: Know Christ. Make Christ Known.
# Planning Period: 2025-2029
# Annual Planning Cycle: Typically August/September

# ============================================
# PRAYER MEETING TARGETS
# ============================================

PRAYER_TARGETS = {
    'quarterly_prayer_meeting_attendance': {
        'description': 'Attendance at quarterly in-person prayer meetings',
        'baseline': {'year': 2025, 'value': 17},
        'targets': {
            2026: 20,
            2029: 25
        }
    },
    'zoom_prayer_meeting_attendance': {
        'description': 'Attendance at weekly Zoom prayer meetings',
        'baseline': {'year': 2025, 'value': 5},
        'targets': {
            2026: 7,
            2029: 10
        }
    },
    'prayer_training_hour_attendance': {
        'description': 'Attendance at prayer training sessions',
        'baseline': {'year': 2025, 'value': 0},
        'targets': {
            2026: 15,
            2029: 20
        },
        'status': 'Not yet tracked in dashboards'
    }
}

# ============================================
# NEWCOMERS & VISITOR TARGETS
# ============================================

NEWCOMERS_TARGETS = {
    'total_visitors_annual': {
        'description': 'Total number of visitors to church per year',
        'baseline': {'year': 2025, 'value': 150},
        'targets': {
            2026: 165,
            2029: 200
        }
    },
    'newcomers_lunch_attendance_pct': {
        'description': 'Percentage of visitors who join Newcomers Lunch',
        'baseline': {'year': 2025, 'value': 0.16},  # 16%
        'targets': {
            2026: 0.20,  # 20%
            2029: 0.25  # 25%
        }
    },
    'visitor_stay_ratio': {
        'description': 'Percentage of visitors who become congregation (stay ratio for whole church)',
        'baseline': {'year': 2025, 'value': 0.175},  # 17.5%
        'targets': {
            2026: 0.19,  # 19%
            2029: 0.22  # 22%
        },
        'note': '20% is considered good'
    }
}

# ============================================
# 10:30 CONGREGATION TARGETS
# ============================================

CONGREGATION_1030_TARGETS = {
    'average_attendance': {
        'description': 'Average weekly attendance at 10:30 service',
        'baseline': {'year': 2024, 'value': 68.53},
        'targets': {
            2025: 75.38,
            2026: 82.92,
            2027: 91.21,
            2028: 100.33,
            2029: 110.37
        },
        'note': '10% annual growth target'
    },
    'visitor_ratio': {
        'description': '10:30 visitor ratio as percentage of congregation',
        'baseline': {'year': 2025, 'value': 1.27},  # 127%
        'targets': {
            2026: 1.35,  # 135%
            2029: 1.50  # 150%
        }
    },
    'stay_numbers_1030': {
        'description': 'Number of new congregation members at 10:30',
        'baseline': {'year': 2025, 'value': 15},  # 15 ytd
        'targets': {
            2026: 22,
            2029: 30
        }
    },
    'stay_ratio': {
        'description': '10:30 stay ratio',
        'baseline': {'year': 2025, 'value': 0.23},  # 23%
        'targets': {
            2026: 0.23,  # 23% (maintain)
            2029: 0.24  # 24%
        },
        'note': '20% is considered good'
    }
}

# ============================================
# NEXTGEN (KIDS CLUB & YOUTH GROUP) TARGETS
# ============================================

NEXTGEN_TARGETS = {
    'church_kids_in_kids_club_pct': {
        'description': 'Percentage of church kids attending Kids Club',
        'baseline': {'year': 2025, 'value': 0.75},  # 75%
        'targets': {
            2026: 0.75,  # 75% (maintain)
            2029: 0.75  # 75% (maintain)
        }
    },
    'church_youth_in_youth_group_pct': {
        'description': 'Percentage of church youth attending Youth Group',
        'baseline': {'year': 2025, 'value': 0.69},  # 69%
        'targets': {
            2029: 0.75  # 75%
        }
    },
    'year6_transition_pct': {
        'description': 'Percentage of year 6 Kids Club kids transitioning to Youth Group',
        'baseline': {'year': 2025, 'value': 0.33},  # 33%
        'targets': {
            2026: 0.50,  # 50%
            2029: 0.50  # 50%
        }
    },
    'kids_club_unchurched_families': {
        'description': 'Kids Club members from unchurched families',
        'baseline': {'year': 2025, 'value': 2},
        'targets': {
            2026: 4,
            2029: 6
        }
    },
    'youth_group_unchurched_families': {
        'description': 'Youth Group members from unchurched families',
        'baseline': {'year': 2025, 'value': 1},
        'targets': {
            2026: 4,
            2029: 6
        }
    },
    'kids_youth_serving_pct': {
        'description': 'Percentage of kids and youth serving on teams/rosters',
        'baseline': {'year': 2025, 'value': 0.55},  # 55%
        'targets': {
            2026: 0.60,  # 60%
            2029: 0.75  # 75%
        }
    },
    'conversions': {
        'description': 'Number of conversions (kids and youth)',
        'baseline': {'year': 2025, 'value': 1},
        'targets': {
            2026: 2,
            2029: 5
        }
    }
}

# ============================================
# GOSPEL COURSE TARGETS (TASTE & SEE)
# ============================================

GOSPEL_COURSE_TARGETS = {
    'people_joining_gospel_course': {
        'description': 'Number of people joining a gospel course',
        'baseline': {'year': 2025, 'value': 17},
        'targets': {
            2026: 11,
            2029: 21
        }
    },
    'people_joining_followup_group': {
        'description': 'Number of people joining a follow up group',
        'baseline': {'year': 2025, 'value': 4},
        'targets': {
            2026: 4,
            2029: 6
        }
    },
    'formal_commitment_to_jesus': {
        'description': 'Number to have made a formal commitment to follow Jesus',
        'baseline': {'year': 2025, 'value': 5},
        'targets': {
            2026: 5,
            2029: 7
        }
    },
    'adult_baptisms': {
        'description': 'Number of adult baptisms',
        'baseline': {'year': 2025, 'value': 0},
        'targets': {
            2026: 2,
            2029: 4
        }
    },
    'followup_to_growth_group': {
        'description': 'Number of people joining a growth group from a follow up group',
        'baseline': {'year': 2025, 'value': 0},
        'targets': {
            2026: 3,
            2029: 4
        }
    }
}

# ============================================
# BIBLE STUDY & APPLYING THE WORD TARGETS
# ============================================

BIBLE_STUDY_TARGETS = {
    'bible_study_attendance_pct': {
        'description': 'Bible Study Group Attendance as % of regular attenders (Adult and Kids & Youth)',
        'baseline': {'year': 2025, 'value': 0.69},  # 69%
        'targets': {
            2026: 0.74,  # 74%
            2029: 0.79  # 79%
        }
    },
    'number_of_bible_study_groups': {
        'description': 'Number of Bible Study Groups (Adult, includes IFF)',
        'baseline': {'year': 2025, 'value': 10},
        'targets': {
            2026: 11,
            2029: 14
        }
    },
    'church_day_away_attendance_pct': {
        'description': 'Church Day Away % of avg attendance (Excluding 8:30)',
        'baseline': {'year': 2025, 'value': 0.71},  # 71%
        'targets': {
            2026: 0.75,  # 75%
            2029: 0.79  # 79%
        },
        'note': 'Target is +10% growth per planning cycle'
    }
}

# ============================================
# TRAINING & DEVELOPMENT TARGETS
# ============================================

TRAINING_TARGETS = {
    'yearly_training_coverage': {
        'description': 'Ministry areas covered by training and feedback',
        'baseline': {'year': 2025, 'areas': ['Sound', 'Music', 'Prayer']},
        'targets': {
            2026: ['Sound', 'Music', 'Prayer', 'Service Leading', 'Bible Reading'],
            2029: 'Complete suite of training'
        },
        'note': 'Regular huddles for service reflection, separate sound/tech training'
    }
}

# ============================================
# NCLS SURVEY DATA POINTS
# ============================================
# Note: NCLS surveys occur periodically (2020, 2022 data shown)
# Diocese averages provided for context

NCLS_TARGETS = {
    'belonging': {
        'strong_sense_belonging': {
            'description': 'Have a strong sense of belonging',
            'diocese_average': 0.93,  # 93%
            'data': {
                2020: 0.92,  # 92%
                2022: 0.96,  # 96%
                2029: 0.96  # Goal: 96%
            }
        },
        'strong_growing_belonging': {
            'description': 'Have a strong and growing sense of belonging',
            'diocese_average': 0.41,  # 41%
            'data': {
                2020: 0.55,  # 55%
                2022: 0.49,  # 49%
                2029: 0.55  # Goal: 55%
            }
        },
        'easy_make_friends': {
            'description': 'Found it easy to make friends in this parish',
            'diocese_average': 0.89,  # 89%
            'data': {
                2020: 0.88,  # 88%
                2022: 0.84,  # 84%
                2029: 0.88  # Goal: 88%
            }
        }
    },
    
    'pastoral_care': {
        'likely_follow_up_drifting': {
            'description': 'Are likely to follow up someone drifting away from church',
            'diocese_average': 0.47,  # 47%
            'data': {
                2020: 0.65,  # 65%
                2022: 0.56,  # 56%
                2029: 0.66  # Goal: 66%
            }
        },
        'certain_follow_up_drifting': {
            'description': 'Are certain to follow up someone drifting away from church',
            'diocese_average': 0.05,  # 5%
            'data': {
                2020: 0.10,  # 10%
                2022: 0.03,  # 3%
                2029: 0.10  # Goal: 10%
            }
        }
    },
    
    'service_and_welcome': {
        'informal_helping_3plus': {
            'description': '3 or more informal ways of helping others',
            'data': {
                2020: 0.57,  # 57%
                2022: 0.62  # 62%
            }
        },
        'personally_welcome_newcomers': {
            'description': 'Always or mostly personally seek to make new arrivals welcome',
            'data': {
                2020: 0.70,  # 70%
                2022: 0.61,  # 61%
                2029: 0.70  # Goal: 70%
            }
        }
    },
    
    'faith_and_inspiration': {
        'faith_influences_decisions': {
            'description': 'My faith influences decisions and actions in my life',
            'diocese_average': 0.91,  # 91%
            'data': {
                2022: 0.92,  # 92%
                2029: 0.95  # Goal: 95%
            }
        },
        'experience_inspiration': {
            'description': 'Always/Usually experience inspiration during services',
            'data': {
                2020: 0.64,  # 64%
                2022: 0.80,  # 80%
                2029: 0.85  # Goal: 85%
            }
        }
    },
    
    'giving_and_growth': {
        'give_10_percent_income': {
            'description': 'Give 10% of income',
            'diocese_average_13': 0.13,  # Diocese 13%
            'fiec_average_31': 0.31,  # FIEC 31%
            'data': {
                2020: 0.24,  # 24%
                2022: 0.25,  # 25%
                2029: 0.28  # Goal: 28%
            }
        },
        'experienced_much_growth': {
            'description': 'Experienced much growth through this church in last 12 months',
            'diocese_average_19': 0.19,  # Diocese 19%
            'fiec_average_55': 0.55,  # FIEC 55%
            'data': {
                2020: 0.26,  # 26%
                2022: 0.35,  # 35%
                2029: 0.45  # Goal: 45%
            }
        },
        'growth_understanding_worship': {
            'description': 'Growth in understanding God during worship services usually/always',
            'diocese_average_70': 0.70,  # Diocese 70%
            'fiec_average_90': 0.90,  # FIEC 90%
            'data': {
                2020: 0.90,  # 90%
                2022: 0.87,  # 87%
                2029: 0.90  # Goal: 90%
            }
        }
    }
}

# ============================================
# DASHBOARD MAPPING
# ============================================
# Which targets are tracked in which dashboards

DASHBOARD_MAPPING = {
    'prayer_dashboard': {
        'tracked': [
            'quarterly_prayer_meeting_attendance',
            'zoom_prayer_meeting_attendance'
        ],
        'not_yet_tracked': [
            'prayer_training_hour_attendance'
        ]
    },
    'newcomers_dashboard': {
        'tracked': [
            'newcomers_lunch_attendance_pct'
        ],
        'not_yet_tracked': []
    },
    'visitor_stay_dashboard': {
        'tracked': [],
        'planned': [
            'total_visitors_annual',
            'visitor_stay_ratio'
        ]
    },
    'congregation_1030_dashboard': {
        'tracked': [
            'average_attendance'
        ],
        'planned': [
            'visitor_ratio',
            'new_congregation_members',
            'stay_ratio'
        ]
    },
    'nextgen_dashboard': {
        'tracked': [
            'church_kids_in_kids_club_pct',
            'church_youth_in_youth_group_pct',
            'kids_youth_serving_pct',
            'conversions'
        ],
        'not_yet_tracked': [
            'year6_transition_pct',
            'kids_club_unchurched_families',
            'youth_group_unchurched_families'
        ]
    },
    'gospel_course_dashboard': {
        'tracked': [],
        'planned': [
            'people_joining_gospel_course',
            'people_joining_followup_group',
            'formal_commitment_to_jesus',
            'adult_baptisms',
            'followup_to_growth_group'
        ]
    },
    'bible_study_dashboard': {
        'tracked': [
            'bible_study_attendance_pct',
            'number_of_bible_study_groups',
            'church_day_away_attendance_pct'
        ],
        'not_yet_tracked': []
    }
}

# ============================================
# STRATEGIC AMBITIONS (2029)
# ============================================

STRATEGIC_AMBITIONS = {
    'gospel_reach': {
        'description': "Make Christ known in Magill and beyond",
        'goals': [
            "Regular baptisms as conversions increase 10% annually",
            "Gospel worker raised from within church for full-time ministry apprenticeship",
            "Members equipped and bold to share gospel personally"
        ]
    },
    'leadership_development': {
        'description': "Raise up new generation of leaders for Magill, Adelaide and beyond",
        'goals': [
            "Lots of people serving in word-based ministry for first time",
            "Training and discipleship programs established",
            "Leadership pipeline for local and wider church"
        ]
    },
    'community_and_growth': {
        'description': "Deep biblical loving community with vibrant serving teams",
        'goals': [
            "Reflect ethnic demographic of Magill and surrounding suburbs",
            "10:30 congregation grows sufficiently that space becomes a problem",
            "60%+ of congregation in small groups"
        ]
    }
}

# ============================================
# NOTES FOR DASHBOARD DEVELOPMENT
# ============================================

DEVELOPMENT_NOTES = """
1. Annual Updates: Review and update targets during planning cycle (Aug/Sep)

2. Baseline Years: 
   - Most metrics use 2025 as baseline year for consistency
   - Exception: 10:30 average_attendance uses 2024 baseline (68.53) with 2025 as first target (75.38)

3. Prorating: Current year targets should be prorated based on year progress for cumulative metrics

4. Data Sources:
   - Elvanto API: Prayer meetings, newcomers lunch, attendance, groups
   - NCLS Surveys: Conducted periodically, provide benchmarking data
   - Manual tracking: Some metrics may require new data collection processes

5. Dashboard Standards:
   - Light background (#ffffff, #f8fafc)
   - Target colors: Emerald (#10b981), Teal (#14b8a6), Cyan (#06b6d4)
   - Save to outputs/ directory
   - Include multi-year targets dynamically
   - Add annotation for not-yet-tracked targets

6. Config.py Structure:
   - Each target area has baseline_year, baseline_value, and targets dict
   - Helper function get_relevant_targets() filters to current/future years
   - Supports flexible year addition without code changes
"""

# ============================================
# VERSION HISTORY
# ============================================

VERSION_INFO = {
    'strategic_plan_version': 'Draft v2',
    'data_file_created': '2024-11-22',
    'last_updated': '2024-11-24',
    'notes': 'Updated with all strategic plan targets, standardized baseline years to 2025'
}
