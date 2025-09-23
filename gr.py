import pandas as pd
from collections import defaultdict

def parse_list(s):
    if not s or s == '[]':
        return []
    s = s.strip('[]')
    return [item.strip() for item in s.split(',')]

# --- Course Data ---
course_data = [
    {'course_id': '101', 'name': 'Data Structures', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F001', 'batch_size': 240, 'preferred_slot': None},
    {'course_id': '102', 'name': 'Digital Logic Design', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F002', 'batch_size': 240, 'preferred_slot': None},
    {'course_id': '103', 'name': 'Mathematics III', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 9, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'Yes', 'faculty_id': 'F003', 'batch_size': 240, 'preferred_slot': None},
    {'course_id': '104', 'name': 'OE - Environmental Studies', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 9, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'Yes', 'faculty_id': 'F010', 'batch_size': 240, 'preferred_slot': 'Wed-2'},
    {'course_id': '201', 'name': 'Operating Systems', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F004', 'batch_size': 180, 'preferred_slot': None},
    {'course_id': '202', 'name': 'Database Systems', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F005', 'batch_size': 180, 'preferred_slot': None},
    {'course_id': '203', 'name': 'Computer Networks', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F006', 'batch_size': 180, 'preferred_slot': None},
    {'course_id': '204', 'name': 'OE - Constitution of India', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 9, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'Yes', 'faculty_id': 'F011', 'batch_size': 180, 'preferred_slot': 'Thu-3'},
    {'course_id': '301', 'name': 'Machine Learning', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F007', 'batch_size': 120, 'preferred_slot': None},
    {'course_id': '302', 'name': 'Cloud Computing', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 6, 'practical_hours': 3, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F008', 'batch_size': 120, 'preferred_slot': None},
    {'course_id': '303', 'name': 'Major Project', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 0, 'practical_hours': 9, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'Yes', 'faculty_id': 'F009', 'batch_size': 120, 'preferred_slot': 'Fri-2,Fri-3'},
    {'course_id': '304', 'name': 'MDM - Seminar', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 9, 'hours_per_week': 9, 'theory_hours': 9, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Seminar', 'central': 'Yes', 'faculty_id': 'F012', 'batch_size': 120, 'preferred_slot': 'Sat-1'},
]
dfCourse = pd.DataFrame(course_data)
dfCourse['divisions_list'] = dfCourse['divisions'].apply(lambda x: x.split(',') if pd.notna(x) else [])
dfCourse['divisions_num'] = dfCourse['divisions_list'].apply(len)

# --- Faculty Data ---
faculty_data = [
    {'faculty_id': 'F001', 'name': 'Dr. A. Sharma', 'expertise': '[101] (Data Structures)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Mon-2]', 'consecutive_limit': 6},
    {'faculty_id': 'F002', 'name': 'Prof. B. Patil', 'expertise': '[102] (Digital Logic)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[]', 'consecutive_limit': 6},
    {'faculty_id': 'F003', 'name': 'Dr. C. Mehta', 'expertise': '[103] (Mathematics III)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Tue-2]', 'consecutive_limit': 6},
    {'faculty_id': 'F004', 'name': 'Prof. D. Kulkarni', 'expertise': '[201] (Operating Sys.)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[]', 'consecutive_limit': 6},
    {'faculty_id': 'F005', 'name': 'Dr. E. Iyer', 'expertise': '[202] (Database Sys.)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Fri-3]', 'consecutive_limit': 6},
    {'faculty_id': 'F006', 'name': 'Dr. F. Singh', 'expertise': '[203] (Comp. Networks)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[]', 'consecutive_limit': 6},
    {'faculty_id': 'F007', 'name': 'Prof. G. Deshmukh', 'expertise': '[301] (Machine Learning)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Wed-1]', 'consecutive_limit': 6},
    {'faculty_id': 'F008', 'name': 'Dr. H. Verma', 'expertise': '[302] (Cloud Computing)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[]', 'consecutive_limit': 6},
    {'faculty_id': 'F009', 'name': 'Prof. I. Rao', 'expertise': '[303] (Major Project)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Fri-2, Fri-3]', 'consecutive_limit': 6},
    {'faculty_id': 'F010', 'name': 'Dr. J. Joshi', 'expertise': '[104] (OE - EVS)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Wed-2]', 'consecutive_limit': 6},
    {'faculty_id': 'F011', 'name': 'Prof. K. Banerjee', 'expertise': '[204] (OE - Constitution)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Thu-3]', 'consecutive_limit': 6},
    {'faculty_id': 'F012', 'name': 'Dr. L. Reddy', 'expertise': '[304] (MDM - Seminar)', 'max_hours_per_week': 40, 'unavailable_slots': '[]', 'preferred_slots': '[Sat-1]', 'consecutive_limit': 6},
]
dfFaculty = pd.DataFrame(faculty_data)
dfFaculty['unavailable_slots'] = dfFaculty['unavailable_slots'].apply(parse_list)
dfFaculty['preferred_slots'] = dfFaculty['preferred_slots'].apply(parse_list)

# --- Room Data ---
room_data = [
    {'room_id': 'R101', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R102', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R201', 'capacity': 60, 'type': 'lab', 'available_slots': None},
    {'room_id': 'R202', 'capacity': 60, 'type': 'lab', 'available_slots': None},
    {'room_id': 'CL1', 'capacity': 60, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'CL2', 'capacity': 60, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'AUD1', 'capacity': 120, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R301', 'capacity': 25, 'type': 'theory', 'available_slots': None},
    {'room_id': 'AUD2', 'capacity': 240, 'type': 'theory', 'available_slots': None},
    {'room_id': 'AUD3', 'capacity': 240, 'type': 'theory', 'available_slots': None},
    {'room_id': 'AUD4', 'capacity': 240, 'type': 'theory', 'available_slots': None},
    {'room_id': 'LAB_BIG', 'capacity': 240, 'type': 'lab', 'available_slots': None},
    {'room_id': 'LAB2', 'capacity': 240, 'type': 'lab', 'available_slots': None},
    {'room_id': 'LAB3', 'capacity': 240, 'type': 'lab', 'available_slots': None},
]
dfRoom = pd.DataFrame(room_data)

# --- Constants ---
COLLEGE_START = 10*60 + 15  # 10:15 AM
COLLEGE_END = 17*60 + 30  # 5:30 PM
THEORY_DURATION = 60  # 1 hour
WORKING_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
BREAKS = [
    {"name": "Lunch Break", "start": 12*60 + 15, "duration": 60},  # 12:15–13:15
    {"name": "Tea Break", "start": 15*60 + 15, "duration": 15},  # 3:15–3:30
]

def minutes_to_time(m):
    return f"{m//60:02d}:{m%60:02d}"

def generate_slots():
    slots = []
    t = COLLEGE_START
    while t < COLLEGE_END:
        brk = next((br for br in BREAKS if br["start"] == t), None)
        if brk:
            br_start = minutes_to_time(brk["start"])
            br_end = minutes_to_time(brk["start"] + brk["duration"])
            slots.append(f"{br_start}-{br_end}")
            t += brk["duration"]
        else:
            start = minutes_to_time(t)
            end = minutes_to_time(t + THEORY_DURATION)
            slots.append(f"{start}-{end}")
            t += THEORY_DURATION
    return slots

def create_all_timetables(dfCourse):
    timetables = {}
    slots = generate_slots()
    grouped = dfCourse.groupby("year")["divisions_num"].max()
    div_map = 'ABCD'
    for year, max_divs in grouped.items():
        for div in range(1, max_divs + 1):
            key = f"Year{year}_Div{div_map[div - 1]}"
            df = pd.DataFrame(
                [["" for _ in slots] for _ in WORKING_DAYS],
                index=WORKING_DAYS,
                columns=slots
            )
            timetables[key] = df
    return timetables

timetables = create_all_timetables(dfCourse)

# --- Slot map ---
slots = generate_slots()
lecture_slot_times = [s for s in slots if not s.startswith('12:15') and not s.startswith('15:15')]
slot_number_map = {}
for day in WORKING_DAYS:
    for slot_num, slot_time in enumerate(lecture_slot_times, 1):
        key = f"{day}-{slot_num}"
        slot_number_map[key] = slot_time

# --- Prepare sessions ---
sessions = []
for idx, row in dfCourse.iterrows():
    year = row['year']
    divs = row['divisions_list']
    faculty = row['faculty_id']
    name = row['name']
    central = row['central'] == 'Yes'
    theory_h = row['theory_hours']
    prac_h = row['practical_hours']
    preferred = row['preferred_slot']
    batch_size = row['batch_size']
    preferred_list = [p.strip() for p in preferred.split(',')] if preferred else []
    if central:
        # For theory
        if theory_h > 0:
            max_cap = dfRoom[dfRoom['type'] == 'theory']['capacity'].max()
            group_size = max_cap // 60
            if group_size == 0:
                print(f"No suitable room for theory {name}")
                continue
            div_groups = [divs[i:i+group_size] for i in range(0, len(divs), group_size)]
            for i in range(theory_h):
                for group in div_groups:
                    sess_pref = preferred_list[i] if i < len(preferred_list) else None
                    sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': group, 'central': True, 'room_type': 'theory', 'batch_size': len(group)*60, 'preferred': sess_pref})
        # For practical
        if prac_h > 0:
            max_cap = dfRoom[dfRoom['type'].str.contains('lab')]['capacity'].max()
            group_size = max_cap // 60
            if group_size == 0:
                print(f"No suitable room for practical {name}")
                continue
            div_groups = [divs[i:i+group_size] for i in range(0, len(divs), group_size)]
            for i in range(prac_h):
                for group in div_groups:
                    sess_pref = preferred_list[i + theory_h] if i + theory_h < len(preferred_list) else None
                    sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': group, 'central': True, 'room_type': 'lab', 'batch_size': len(group)*60, 'preferred': sess_pref})
    else:
        for div in divs:
            pref_idx = 0
            for i in range(theory_h):
                sess_pref = preferred_list[pref_idx] if pref_idx < len(preferred_list) else None
                pref_idx += 1
                sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'theory', 'batch_size': 60, 'preferred': sess_pref})
            for i in range(prac_h):
                sess_pref = preferred_list[pref_idx] if pref_idx < len(preferred_list) else None
                pref_idx += 1
                sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'lab', 'batch_size': 60, 'preferred': sess_pref})

# --- Sort sessions ---
sessions = sorted(sessions, key=lambda s: (s['preferred'] is None, -len(s['divs'])))

# --- Scheduling structures ---
faculty_schedule = defaultdict(list)
room_schedule = defaultdict(list)
faculty_hours = defaultdict(int)

# --- Schedule sessions ---
for session in sessions:
    possible_slots = []
    faculty = session['faculty']
    max_hours = dfFaculty[dfFaculty['faculty_id'] == faculty]['max_hours_per_week'].item()
    consecutive_limit = dfFaculty[dfFaculty['faculty_id'] == faculty]['consecutive_limit'].item()
    unavailable = dfFaculty[dfFaculty['faculty_id'] == faculty]['unavailable_slots'].item()
    room_type = session['room_type']
    if room_type == 'Seminar':
        room_type = 'theory'
    for day in WORKING_DAYS:
        day_hours = len([s for d, s in faculty_schedule[faculty] if d == day])
        if day_hours >= 6:
            continue
        for slot_num in range(1, 7):
            key = f"{day}-{slot_num}"
            slot_time = slot_number_map[key]
            if session['preferred'] and session['preferred'] != key:
                continue
            if key in unavailable:
                continue
            if any(d == day and s == slot_time for d, s in faculty_schedule[faculty]):
                continue
            if faculty_hours[faculty] + 1 > max_hours:
                continue
            day_slots = [s for d, s in faculty_schedule[faculty] if d == day]
            day_indices = sorted([lecture_slot_times.index(s) for s in day_slots])
            new_index = lecture_slot_times.index(slot_time)
            day_indices.append(new_index)
            day_indices = sorted(day_indices)
            curr_con = 1
            max_con = 1
            for j in range(1, len(day_indices)):
                if day_indices[j] == day_indices[j-1] + 1:
                    curr_con += 1
                    max_con = max(max_con, curr_con)
                else:
                    curr_con = 1
            if max_con > consecutive_limit:
                continue
            possible_rooms = dfRoom[(dfRoom['type'] == 'theory') if room_type == 'theory' else dfRoom['type'].str.contains('lab')]
            room_found = False
            chosen_room = None
            for r_idx, r_row in possible_rooms.iterrows():
                r_id = r_row['room_id']
                if r_row['capacity'] < session['batch_size']:
                    continue
                if any(d == day and s == slot_time for d, s in room_schedule[r_id]):
                    continue
                chosen_room = r_id
                room_found = True
                break
            if not room_found:
                continue
            empty = True
            for div in session['divs']:
                key_tt = f"Year{session['year']}_Div{div}"
                if timetables[key_tt].at[day, slot_time] != "":
                    empty = False
                    break
            if not empty:
                continue
            possible_slots.append((day, slot_time, chosen_room))
    if not possible_slots:
        print(f"Cannot schedule session: {session}")
        continue
    def get_day_hours(d):
        return len([s for dd, s in faculty_schedule[faculty] if dd == d])
    possible_slots = sorted(possible_slots, key=lambda x: get_day_hours(x[0]))
    day, slot_time, room = possible_slots[0]
    faculty_schedule[faculty].append((day, slot_time))
    room_schedule[room].append((day, slot_time))
    faculty_hours[faculty] += 1
    entry = f"{session['name']} ({session['type']}) - {faculty} - {room}"
    for div in session['divs']:
        key_tt = f"Year{session['year']}_Div{div}"
        timetables[key_tt].at[day, slot_time] = entry

# --- Fill break slots ---
break_slots = [s for s in slots if s.startswith('12:15') or s.startswith('15:15')]
for key in timetables:
    for day in WORKING_DAYS:
        for bs in break_slots:
            if bs.startswith('12:15'):
                timetables[key].at[day, bs] = "Lunch Break"
            else:
                timetables[key].at[day, bs] = "Tea Break"

# --- Print and export timetables ---
with pd.ExcelWriter("College_Timetables99.xlsx") as writer:
    for key in timetables:
        print(f"\nTimetable for {key}:")
        print(timetables[key].to_string(justify="center"))
        sheet_name = key.replace(" ", "_")[:31]
        timetables[key].to_excel(writer, sheet_name=sheet_name)
print("\nAll timetables have been exported to 'College_Timetables99.xlsx'.")