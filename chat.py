import pandas as pd
from collections import defaultdict

def parse_list(s):
    if not s or s == '[]':
        return []
    s = s.strip('[]')
    return [item.strip() for item in s.split(',')]

# --- Course Data ---
course_data = [
    {'course_id': '101', 'name': 'Data Structures', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F001', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '102', 'name': 'Digital Logic Design', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F002', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '103', 'name': 'Mathematics III', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 3, 'hours_per_week': 3, 'theory_hours': 3, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'No', 'faculty_id': 'F003', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '104', 'name': 'OE - Environmental Studies', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 2, 'hours_per_week': 2, 'theory_hours': 2, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'Yes', 'faculty_id': 'F010', 'batch_size': 240, 'preferred_slot': 'Wed-2'},
    {'course_id': '201', 'name': 'Operating Systems', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F004', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '202', 'name': 'Database Systems', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F005', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '203', 'name': 'Computer Networks', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F006', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '204', 'name': 'OE - Constitution of India', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 2, 'hours_per_week': 2, 'theory_hours': 2, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'Yes', 'faculty_id': 'F011', 'batch_size': 180, 'preferred_slot': 'Thu-3'},
    {'course_id': '301', 'name': 'Machine Learning', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F007', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '302', 'name': 'Cloud Computing', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 4, 'hours_per_week': 4, 'theory_hours': 3, 'practical_hours': 1, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F008', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '303', 'name': 'Major Project', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 8, 'hours_per_week': 8, 'theory_hours': 0, 'practical_hours': 8, 'lab_required': 'Yes', 'room_type': 'Lab', 'central': 'No', 'faculty_id': 'F009', 'batch_size': 60, 'preferred_slot': 'Fri-2,3'},
    {'course_id': '304', 'name': 'MDM - Seminar', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 2, 'hours_per_week': 2, 'theory_hours': 2, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Seminar', 'central': 'Yes', 'faculty_id': 'F012', 'batch_size': 120, 'preferred_slot': 'Sat-1'},
]
dfCourse = pd.DataFrame(course_data)
dfCourse['divisions_list'] = dfCourse['divisions'].apply(lambda x: x.split(',') if pd.notna(x) else [])
dfCourse['divisions_num'] = dfCourse['divisions_list'].apply(len)

# --- Faculty Data ---
faculty_data = [
    {'faculty_id': 'F001', 'name': 'Dr. A. Sharma', 'expertise': '[101] (Data Structures)', 'max_hours_per_week': 14, 'unavailable_slots': '[Tue-1, Thu-4]', 'preferred_slots': '[Mon-2]', 'consecutive_limit': 2},
    {'faculty_id': 'F002', 'name': 'Prof. B. Patil', 'expertise': '[102] (Digital Logic)', 'max_hours_per_week': 12, 'unavailable_slots': '[Wed-3]', 'preferred_slots': '[]', 'consecutive_limit': 3},
    {'faculty_id': 'F003', 'name': 'Dr. C. Mehta', 'expertise': '[103] (Mathematics III)', 'max_hours_per_week': 10, 'unavailable_slots': '[Fri-1]', 'preferred_slots': '[Tue-2]', 'consecutive_limit': 2},
    {'faculty_id': 'F004', 'name': 'Prof. D. Kulkarni', 'expertise': '[201] (Operating Sys.)', 'max_hours_per_week': 14, 'unavailable_slots': '[Mon-4, Thu-2]', 'preferred_slots': '[]', 'consecutive_limit': 2},
    {'faculty_id': 'F005', 'name': 'Dr. E. Iyer', 'expertise': '[202] (Database Sys.)', 'max_hours_per_week': 14, 'unavailable_slots': '[Wed-2]', 'preferred_slots': '[Fri-3]', 'consecutive_limit': 2},
    {'faculty_id': 'F006', 'name': 'Dr. F. Singh', 'expertise': '[203] (Comp. Networks)', 'max_hours_per_week': 14, 'unavailable_slots': '[Thu-1]', 'preferred_slots': '[]', 'consecutive_limit': 3},
    {'faculty_id': 'F007', 'name': 'Prof. G. Deshmukh', 'expertise': '[301] (Machine Learning)', 'max_hours_per_week': 12, 'unavailable_slots': '[Mon-3, Fri-4]', 'preferred_slots': '[Wed-1]', 'consecutive_limit': 2},
    {'faculty_id': 'F008', 'name': 'Dr. H. Verma', 'expertise': '[302] (Cloud Computing)', 'max_hours_per_week': 12, 'unavailable_slots': '[Tue-4]', 'preferred_slots': '[]', 'consecutive_limit': 2},
    {'faculty_id': 'F009', 'name': 'Prof. I. Rao', 'expertise': '[303] (Major Project)', 'max_hours_per_week': 16, 'unavailable_slots': '[]', 'preferred_slots': '[Fri-2, Fri-3]', 'consecutive_limit': 4},
    {'faculty_id': 'F010', 'name': 'Dr. J. Joshi', 'expertise': '[104] (OE - EVS)', 'max_hours_per_week': 8, 'unavailable_slots': '[Mon-1]', 'preferred_slots': '[Wed-2]', 'consecutive_limit': 2},
    {'faculty_id': 'F011', 'name': 'Prof. K. Banerjee', 'expertise': '[204] (OE - Constitution)', 'max_hours_per_week': 8, 'unavailable_slots': '[Thu-2]', 'preferred_slots': '[Thu-3]', 'consecutive_limit': 2},
    {'faculty_id': 'F012', 'name': 'Dr. L. Reddy', 'expertise': '[304] (MDM - Seminar)', 'max_hours_per_week': 6, 'unavailable_slots': '[]', 'preferred_slots': '[Sat-1]', 'consecutive_limit': 2},
]
dfFaculty = pd.DataFrame(faculty_data)
dfFaculty['unavailable_slots'] = dfFaculty['unavailable_slots'].apply(parse_list)
dfFaculty['preferred_slots'] = dfFaculty['preferred_slots'].apply(parse_list)

# --- Room Data ---
room_data = [
    {'room_id': 'R101', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R102', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R201', 'capacity': 30, 'type': 'lab', 'available_slots': None},
    {'room_id': 'R202', 'capacity': 30, 'type': 'lab', 'available_slots': None},
    {'room_id': 'CL1', 'capacity': 40, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'CL2', 'capacity': 40, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'AUD1', 'capacity': 120, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R301', 'capacity': 25, 'type': 'theory', 'available_slots': None},
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
    preferred_list = preferred.split(',') if preferred else []
    if central:
        for i in range(theory_h):
            sess_pref = preferred_list[i] if i < len(preferred_list) else None
            sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': divs, 'central': True, 'room_type': 'theory', 'batch_size': batch_size, 'preferred': sess_pref})
        for i in range(prac_h):
            sess_pref = preferred_list[i + theory_h] if i + theory_h < len(preferred_list) else None
            sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': divs, 'central': True, 'room_type': 'lab', 'batch_size': batch_size, 'preferred': sess_pref})
    else:
        for div in divs:
            for i in range(theory_h):
                sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'theory', 'batch_size': 60, 'preferred': None})
            for i in range(prac_h):
                sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'lab', 'batch_size': 60, 'preferred': None})

# --- Sort sessions ---
sessions = sorted(sessions, key=lambda s: s['preferred'] is None)

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
            # consecutive check relaxed
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
            if max_con > consecutive_limit + 1:
                continue
            # back-to-back check
            back_to_back = False
            for div in session['divs']:
                key_tt = f"Year{session['year']}_Div{div}"
                day_schedule = timetables[key_tt].loc[day]
                slot_index = lecture_slot_times.index(slot_time)
                if slot_index > 0 and session['name'] in str(day_schedule[lecture_slot_times[slot_index - 1]]):
                    back_to_back = True
                    break
                if slot_index < len(lecture_slot_times) - 1 and session['name'] in str(day_schedule[lecture_slot_times[slot_index + 1]]):
                    back_to_back = True
                    break
            if back_to_back:
                continue
            # room availability
            room_id = None
            for r_idx, r_row in dfRoom[dfRoom['type'] == room_type].iterrows():
                r_id = r_row['room_id']
                if all(timetables[f"Year{session['year']}_Div{div}"].at[day, slot_time] == "" for div in session['divs']):
                    room_id = r_id
                    break
            if room_id:
                possible_slots.append((day, slot_time, room_id))
    # fallback
    if not possible_slots:
        for day in WORKING_DAYS:
            for slot_time in lecture_slot_times:
                for r_idx, r_row in dfRoom[dfRoom['type'] == room_type].iterrows():
                    r_id = r_row['room_id']
                    if all(timetables[f"Year{session['year']}_Div{div}"].at[day, slot_time] == "" for div in session['divs']):
                        possible_slots = [(day, slot_time, r_id)]
                        break
                if possible_slots:
                    break
            if possible_slots:
                break
    if not possible_slots:
        print(f"⚠ Could not place session: {session}")
        continue
    chosen = possible_slots[0]
    day, slot_time, room_id = chosen
    for div in session['divs']:
        key_tt = f"Year{session['year']}_Div{div}"
        timetables[key_tt].at[day, slot_time] = f"{session['name']} ({room_id})"
    faculty_schedule[faculty].append((day, slot_time))
    faculty_hours[faculty] += 1
    room_schedule[room_id].append((day, slot_time))

# --- Display sample timetable ---
for key, df in timetables.items():
    print(f"\nTimetable for {key}:\n")
    print(df)

# --- Export all timetables to Excel ---
with pd.ExcelWriter("University_Timetables.xlsx") as writer:
    for key, df in timetables.items():
        df.to_excel(writer, sheet_name=key)
print(" Timetables exported to University_Timetables.xlsx")
