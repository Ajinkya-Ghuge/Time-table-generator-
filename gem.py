import pandas as pd
from collections import defaultdict
import random

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
    # New dummy subjects to fill gaps
    {'course_id': '901', 'name': 'Elective A', 'program': 'B.Tech CSE', 'year': 2, 'divisions': 'A,B,C,D', 'credit': 1, 'hours_per_week': 1, 'theory_hours': 1, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'No', 'faculty_id': 'F999', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '902', 'name': 'Elective B', 'program': 'B.Tech CSE', 'year': 3, 'divisions': 'A,B,C', 'credit': 1, 'hours_per_week': 1, 'theory_hours': 1, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'No', 'faculty_id': 'F999', 'batch_size': 60, 'preferred_slot': None},
    {'course_id': '903', 'name': 'Elective C', 'program': 'B.Tech CSE', 'year': 4, 'divisions': 'A,B', 'credit': 1, 'hours_per_week': 1, 'theory_hours': 1, 'practical_hours': 0, 'lab_required': 'No', 'room_type': 'Classroom', 'central': 'No', 'faculty_id': 'F999', 'batch_size': 60, 'preferred_slot': None},
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
    # Dummy faculty for new elective courses with high capacity
    {'faculty_id': 'F999', 'name': 'Tutor', 'expertise': '[901, 902, 903] (Electives)', 'max_hours_per_week': 20, 'unavailable_slots': '[]', 'preferred_slots': '[]', 'consecutive_limit': 6},
]
dfFaculty = pd.DataFrame(faculty_data)
dfFaculty['unavailable_slots'] = dfFaculty['unavailable_slots'].apply(parse_list)
dfFaculty['preferred_slots'] = dfFaculty['preferred_slots'].apply(parse_list)

# --- Room Data ---
room_data = [
    {'room_id': 'R101', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R102', 'capacity': 60, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R103', 'capacity': 60, 'type': 'theory', 'available_slots': None}, # Added room
    {'room_id': 'R104', 'capacity': 60, 'type': 'theory', 'available_slots': None}, # Added room
    {'room_id': 'R201', 'capacity': 30, 'type': 'lab', 'available_slots': None},
    {'room_id': 'R202', 'capacity': 30, 'type': 'lab', 'available_slots': None},
    {'room_id': 'CL1', 'capacity': 40, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'CL2', 'capacity': 40, 'type': 'computer_lab', 'available_slots': None},
    {'room_id': 'CL3', 'capacity': 40, 'type': 'computer_lab', 'available_slots': None}, # Added room
    {'room_id': 'AUD1', 'capacity': 120, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R301', 'capacity': 25, 'type': 'theory', 'available_slots': None},
    {'room_id': 'R302', 'capacity': 25, 'type': 'lab', 'available_slots': None},
    {'room_id': 'R303', 'capacity': 25, 'type': 'lab', 'available_slots': None},
    {'room_id': 'CR1', 'capacity': 250, 'type': 'theory', 'available_slots': None},
    {'room_id': 'CR2', 'capacity': 200, 'type': 'theory', 'available_slots': None},
]
dfRoom = pd.DataFrame(room_data)

# --- Constants ---
COLLEGE_START = 10*60 + 15 # 10:15 AM
COLLEGE_END = 17*60 + 30 # 5:30 PM
THEORY_DURATION = 60 # 1 hour
WORKING_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
BREAKS = [
    {"name": "Lunch Break", "start": 12*60 + 15, "duration": 60}, # 12:15–13:15
    {"name": "Tea Break", "start": 15*60 + 15, "duration": 15}, # 3:15–3:30
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
    room_type = row['room_type']
    preferred_list = preferred.split(',') if preferred else []
    
    # Calculate a "difficulty" score for each session to prioritize scheduling
    score = 0
    if preferred:
        score += 10 # Prioritize sessions with preferred slots
    if central:
        score += 5 # Central courses are harder to place
    if room_type == 'Lab':
        score += 3 # Labs are a scarcer resource
    score += (theory_h + prac_h) # Longer sessions are harder to place
    
    if central:
        for i in range(theory_h):
            sess_pref = preferred_list[i] if i < len(preferred_list) else None
            sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': divs, 'central': True, 'room_type': 'theory', 'batch_size': batch_size, 'preferred': sess_pref, 'score': score})
        for i in range(prac_h):
            sess_pref = preferred_list[i + theory_h] if i + theory_h < len(preferred_list) else None
            sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': divs, 'central': True, 'room_type': 'lab', 'batch_size': batch_size, 'preferred': sess_pref, 'score': score})
    else:
        for div in divs:
            for i in range(theory_h):
                sessions.append({'type': 'theory', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'theory', 'batch_size': 60, 'preferred': None, 'score': score})
            for i in range(prac_h):
                sessions.append({'type': 'practical', 'course_id': row['course_id'], 'name': name, 'faculty': faculty, 'year': year, 'divs': [div], 'central': False, 'room_type': 'lab', 'batch_size': 60, 'preferred': None, 'score': score})

# --- Sort sessions by score (descending) ---
sessions.sort(key=lambda s: s['score'], reverse=True)

# --- Scheduling structures ---
faculty_schedule = defaultdict(list)
room_schedule = defaultdict(list)
faculty_hours = defaultdict(int)

# --- Schedule sessions ---
for session in sessions:
    possible_slots_with_scores = []
    faculty = session['faculty']
    max_hours = dfFaculty[dfFaculty['faculty_id'] == faculty]['max_hours_per_week'].item()
    consecutive_limit = dfFaculty[dfFaculty['faculty_id'] == faculty]['consecutive_limit'].item()
    unavailable = dfFaculty[dfFaculty['faculty_id'] == faculty]['unavailable_slots'].item()
    room_type = session['room_type']
    if session['name'] == 'MDM - Seminar':
        room_type = 'theory'

    for day in WORKING_DAYS:
        # Check faculty's daily hour limit
        day_hours = len([s for d, s in faculty_schedule[faculty] if d == day])
        if day_hours >= 6:
            continue
        
        for slot_num, slot_time in enumerate(lecture_slot_times, 1):
            key = f"{day}-{slot_num}"

            # Check preferred slot constraint
            if session['preferred'] and session['preferred'] != key:
                continue

            # Check faculty unavailability
            if key in unavailable:
                continue

            # Check faculty busy in this slot
            if any(d == day and s == slot_time for d, s in faculty_schedule[faculty]):
                continue

            # Check faculty weekly hour limit
            if faculty_hours[faculty] + 1 > max_hours:
                continue

            # Check consecutive hour constraint
            day_slots = sorted([lecture_slot_times.index(s) for d, s in faculty_schedule[faculty] if d == day])
            day_indices = day_slots + [lecture_slot_times.index(slot_time)]
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

            # Find a suitable and available room
            possible_rooms = dfRoom[dfRoom['type'] == room_type]
            if room_type == 'lab':
                possible_rooms = dfRoom[dfRoom['type'].str.contains('lab')]
            
            # Find an available room that meets capacity
            chosen_room = None
            for r_idx, r_row in possible_rooms.iterrows():
                r_id = r_row['room_id']
                if any(d == day and s == slot_time for d, s in room_schedule[r_id]):
                    continue
                if session['batch_size'] <= r_row['capacity']:
                    chosen_room = r_id
                    break
            
            if not chosen_room:
                continue

            # Check if all divisions for the session are free in this slot
            divisions_are_empty = True
            for div in session['divs']:
                key_tt = f"Year{session['year']}_Div{div}"
                if timetables[key_tt].at[day, slot_time] != "":
                    divisions_are_empty = False
                    break
            if not divisions_are_empty:
                continue

            # Check for same-day subject constraint
            # NOTE: This constraint is commented out to allow for a fuller timetable
            # same_day_subject = False
            # for div in session['divs']:
            #     key_tt = f"Year{session['year']}_Div{div}"
            #     day_schedule = timetables[key_tt].loc[day]
            #     if any(session['name'] in str(item) for item in day_schedule if item):
            #         same_day_subject = True
            #         break
            # if same_day_subject:
            #     continue
            
            # Score this potential slot
            slot_score = 0
            # Preferred slot gives a huge bonus
            if session['preferred'] == key:
                slot_score += 100
            # Prefer slots that have fewer available rooms
            available_rooms_count = sum(1 for r_idx, r_row in possible_rooms.iterrows() if not any(d == day and s == slot_time for d, s in room_schedule[r_row['room_id']]))
            slot_score += (len(possible_rooms) - available_rooms_count) * 2
            
            # Prefer slots that are earlier in the day
            slot_score += (len(lecture_slot_times) - slot_num) * 1

            possible_slots_with_scores.append((slot_score, day, slot_time, chosen_room))

    if not possible_slots_with_scores:
        print(f"Cannot schedule session: {session['name']} for {session['divs']} (no available slot)")
        continue

    # Select the slot with the highest score
    possible_slots_with_scores.sort(key=lambda x: x[0], reverse=True)
    day, slot_time, room = possible_slots_with_scores[0][1:]

    # Commit the scheduling
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
with pd.ExcelWriter("College_Timetables_v3.xlsx") as writer:
    for key in timetables:
        print(f"\nTimetable for {key}:")
        print(timetables[key].to_string(justify="center"))
        sheet_name = key.replace(" ", "_")[:31]
        timetables[key].to_excel(writer, sheet_name=sheet_name)
print("\nAll timetables have been exported to 'College_Timetables_v3.xlsx'.")