from ortools.sat.python import cp_model
import pandas as pd

# Load data
dfCourse = pd.read_excel("Course(Subjects).xlsx")
dfFaculty = pd.read_excel("Faculty.xlsx")
dfRoom = pd.read_excel("Room.xlsx")

# Clean data
dfCourse = dfCourse.apply(lambda col: col.map(lambda x: x.lower() if isinstance(x, str) else x))
dfFaculty = dfFaculty.apply(lambda col: col.map(lambda x: x.lower() if isinstance(x, str) else x))
dfRoom = dfRoom.apply(lambda col: col.map(lambda x: x.lower() if isinstance(x, str) else x))

# Parse divisions per year
divisions_per_year = {}
for y in dfCourse['year'].unique():
    div_str = dfCourse[dfCourse['year'] == y]['divisions'].iloc[0]
    divisions_per_year[y] = div_str.split(',')

# Courses per year
courses_per_year = {}
for y in dfCourse['year'].unique():
    courses_per_year[y] = dfCourse[dfCourse['year'] == y]['course_id'].tolist()

# Faculty for course
faculty_for_course = dfCourse.set_index('course_id')['faculty_id'].to_dict()

# Subject name for course
subject_name_for_course = dfCourse.set_index('course_id')['name'].to_dict()

# Max hours for faculty
max_hours_for_faculty = dfFaculty.set_index('faculty_id')['max_hours_per_week'].to_dict()

# Unavailable slots for faculty
unavailable_for_faculty = {}
for row in dfFaculty.itertuples():
    if isinstance(row.unavailable_slots, str) and row.unavailable_slots.strip('[]'):
        u = row.unavailable_slots[1:-1].split(',')
        unavailable_for_faculty[row.faculty_id] = [s.strip() for s in u if s.strip()]
    else:
        unavailable_for_faculty[row.faculty_id] = []

# Consecutive limit for faculty
consecutive_for_faculty = dfFaculty.set_index('faculty_id')['consecutive_limit'].to_dict()

# Hours per week for course
hours_for_course = dfCourse.set_index('course_id')['hours_per_week'].to_dict()

# Batch size for course
batch_size_for_course = dfCourse.set_index('course_id')['batch_size'].to_dict()

# Room type for course
room_type_for_course = dfCourse.set_index('course_id')['room_type'].to_dict()

# Lab required for course
lab_required_for_course = dfCourse.set_index('course_id')['lab_required'].to_dict()

# Central course flag
central_for_course = dfCourse.set_index('course_id')['central'].to_dict()

# Preferred slot for course
preferred_slot_for_course = dfCourse.set_index('course_id')['preferred_slot'].to_dict()

# Room data
room_capacity = dfRoom.set_index('room_id')['capacity'].to_dict()
room_type = dfRoom.set_index('room_id')['type'].to_dict()

# Assume division size
DIVISION_SIZE = 60

# Define days and slots
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
SLOTS = [
    ("S1", "10:15-11:15"),
    ("S2", "11:15-12:15"),
    ("Lunch", "12:15-13:15"),
    ("S3", "13:15-14:15"),
    ("S4", "14:15-15:15"),
    ("Tea", "15:15-15:30"),
    ("S5", "15:30-16:30"),
    ("S6", "16:30-17:30"),
]
FIXED = {
    "Lunch": "Lunch",
    "Tea": "Tea",
}
LECTURE_SLOTS = [s for s, _ in SLOTS if s not in FIXED]

# Slot and day maps for parsing unavailable and preferred slots
slot_map = {'1': 'S1', '2': 'S2', '3': 'S3', '4': 'S4', '5': 'S5', '6': 'S6'}
day_map = {'mon': 'Mon', 'tue': 'Tue', 'wed': 'Wed', 'thu': 'Thu', 'fri': 'Fri', 'sat': 'Sat'}

# Generate timetables year by year
out_path = "timetable_auto_manage.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    for y in dfCourse['year'].unique():
        divs = divisions_per_year[y]
        subjects = courses_per_year[y]
        rooms = dfRoom['room_id'].tolist()

        model = cp_model.CpModel()

        # Variables: x[d, day, slot, subj, room] = 1 if subject is assigned to division, day, slot, room
        x = {}
        for d in divs:
            for day in DAYS:
                for slot_key, _ in SLOTS:
                    if slot_key in FIXED:
                        continue
                    for subj in subjects:
                        for r in rooms:
                            x[(d, day, slot_key, subj, r)] = model.NewBoolVar(f"x_{d}_{day}_{slot_key}_{subj}_{r}")

        # Constraint 1: At most one subject and room per slot per division
        for d in divs:
            for day in DAYS:
                for slot_key in LECTURE_SLOTS:
                    model.Add(sum(x[(d, day, slot_key, subj, r)] for subj in subjects for r in rooms) <= 1)

        # Constraint 2: No subject repeat in same day for a division
        for d in divs:
            for day in DAYS:
                for subj in subjects:
                    model.Add(
                        sum(
                            x[(d, day, slot_key, subj, r)]
                            for slot_key in LECTURE_SLOTS
                            for r in rooms
                        ) <= 1
                    )

        # Constraint 3: Faculty teaching overlap (can teach only one group at a time)
        for day in DAYS:
            for slot_key in LECTURE_SLOTS:
                for subj in subjects:
                    model.Add(
                        sum(x[(d, day, slot_key, subj, r)] for d in divs for r in rooms) <= 1
                    )

        # Constraint 4: Batch size limit (max divisions per slot per subject)
        for day in DAYS:
            for slot_key in LECTURE_SLOTS:
                for subj in subjects:
                    batch = batch_size_for_course[subj]
                    max_div = batch // DIVISION_SIZE
                    model.Add(
                        sum(x[(d, day, slot_key, subj, r)] for d in divs for r in rooms) <= max_div
                    )

        # Constraint 5: Faculty weekly hours <= max
        for subj in subjects:
            f = faculty_for_course[subj]
            max_h = max_hours_for_faculty[f]
            model.Add(
                sum(
                    x[(d, day, slot_key, subj, r)]
                    for d in divs
                    for day in DAYS
                    for slot_key in LECTURE_SLOTS
                    for r in rooms
                ) <= max_h
            )

        # Constraint 6: No continuous teaching beyond limit for faculty
        for subj in subjects:
            f = faculty_for_course[subj]
            con_limit = consecutive_for_faculty[f]
            for day in DAYS:
                for i in range(len(LECTURE_SLOTS) - con_limit):
                    s_seq = LECTURE_SLOTS[i : i + con_limit + 1]
                    model.Add(
                        sum(x[(d, day, s, subj, r)] for d in divs for s in s_seq for r in rooms) <= con_limit
                    )

        # Constraint 7: Required hours per division per subject (== for non-central, adjusted for central)
        for subj in subjects:
            h = hours_for_course[subj]
            if central_for_course[subj] == 'yes':
                # Central courses: all divisions in same slot
                for day in DAYS:
                    for slot_key in LECTURE_SLOTS:
                        for r in rooms:
                            # If assigned to one division, assign to all relevant divisions
                            for d1 in divs:
                                model.Add(
                                    sum(x[(d, day, slot_key, subj, r)] for d in divs) == len(divs) * x[(d1, day, slot_key, subj, r)]
                                )
                # Total hours for central course (shared across divisions)
                model.Add(
                    sum(
                        x[(divs[0], day, slot_key, subj, r)]
                        for day in DAYS
                        for slot_key in LECTURE_SLOTS
                        for r in rooms
                    ) == h
                )
            else:
                # Non-central courses: hours per division
                for d in divs:
                    model.Add(
                        sum(
                            x[(d, day, slot_key, subj, r)]
                            for day in DAYS
                            for slot_key in LECTURE_SLOTS
                            for r in rooms
                        ) == h
                    )

        # Constraint 8: Unavailable slots for faculty
        for subj in subjects:
            f = faculty_for_course[subj]
            for u in unavailable_for_faculty[f]:
                if '-' in u:
                    try:
                        day_str, slot_str = u.split('-')
                        day = day_map.get(day_str.strip(), None)
                        slot_key = slot_map.get(slot_str.strip(), None)
                        if day and slot_key:
                            model.Add(sum(x[(d, day, slot_key, subj, r)] for d in divs for r in rooms) == 0)
                    except ValueError:
                        print(f"Warning: Invalid unavailable slot format '{u}' for faculty {f}. Skipping.")
                else:
                    print(f"Warning: Invalid unavailable slot format '{u}' for faculty {f}. Skipping.")

        # Constraint 9: Room compatibility (room type and lab requirement)
        for subj in subjects:
            req_room_type = room_type_for_course[subj]
            is_lab = lab_required_for_course[subj] == 'yes'
            for r in rooms:
                r_type = room_type[r]
                # Allow only compatible room types
                if req_room_type == 'classroom' and r_type != 'theory':
                    for d in divs:
                        for day in DAYS:
                            for slot_key in LECTURE_SLOTS:
                                model.Add(x[(d, day, slot_key, subj, r)] == 0)
                elif req_room_type == 'lab' and r_type != 'lab' and r_type != 'computer_lab':
                    for d in divs:
                        for day in DAYS:
                            for slot_key in LECTURE_SLOTS:
                                model.Add(x[(d, day, slot_key, subj, r)] == 0)
                elif req_room_type == 'seminar' and r_type != 'theory':
                    for d in divs:
                        for day in DAYS:
                            for slot_key in LECTURE_SLOTS:
                                model.Add(x[(d, day, slot_key, subj, r)] == 0)
                # Lab-required courses must be in lab or computer_lab
                if is_lab and r_type not in ['lab', 'computer_lab']:
                    for d in divs:
                        for day in DAYS:
                            for slot_key in LECTURE_SLOTS:
                                model.Add(x[(d, day, slot_key, subj, r)] == 0)

        # Constraint 10: Room capacity
        for day in DAYS:
            for slot_key in LECTURE_SLOTS:
                for r in rooms:
                    model.Add(
                        sum(
                            x[(d, day, slot_key, subj, r)] * batch_size_for_course[subj]
                            for d in divs
                            for subj in subjects
                        ) <= room_capacity[r]
                    )

        # Constraint 11: No room double-booking
        for day in DAYS:
            for slot_key in LECTURE_SLOTS:
                for r in rooms:
                    model.Add(
                        sum(x[(d, day, slot_key, subj, r)] for d in divs for subj in subjects) <= 1
                    )

        # Constraint 12: Preferred slots
        for subj in subjects:
            pref = preferred_slot_for_course[subj]
            if isinstance(pref, str) and '-' in pref:
                try:
                    day_str, slot_str = pref.split('-')
                    pref_day = day_map.get(day_str.strip(), None)
                    pref_slot = slot_map.get(slot_str.strip(), None)
                    if pref_day and pref_slot:
                        # Encourage scheduling in preferred slot
                        model.Add(
                            sum(
                                x[(d, pref_day, pref_slot, subj, r)]
                                for d in divs
                                for r in rooms
                            ) >= sum(
                                x[(d, day, slot_key, subj, r)]
                                for d in divs
                                for day in DAYS
                                for slot_key in LECTURE_SLOTS
                                for r in rooms
                                if day != pref_day or slot_key != pref_slot
                            )
                        )
                except ValueError:
                    print(f"Warning: Invalid preferred slot format '{pref}' for subject {subj}. Skipping.")

        # Objective: Maximize assigned lectures
        assign_vars = [v for v in x.values()]
        model.Maximize(sum(assign_vars))

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 60  # Increased for better solutions
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)

        # Build timetable from solution
        timetable = {}
        for d in divs:
            rows = []
            for slot_key, timing in SLOTS:
                row = {"Slot": slot_key, "Timing": timing}
                for day in DAYS:
                    if slot_key in FIXED:
                        row[day] = FIXED[slot_key]
                        continue
                    assigned_subj = None
                    assigned_room = None
                    for subj in subjects:
                        for r in rooms:
                            if (d, day, slot_key, subj, r) in x and solver.Value(x[(d, day, slot_key, subj, r)]) == 1:
                                assigned_subj = subj
                                assigned_room = r
                                break
                        if assigned_subj:
                            break
                    row[day] = f"{subject_name_for_course[assigned_subj]} ({assigned_room})" if assigned_subj else "Free"
                rows.append(row)
            timetable[d] = pd.DataFrame(rows)[["Slot", "Timing"] + DAYS]

        # Save to Excel
        for d in divs:
            timetable[d].to_excel(writer, sheet_name=f"Year{y}_Div{d}", index=False)

print("Timetable exported to:", out_path)
if status == cp_model.OPTIMAL:
    print("Solver status: OPTIMAL (best possible under constraints).")
elif status == cp_model.FEASIBLE:
    print("Solver status: FEASIBLE (valid assignment, not proven optimal).")
else:
    print("Solver status: no feasible/optimal solution; output contains fixed slots and 'Free' placeholders.")