from ortools.sat.python import cp_model
import pandas as pd


DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]

SLOTS = [
    ("HM",    "09:00-10:15"),
    ("S1",    "10:15-11:15"),
    ("S2",    "11:15-12:15"),
    ("Lunch", "12:15-01:15"),
    ("S3",    "01:15-02:15"),
    ("S4",    "02:15-03:15"),
    ("Tea",   "03:15-03:30"),
    ("MDM",   "03:30-05:30"),
]

DIVISIONS = ["A", "B", "C" ]
SUBJECTS = ["FOIL", "DAA", "DBMS", "IDS", "SUB5"]

# Fixed slots 
FIXED = {
    "HM": "HM",
    "Lunch": "Break",
    "Tea": "Break",
    "MDM": "MDM",
}

# Faculty weekly hours. 
MAX_HOURS = 14


model = cp_model.CpModel()

x = {}
for d in DIVISIONS:
    for day in DAYS:
        for slot_key, _ in SLOTS:
            if slot_key in FIXED:
                continue
            for subj in SUBJECTS:
                x[(d, day, slot_key, subj)] = model.NewBoolVar(f"x_{d}_{day}_{slot_key}_{subj}")


# Constraints

# 1) At most one subject per non-fixed slot per division (allows empty -> Free)
for d in DIVISIONS:
    for day in DAYS:
        for slot_key, _ in SLOTS:
            if slot_key in FIXED:
                continue
            model.Add(sum(x[(d, day, slot_key, subj)] for subj in SUBJECTS) <= 1)

# 2) No subject repeat in same day for a division 
for d in DIVISIONS:
    for day in DAYS:
        for subj in SUBJECTS:
            model.Add(
                sum(
                    x[(d, day, slot_key, subj)]
                    for slot_key, _ in SLOTS
                    if slot_key not in FIXED
                ) <= 1
            )

# 3) No subject overlap across divisions at same slot 
for day in DAYS:
    for slot_key, _ in SLOTS:
        if slot_key in FIXED:
            continue
        for subj in SUBJECTS:
            model.Add(
                sum(x[(d, day, slot_key, subj)] for d in DIVISIONS) <= 1
            )

# 4) Faculty weekly hours 
for subj in SUBJECTS:
    model.Add(
        sum(
            x[(d, day, slot_key, subj)]
            for d in DIVISIONS
            for day in DAYS
            for slot_key, _ in SLOTS
            if slot_key not in FIXED
        ) <= MAX_HOURS
    )

#5)  No continuous teaching for faculty n
#skip jar empty asal
slot_order = [slot_key for slot_key, _ in SLOTS]
for subj in SUBJECTS:
    for day in DAYS:
        for i in range(len(slot_order) - 1):
            s1 = slot_order[i]
            s2 = slot_order[i + 1]
            if s1 in FIXED or s2 in FIXED:
                continue
           
            model.Add(
                sum(x[(d, day, s1, subj)] for d in DIVISIONS) +
                sum(x[(d, day, s2, subj)] for d in DIVISIONS) <= 1
            )


# Objective ki maximize number of assigned lectures asave

assign_vars = [v for v in x.values()]
model.Maximize(sum(assign_vars))

# Solve

solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 30

solver.parameters.num_search_workers = 8

status = solver.Solve(model)

# Build timetable output (always produces something)

def build_timetable_from_solution(solver):
    timetable = {}
    for d in DIVISIONS:
        rows = []
        for slot_key, timing in SLOTS:
            row = {"Slot": slot_key, "Timing": timing}
            for day in DAYS:
                
                if slot_key in FIXED:
                    row[day] = FIXED[slot_key]
                    continue
                
                assigned_subj = None
                for subj in SUBJECTS:
                    var = x[(d, day, slot_key, subj)]
                    if solver.Value(var) == 1:
                        assigned_subj = subj
                        break
                row[day] = assigned_subj if assigned_subj is not None else "Free"
            rows.append(row)
        timetable[d] = pd.DataFrame(rows)[["Slot", "Timing"] + DAYS]
    return timetable

if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    timetable = build_timetable_from_solution(solver)
else:
    # safety net 
    timetable = {}
    for d in DIVISIONS:
        rows = []
        for slot_key, timing in SLOTS:
            row = {"Slot": slot_key, "Timing": timing}
            for day in DAYS:
                row[day] = FIXED[slot_key] if slot_key in FIXED else "Free"
            rows.append(row)
        timetable[d] = pd.DataFrame(rows)[["Slot", "Timing"] + DAYS]


# Save to Excel

out_path = "timetable_auto_manage.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    for d in DIVISIONS:
        timetable[d].to_excel(writer, sheet_name=f"Div_{d}", index=False)

print("Timetable exported to:", out_path)
if status == cp_model.OPTIMAL:
    print("Solver status: OPTIMAL (best possible under constraints).")
elif status == cp_model.FEASIBLE:
    print("Solver status: FEASIBLE (found a valid assignment, not proven optimal).")
else:
    print("Solver status: no feasible/optimal solution; output contains fixed slots and 'Free' placeholders.")
