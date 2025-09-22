import os
import re
import pandas as pd
from collections import defaultdict
from ortools.sat.python import cp_model

# ---------------- Config ----------------
WORKING_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
SLOTS = [
    ("S1", "10:15-11:15"),
    ("S2", "11:15-12:15"),
    ("Lunch", "12:15-1:15"),
    ("S3", "1:15-2:15"),
    ("S4", "2:15-3:15"),
    ("Tea", "3:15-3:30"),
    ("S5", "3:30-4:30"),
    ("S6", "4:30-5:30"),
]
NON_FIXED_SLOTS = [s for s, _ in SLOTS if s not in ["Lunch", "Tea"]]

COURSE_PATH = "Course(Subjects).xlsx"
ROOM_PATH = "Room.xlsx"
FAC_PATH = "Faculty.xlsx"

# ---------------- Load data ----------------
courses_df = pd.read_excel(COURSE_PATH)
rooms_df = pd.read_excel(ROOM_PATH)
fac_df = pd.read_excel(FAC_PATH)

# Normalize column names
courses_df.columns = courses_df.columns.str.lower()
rooms_df.columns = rooms_df.columns.str.lower()
fac_df.columns = fac_df.columns.str.lower()

# ---------------- Process courses ----------------
subjects = courses_df["course_id"].astype(str).tolist()
subject_info = {}
DIVISIONS = []

for _, r in courses_df.iterrows():
    subj_id = str(r["course_id"])
    year = int(r["year"])

    # Handle divisions column: might be "4" or "A,B,C,D"
    div_val = str(r["divisions"]).strip()
    if div_val.isdigit():
        divs = [f"D{d}" for d in range(1, int(div_val) + 1)]
    else:
        divs = [d.strip() for d in div_val.split(",") if d.strip()]

    # Save subject info
    subject_info[subj_id] = {
        "year": year,
        "divisions": divs,
        "theory_hours": int(r.get("theory_hours", 0)),
        "lab_required": str(r.get("lab_required", "")).lower()
        in ("1", "true", "yes"),
    }

    # Build divisions list like Y2_A, Y2_B
    for d in divs:
        DIVISIONS.append(f"Y{year}_{d}")

DIVISIONS = sorted(set(DIVISIONS))

# ---------------- Process rooms ----------------
rooms = rooms_df["room_id"].astype(str).tolist()
room_info = {
    str(r["room_id"]): {
        "is_lab": str(r.get("is_lab", "")).lower() in ("1", "true", "yes", "lab")
        or "lab" in str(r["room_id"]).lower()
    }
    for _, r in rooms_df.iterrows()
}

# ---------------- Process faculty ----------------
faculty_list = fac_df["faculty"].astype(str).tolist()
faculty_info = {}
subject_to_faculty = defaultdict(list)

for _, r in fac_df.iterrows():
    f = str(r["faculty"])
    maxh = int(r.get("max_hours", 15))
    subs = []
    if "course_id" in r and pd.notna(r["course_id"]):
        subs = re.split(r"[;,]", str(r["course_id"]))
        subs = [s.strip() for s in subs if s.strip() != ""]
    faculty_info[f] = {"max_hours": maxh, "subjects": subs}
    for s in subs:
        subject_to_faculty[s].append(f)

# ---------------- Model ----------------
model = cp_model.CpModel()
teach = {}
room_for = {}
fac_for = {}

for d in DIVISIONS:
    year = int(d.split("_")[0][1:])
    div_label = d.split("_")[1]
    for day in WORKING_DAYS:
        for slot in NON_FIXED_SLOTS:
            for subj in subjects:
                info = subject_info[subj]
                if info["year"] != year:
                    continue
                if div_label not in info["divisions"]:
                    continue
                teach[(d, day, slot, subj)] = model.NewBoolVar(
                    f"teach_{d}_{day}_{slot}_{subj}"
                )
                for r in rooms:
                    if info["lab_required"] and not room_info[r]["is_lab"]:
                        continue
                    room_for[(d, day, slot, subj, r)] = model.NewBoolVar(
                        f"room_{d}_{day}_{slot}_{subj}_{r}"
                    )
                for f in subject_to_faculty.get(subj, []):
                    fac_for[(d, day, slot, subj, f)] = model.NewBoolVar(
                        f"fac_{d}_{day}_{slot}_{subj}_{f}"
                    )

# ---------------- Constraints ----------------
# one subject per div/day/slot
for d in DIVISIONS:
    for day in WORKING_DAYS:
        for slot in NON_FIXED_SLOTS:
            model.Add(
                sum(teach.get((d, day, slot, s), 0) for s in subjects) <= 1
            )

# no repeat same subject per day/div
for d in DIVISIONS:
    for day in WORKING_DAYS:
        for s in subjects:
            model.Add(
                sum(
                    teach.get((d, day, slot, s), 0) for slot in NON_FIXED_SLOTS
                )
                <= 1
            )

# subject not in two divisions same slot
for day in WORKING_DAYS:
    for slot in NON_FIXED_SLOTS:
        for s in subjects:
            model.Add(
                sum(teach.get((d, day, slot, s), 0) for d in DIVISIONS) <= 1
            )

# faculty link + no double-booking
for (d, day, slot, s), tvar in teach.items():
    facs = [
        fac_for[(d, day, slot, s, f)]
        for f in subject_to_faculty.get(s, [])
        if (d, day, slot, s, f) in fac_for
    ]
    if facs:
        model.Add(sum(facs) == tvar)
for f in faculty_list:
    for day in WORKING_DAYS:
        for slot in NON_FIXED_SLOTS:
            model.Add(
                sum(
                    fac_for.get((d, day, slot, s, f), 0)
                    for d in DIVISIONS
                    for s in subjects
                )
                <= 1
            )

# room link + no double-booking
for (d, day, slot, s), tvar in teach.items():
    rvars = [
        room_for[(d, day, slot, s, r)]
        for r in rooms
        if (d, day, slot, s, r) in room_for
    ]
    if rvars:
        model.Add(sum(rvars) == tvar)
for r in rooms:
    for day in WORKING_DAYS:
        for slot in NON_FIXED_SLOTS:
            model.Add(
                sum(
                    room_for.get((d, day, slot, s, r), 0)
                    for d in DIVISIONS
                    for s in subjects
                )
                <= 1
            )

# required hours exactly
for s, info in subject_info.items():
    req = info["theory_hours"]
    if req <= 0:
        continue
    for d in DIVISIONS:
        year = int(d.split("_")[0][1:])
        div_label = d.split("_")[1]
        if year != info["year"]:
            continue
        if div_label not in info["divisions"]:
            continue
        model.Add(
            sum(
                teach.get((d, day, slot, s), 0)
                for day in WORKING_DAYS
                for slot in NON_FIXED_SLOTS
            )
            == req
        )

# ---------------- Solve ----------------
solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 60
status = solver.Solve(model)
print("Status:", solver.StatusName(status))

# ---------------- Output ----------------
flat = []
for d in DIVISIONS:
    for slot, timing in SLOTS:
        for day in WORKING_DAYS:
            if slot in ["Lunch", "Tea"]:
                flat.append([d, day, slot, timing, slot, None, None])
                continue
            sub = None
            fac = None
            room = None
            for s in subjects:
                if (
                    teach.get((d, day, slot, s))
                    and solver.Value(teach[(d, day, slot, s)]) == 1
                ):
                    sub = s
                    for f in subject_to_faculty.get(s, []):
                        if fac_for.get((d, day, slot, s, f)) and solver.Value(
                            fac_for[(d, day, slot, s, f)]
                        ) == 1:
                            fac = f
                    for r in rooms:
                        if room_for.get((d, day, slot, s, r)) and solver.Value(
                            room_for[(d, day, slot, s, r)]
                        ) == 1:
                            room = r
                    break
            flat.append([d, day, slot, timing, sub or "Free", fac, room])

df = pd.DataFrame(
    flat, columns=["Division", "Day", "Slot", "Timing", "Subject", "Faculty", "Room"]
)
df.to_csv("debug_assignments.csv", index=False)

with pd.ExcelWriter("timetable_auto_manage.xlsx") as writer:
    for d, grp in df.groupby("Division"):
        pivot = (
            grp.pivot(index=["Slot", "Timing"], columns="Day", values="Subject")
            .reset_index()
        )
        pivot.to_excel(writer, sheet_name=d[:31], index=False)

print("Exported timetable_auto_manage.xlsx + debug_assignments.csv")
