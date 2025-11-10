# College Timetable Scheduling System

This Python-based project generates fully populated, conflict-free timetables for B.Tech CSE students in Years 2, 3, and 4. It schedules classes for multiple divisions, ensuring all lecture slots are filled while respecting faculty, room, and student constraints. The output is exported to an Excel file (`College_Timetables.xlsx`) with a separate sheet for each division.

## Project Overview
- **Purpose**: Creates weekly timetables for college divisions, scheduling theory and practical classes without gaps.
- **Scope**: Covers Years 2 (Divisions A, B, C, D), 3 (A, B, C), and 4 (A, B), with courses like Data Structures, Operating Systems, and Machine Learning.
- **Output**: Excel file with timetables, each showing 6 days (Monday–Saturday), 6 lecture slots/day, and fixed breaks (Lunch and Tea).
- **Key Feature**: All slots are filled using a modified dataset with increased course hours (36 hours/year) and centralized scheduling.

## Setup Instructions
1. **Install Python**: Ensure Python 3.6+ is installed on your system.
2. **Install Dependencies**:
   ```bash
   pip install pandas
   ```
   - `pandas`: Used for data handling and Excel output.
3. **Prepare the Script**:
   - Save the provided Python script as `timetable_scheduler.py`.
   - Ensure it’s in a directory with write permissions to avoid errors when saving the Excel file.
4. **Run the Script**:
   ```bash
   python timetable_scheduler.py
   ```
   - The script generates `College_Timetables.xlsx` in the same directory.
   - Console output displays each timetable for review.

## Dataset Description
The system uses three datasets defined in the script:
- **Courses**:
  - Lists courses (e.g., Data Structures, Machine Learning) with details like year, divisions, hours (9 hours/course), and faculty.
  - All courses are centralized (shared across divisions) with batch sizes of 120–240.
- **Faculty**:
  - Specifies faculty details (e.g., ID, expertise, max hours of 40/week).
  - No unavailable slots; preferred slots (e.g., Wed-2 for OE - Environmental Studies) are included.
- **Rooms**:
  - Includes rooms with capacities (25–240) and types (theory, lab, computer_lab).
  - Added large rooms (e.g., AUD2, LAB_BIG) to support big batches.

## Scheduling Constraints
The system follows these rules to ensure a valid, fully populated timetable:
1. **Faculty Work Limits**: Each faculty can teach up to 40 hours per week and no more than 6 slots per day to prevent overworking.
2. **Consecutive Teaching Limit**: Faculty can teach up to 6 consecutive slots per day to maintain reasonable schedules and avoid exhaustion.
3. **No Unavailable Slots**: Faculty have no restricted time slots, allowing maximum flexibility for scheduling classes.
4. **Preferred Slots**: Courses with preferred time slots (e.g., Sat-1 for MDM - Seminar, Wed-2 for OE - Environmental Studies) are prioritized for those slots when possible.
5. **Room Exclusivity**: No two classes can be scheduled in the same room at the same time, ensuring rooms are booked only when free
6. **Room Capacity**: Rooms must have enough seats for the class size (e.g., 240-student classes require large rooms like AUD2 or LAB_BIG).
7. **Room Type Matching**: Theory classes are assigned to classrooms or auditoriums, while practical classes are assigned to labs, with no mismatches.
8. **No Student Conflicts**: Each division (student group) can have only one class per time slot, preventing overlaps in their schedules.
9. **Fixed Breaks**: Lunch Break (12:15–13:15) and Tea Break (15:15–15:30) are reserved daily, with no classes scheduled during these times.
10. **Centralized Courses**: Central courses (e.g., Data Structures for all Year 2 divisions) are scheduled once for multiple divisions in a single session, using large rooms to save time and resources.
11. **Balanced Distribution**: Classes are spread evenly across days, prioritizing slots on less busy days for faculty to balance their workload.


## How It Works
- **Input Processing**: Reads course, faculty, and room data from the script’s datasets.
- **Session Creation**: Generates sessions (theory/practical) based on course hours, grouping divisions for central courses.
- **Scheduling Logic**: Assigns sessions to slots, respecting constraints, using large rooms for big batches, and prioritizing preferred slots.
- **Output**: Fills all 36 lecture slots/week (6 days × 6 slots) with courses, adds breaks, and exports timetables to Excel.
- **Error Handling**: Logs unschedulable sessions (none in this case due to sufficient hours and rooms).

## Sample Output
- **File**: `College_Timetables.xlsx`
- **Sheets**: One per division (e.g., Year2_DivA, Year3_DivB, Year4_DivA).
- **Format**: Rows (days: Mon–Sat), columns (slots: 10:15–17:30), entries like "Data Structures (theory) - F001 - AUD2".
- **Example (Year 2, Monday)**: Data Structures (theory), Data Structures (practical), Lunch Break, Digital Logic Design (theory), Digital Logic Design (practical), Tea Break, Mathematics III (theory), Mathematics III (theory).
- **Result**: No empty slots; all divisions in a year share identical schedules due to central courses.

## Limitations and Future Improvements
- **Current Limitations**:
  - Fixed dataset requires manual updates for new courses or faculty.
  - Assumes all courses are centralized, which may not suit all colleges.
- **Suggested Enhancements**:
  - Add a user interface for inputting data.
  - Optimize for varied course distributions (e.g., fewer hours for some subjects).
  - Integrate with a database for dynamic updates.
  - Handle partial scheduling failures with fallback options.

## Troubleshooting
- **Permission Errors**: Ensure the script has write access to the output directory. Move it out of cloud folders (e.g., OneDrive) or run as administrator.
- **Missing Modules**: Verify `pandas` is installed (`pip show pandas`). Install `openpyxl` if Excel errors occur (`pip install openpyxl`).
- **Empty Slots**: If slots are empty (unlikely with this dataset), check course hours (must total 36/week/year) or reduce constraints.

## License
MIT License. See the [LICENSE](LICENSE) file for details.

## Contact
For issues or suggestions, please open an issue on the repository or contact the developer.
