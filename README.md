College Timetable Scheduling System

This project is a Python-based timetable scheduling system designed to create conflict-free, fully populated timetables for a college, specifically for B.Tech CSE students in Years 2, 3, and 4. It uses the provided dataset to schedule classes while adhering to constraints and exports the results to an Excel file (College_Timetables.xlsx). The system ensures all lecture slots are filled without gaps by using a modified dataset that balances course hours and resources.

Features





Generates timetables for multiple divisions (e.g., Year2_DivA, Year3_DivB, Year4_DivA) across a 6-day week (Monday to Saturday).



Schedules both theory and practical classes, respecting faculty, room, and student constraints.



Exports timetables to an Excel file with separate sheets for each division.



Handles central courses (shared across divisions) and individual division classes.



Ensures no vacant lecture slots by allocating exactly enough course hours to fill the schedule.

Prerequisites





Python 3.6+



Required Libraries:





pandas: For data handling and Excel output.



Install via: pip install pandas

How to Run





Clone the Repository:

git clone <repository_url>
cd <repository_directory>



Install Dependencies:

pip install pandas



Run the Script:





Ensure the Python script (timetable_scheduler.py) is in your working directory.



Execute: python timetable_scheduler.py



Output:





The script generates College_Timetables.xlsx with timetables for each division.



Console output shows a preview of each timetable.

Dataset

The system uses three main datasets defined in the script:





Course Data: Details about courses (e.g., course ID, name, year, divisions, hours, faculty, room type, batch size).





Modified to make all courses central, with each course set to 9 hours/week (total 36 hours/year) to fill all slots.



Faculty Data: Information about teachers (e.g., ID, expertise, max hours, unavailable/preferred slots, consecutive limits).





Adjusted to allow 40 max hours/week, no unavailable slots, and up to 6 consecutive hours.



Room Data: Room details (e.g., ID, capacity, type).





Added large-capacity rooms (e.g., AUD2, AUD3, LAB_BIG) to handle big batches (120-240 students).

Constraints Followed

The scheduling algorithm adheres to the following rules to ensure a valid timetable:





Faculty Work Limits: Each faculty has a maximum weekly teaching limit (40 hours) and a daily limit (6 slots).



No Back-to-Back Overload: Faculty can teach up to 6 consecutive slots per day to avoid exhaustion.



Faculty Unavailable Times: No classes are scheduled during a faculty’s unavailable slots (set to none in this version).



Preferred Times: Courses or faculty with preferred slots (e.g., Wed-2 for OE - Environmental Studies) are prioritized for those times.



Room Availability: No two classes can use the same room simultaneously.



Room Size Fit: Rooms must have enough capacity for the class’s batch size (e.g., 240-student classes use AUD2 or LAB_BIG).



Room Type Match: Theory classes use classrooms/auditoriums; practical classes use labs.



No Student Overlaps: A division cannot have multiple classes in the same time slot.



Fixed Breaks: Lunch (12:15-13:15) and tea break (15:15-15:30) are reserved daily.



Central Course Sharing: Central courses are scheduled once for multiple divisions in a large room.



Even Distribution: Classes are spread across days, prioritizing less busy days for faculty.

Output





Excel File: College_Timetables.xlsx contains a sheet for each division (e.g., Year2_DivA, Year3_DivB).



Timetable Structure: Each sheet covers Monday to Saturday, with 6 lecture slots/day plus breaks. Entries include course name, type (theory/practical), faculty ID, and room ID (e.g., "Data Structures (theory) - F001 - AUD2").



No Vacancies: All lecture slots are filled, with central courses ensuring efficient use of time and resources.

Example Output

For Year 2 (Divisions A, B, C, D):





Identical schedules due to central courses.



Example (Monday): Data Structures (theory), Data Structures (practical), Lunch Break, Digital Logic Design (theory), Digital Logic Design (practical), Tea Break, Mathematics III (theory), Mathematics III (theory).



Uses large rooms (AUD2, LAB_BIG) to accommodate 240 students.

Notes





The dataset was modified to ensure full slot coverage (36 hours/week per year, all central courses, more large rooms).



If a session cannot be scheduled (e.g., due to room or faculty conflicts), it is logged in the console but does not affect the final timetable due to sufficient hours.



For production use, consider adding:





Optimization for balanced subject distribution.



User interface for input/output.



Database integration for dynamic data.



Handling of edge cases (e.g., partial scheduling failures).
