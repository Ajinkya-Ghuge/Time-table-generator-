# Time-table-generator-

The timetable scheduling system follows these main constraints (rules) to make sure everything runs smoothly without conflicts. I'll explain them simply, like rules for organizing a school day:

Faculty Work Limits: Each teacher has a max number of hours they can teach per week (e.g., 40 hours max now, to avoid overworking them). They also can't teach more than 6 classes in one day.
No Back-to-Back Overload for Teachers: Teachers have a limit on how many classes they can do in a row without a break (e.g., up to 6 now, but it prevents too many consecutive slots to avoid tiredness).
Teacher Unavailable Times: Teachers have specific times they're not free (e.g., certain days or slots marked as unavailable), so no classes are scheduled then.
Preferred Times for Classes: Some classes or teachers have favorite time slots (e.g., a seminar on Saturday morning), and the system tries to use those first if possible.
Room Availability: No two classes can use the same room at the same time. Rooms are booked only when free.
Room Size Fit: The room must be big enough for the number of students (e.g., a class of 240 needs a large auditorium like AUD2, not a small lab).
Room Type Match: Theory classes go in classrooms or auditoriums; practical (lab) classes go in labs. No mixing them up.
No Overlaps for Students/Divisions: A student group (division) can't have two classes at the exact same time. Their timetable must have only one thing per slot.
Fixed Breaks: Lunch (12:15-13:15) and tea break (15:15-15:30) are always there every day, no classes during those.
Central Course Sharing: Some classes are "central" (shared by multiple divisions), so one session covers everyone in a big group, saving time and rooms.
Even Distribution: The system tries to spread classes across days evenly, starting with less busy days for teachers.
