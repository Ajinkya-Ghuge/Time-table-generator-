@app.route('/download_sample_csv')
def download_sample_csv():
    """Download perfect sample CSV files"""
    
    # Courses CSV
    courses_csv = """course_id,name,program,year,divisions,credit,hours_per_week,theory_hours,practical_hours,lab_required,room_type,central,faculty_id,batch_size,preferred_slot
CSE201,Data Structures,B.Tech CSE,2,"A,B",4,6,4,2,Yes,Lab,No,F101,60,"Mon-1"
CSE202,Digital Circuits,B.Tech CSE,2,"A,B",4,6,4,2,Yes,Lab,No,F102,60,"Tue-1"
MAT203,Mathematics III,B.Tech CSE,2,"A,B",4,6,6,0,No,Classroom,No,F103,60,"Wed-1"
PHY204,Physics Lab,B.Tech CSE,2,"A,B",3,4,0,4,Yes,Lab,No,F104,30,"Thu-1, Thu-3"
ENG205,Communication Skills,B.Tech CSE,2,"A,B",3,4,4,0,No,Classroom,Yes,F105,120,"Fri-1"
PRO206,Programming Basics,B.Tech CSE,2,"A,B",4,6,4,2,Yes,Lab,No,F106,60,"Sat-1"
CSE207,Computer Organization,B.Tech CSE,2,"A,B",4,6,4,2,Yes,Lab,No,F107,60,"Mon-3"
WEB208,Web Development,B.Tech CSE,2,"A,B",4,6,4,2,Yes,Lab,No,F108,60,"Tue-3""""
    
    # Faculty CSV
    faculty_csv = """faculty_id,name,expertise,max_hours_per_week,unavailable_slots,preferred_slots,consecutive_limit
F101,Dr. Anil Sharma,"[CSE201] (Data Structures)",24,"[]","[Mon-1, Mon-2, Mon-3, Mon-4, Mon-5, Mon-6]",6
F102,Prof. Priya Patel,"[CSE202] (Digital Circuits)",24,"[]","[Tue-1, Tue-2, Tue-3, Tue-4, Tue-5, Tue-6]",6
F103,Dr. Raj Kumar,"[MAT203] (Mathematics III)",24,"[]","[Wed-1, Wed-2, Wed-3, Wed-4, Wed-5, Wed-6]",6
F104,Dr. Sunita Singh,"[PHY204] (Physics Lab)",24,"[]","[Thu-1, Thu-2, Thu-3, Thu-4, Thu-5, Thu-6]",6
F105,Prof. Michael Brown,"[ENG205] (Communication Skills)",24,"[]","[Fri-1, Fri-2, Fri-3, Fri-4, Fri-5, Fri-6]",6
F106,Dr. Sarah Chen,"[PRO206] (Programming Basics)",24,"[]","[Sat-1, Sat-2, Sat-3, Sat-4, Sat-5, Sat-6]",6
F107,Prof. David Wang,"[CSE207] (Computer Organization)",24,"[]","[Mon-1, Mon-2, Mon-3, Mon-4, Mon-5, Mon-6]",6
F108,Dr. Emily Garcia,"[WEB208] (Web Development)",24,"[]","[Tue-1, Tue-2, Tue-3, Tue-4, Tue-5, Tue-6]",6"""
    
    # Rooms CSV
    rooms_csv = """room_id,capacity,type,available_slots
LT-101,60,theory,
LT-102,60,theory,
LT-103,60,theory,
LT-104,60,theory,
LAB-A,30,lab,
LAB-B,30,lab,
LAB-C,30,lab,
LAB-D,30,lab,
AUD-1,120,theory,
LAB-COMP1,40,computer_lab,
LAB-COMP2,40,computer_lab,"""
    
    file_type = request.args.get('file', 'courses')
    
    if file_type == 'courses':
        return send_file(
            io.BytesIO(courses_csv.encode('utf-8')),
            as_attachment=True,
            download_name='perfect_courses.csv',
            mimetype='text/csv'
        )
    elif file_type == 'faculty':
        return send_file(
            io.BytesIO(faculty_csv.encode('utf-8')),
            as_attachment=True,
            download_name='perfect_faculty.csv',
            mimetype='text/csv'
        )
    elif file_type == 'rooms':
        return send_file(
            io.BytesIO(rooms_csv.encode('utf-8')),
            as_attachment=True,
            download_name='perfect_rooms.csv',
            mimetype='text/csv'
        )
    else:
        # Download all three as ZIP
        import zipfile
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            zip_file.writestr('perfect_courses.csv', courses_csv)
            zip_file.writestr('perfect_faculty.csv', faculty_csv)
            zip_file.writestr('perfect_rooms.csv', rooms_csv)
        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name='perfect_timetable_data.zip',
            mimetype='application/zip'
        )