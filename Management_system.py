import mysql.connector
import pandas as pd

cursor=None
try:
    conn = mysql.connector.connect(host='localhost', user='root', password='waqasbruh404')
    if conn.is_connected():
        print("Connection Established.....")
    # Create a cursor to interact with the database
    cursor = conn.cursor()

    # Create the database if it doesn't exist
    cursor.execute("CREATE DATABASE IF NOT EXISTS StudentsDB")
    cursor.execute("USE StudentsDB")

    table_creating_query="""
CREATE TABLE IF NOT EXISTS Student_Info (
    Rollnumber INT PRIMARY KEY,
    Name VARCHAR(20),
    Course VARCHAR(20),
    Marks INT
)
"""
    cursor.execute(table_creating_query)

    conn.commit()

    select_query="""SELECT * FROM Student_Info"""
    cursor.execute(select_query)
    rows=cursor.fetchall()

    column_headers = [desc[0] for desc in cursor.description]


    def adding_to_excel(rows):
        if rows:
            df = pd.DataFrame(rows, columns=column_headers)
            # Save the DataFrame to an Excel file
            Python_For_Data = "student_data.xlsx"
            df.to_excel(Python_For_Data, index=False)
            print(f"Data saved to {Python_For_Data}")
        else:
            print("No records found in the table.")

    def add_student():
        rollnumber = int(input("Enter Rollnumber: "))
        name = input("Enter Name: ")
        course = input("Enter Course: ")
        marks = int(input("Enter Marks: "))

        insert_query = "INSERT INTO Student_Info (Rollnumber, Name, Course, Marks) VALUES (%s, %s, %s, %s)"
        data=(rollnumber,name,course,marks)
        cursor.execute(insert_query,data)
        select_query="""SELECT * FROM Student_Info"""
        cursor.execute(select_query)
        rows=cursor.fetchall()
        adding_to_excel(rows)
    
    def show_student_list():
        print(pd.read_excel("student_data.xlsx"))

    def edit_student_info():
        getting_rollnumber=int(input("Enter Rollnumber of Student you want to edit: "))
        getting_rollnumber_from_excel=pd.read_excel("student_data.xlsx")
        if(getting_rollnumber in getting_rollnumber_from_excel["Rollnumber"]):
            getting_rollnumber_from_excel = getting_rollnumber_from_excel.loc[getting_rollnumber_from_excel['Rollnumber'] != getting_rollnumber]
            new_name = input("Enter Name: ")
            new_course = input("Enter Course: ")
            new_marks = int(input("Enter Marks: "))

            insert_query = "INSERT INTO Student_Info (Rollnumber,Name, Course, Marks) VALUES (%s, %s, %s, %s)"
            data=(getting_rollnumber,new_name,new_course,new_marks)
            cursor.execute(insert_query,data)
            select_query="""SELECT * FROM Student_Info"""
            cursor.execute(select_query)
            rows=cursor.fetchall()
            adding_to_excel(rows)
        else:
            print(f"Rollnumber: {getting_rollnumber} not found!")
    
    def delete_student():
        getting_rollnumber=int(input("Enter Rollnumber of Student you want to Delete: "))
        getting_rollnumber_from_excel=pd.read_excel("student_data.xlsx")
        if(getting_rollnumber in getting_rollnumber_from_excel["Rollnumber"]):
            getting_rollnumber_from_excel = getting_rollnumber_from_excel.loc[getting_rollnumber_from_excel['Rollnumber'] != getting_rollnumber]
            getting_rollnumber_from_excel.to_excel("student_data.xlsx", index=False)

    def register_student_in_course():
        add_student()

    def print_grades():
        def Calculating_Grades(marks):
            if(marks>90):
                return "A+"
            elif(marks>80):
                return "A"
            elif(marks>70):
                return "B"
            elif(marks>60):
                return "C"
            elif(marks>50):
                return "D"
            else:
                return "E"
        getting_rollnumber_from_excel=pd.read_excel("student_data.xlsx")
        new_column=[Calculating_Grades(getting_rollnumber_from_excel['marks'])] * len(getting_rollnumber_from_excel)
        getting_rollnumber_from_excel['Grades']=new_column
        # Display specific columns
        selected_columns = ['Rollnumber', 'Marks', 'Grades']
        df_selected = getting_rollnumber_from_excel[selected_columns]
        # print("DataFrame with Selected Columns:")
        print(df_selected)


    while(True):
        print("\n===== University Admission Office Management System =====")
        print("1. Add Student")
        print("2. Show List of Students")
        print("3. Edit Student Information")
        print("4. Delete Student")
        print("5. Register Student in a Course")
        print("6. Enter Marks of a Student or Complete Class")
        print("7. Print Grades of Students")
        print("8. Exit\n")
        choice=int(input("Enter Choice(letter): "))
        if (choice == 1):
            add_student()
        elif (choice == 2):
            show_student_list()
        elif (choice == 3):
            edit_student_info()
        elif (choice == 4):
            delete_student()
        elif (choice == 5):
            register_student_in_course()
        elif (choice == 7):
            print_grades()
        elif (choice == 8):
            break
        else:
            print("Invalid choice. Please enter a number between 1 and 8: ")
            choice=int(input("Enter Choice(letter): "))
        
except mysql.connector.Error as err:
    print(f"Error: {err}")

finally:
    # Close the cursor and connection in a finally block
    if 'cursor' in locals() and cursor:
        cursor.close()
    if 'conn' in locals() and conn.is_connected():
        conn.close()
        print("Connection Closed.....")
