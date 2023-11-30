import csv
import tkinter
from datetime import datetime
from tkinter import *
from tkinter import messagebox, ttk

import pymysql
import xlsxwriter

window = tkinter.Tk()
window.title("SmartScore")
window.geometry("820x640")
window.configure(bg="#000000")
my_tree=ttk.Treeview(window,show='headings',height=20)
style = ttk.Style()

placeholderArray=['','','','','','','','']
numeric='1234567890'
alpha='ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def connection():
    conn=pymysql.connect(
        host='localhost',
        user='root',
        password='',
        db='sms'
    )
    return conn

conn=connection()
cursor=conn.cursor()

for i in range(0,8):
    placeholderArray[i]=tkinter.StringVar()
    
def read():
    cursor.connection.ping()
    sql=f"SELECT `first_name`, `middle_name`, `surname`, `year_level`, `program`, `student_id`, `midterm_grade`, `finals_grade`, `GWA` FROM student_data ORDER BY `id` DESC"
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.commit()
    conn.close()
    return results

def refreshTable():
    for data in my_tree.get_children():
        my_tree.delete(data)
    for array in read():
        my_tree.insert(parent='',index='end',iid=array,text="",values=(array),tag="orow")
    my_tree.tag_configure('orow',background="#EEEEEE")
    my_tree.pack()
    
def setph(word,num):
    for ph in range(0,8):
        if ph == num:
            placeholderArray[ph].set(word)
            
def save_button_click():
    save()

def update_button_click():
    update()

def delete_button_click():
    delete()

def select_button_click():
    select()

def find_button_click():
    find()

def clear_button_click():
    clear()

def export_button_click():
    exportExcel()

def save():
    FirstName = str(FirstNameEntry.get())
    MiddleName = str(MiddleNameEntry.get())
    Surname = str(SurnameEntry.get())
    YearLevel = str(YearCombo.get())
    Program = str(ProgramCombo.get())
    StudentID = str(StudentIDEntry.get())
    MidtermGrade = float(MidtermGradeEntry.get())
    FinalsGrade = float(FinalsGradeEntry.get())

    GWA = (MidtermGrade + FinalsGrade) / 2

    valid = True

    if not (FirstName and FirstName.strip()) or not (MiddleName and MiddleName.strip()) or not (
            Surname and Surname.strip()) or not (YearLevel and YearLevel.strip()) or not (
            Program and Program.strip()) or not (StudentID and StudentID.strip()) or not (
            MidtermGrade and str(MidtermGrade).strip()) or not (
            FinalsGrade and str(FinalsGrade).strip()):
        messagebox.showwarning("", "Please fill up all entries")
        return

    cursor.connection.ping()
    sql = f"SELECT * FROM student_data WHERE `student_id` = '{StudentID}' "
    cursor.execute(sql)
    checkItemNo = cursor.fetchall()

    if len(checkItemNo) > 0:
        messagebox.showwarning("", "Student ID already used")
    else:
        try:
            cursor.connection.ping()
            sql = f"INSERT INTO student_data (`first_name`, `middle_name`, `surname`, `year_level`, `program`, `student_id`, `midterm_grade`, `finals_grade`, `GWA`) VALUES ('{FirstName}','{MiddleName}','{Surname}','{YearLevel}','{Program}','{StudentID}','{MidtermGrade}','{FinalsGrade}','{GWA}')"
            cursor.execute(sql)
            conn.commit()
        except Exception as e:
            print(e)
            messagebox.showwarning("", "Error while saving ref: " + str(e))
        finally:
            conn.close()

        refreshTable()

    

def update():
    selectedItemId = ''
    try:
        selectedItem = my_tree.selection()[0]
        selectedItemId = str(my_tree.item(selectedItem)['values'][5])
    except:
        messagebox.showwarning("", "Please select a data row")
    print(selectedItemId)
    FirstName = str(FirstNameEntry.get())
    MiddleName = str(MiddleNameEntry.get())
    Surname = str(SurnameEntry.get())
    YearLevel = str(YearCombo.get())
    Program = str(ProgramCombo.get())
    StudentID = str(StudentIDEntry.get())
    MidtermGrade = float(MidtermGradeEntry.get())
    FinalsGrade = float(FinalsGradeEntry.get())
    
    if not(FirstName and FirstName.strip()) or not(MiddleName and MiddleName.strip()) or not(Surname and Surname.strip()) or not(YearLevel and YearLevel.strip()) or not(Program and Program.strip()) or not(StudentID and StudentID.strip()) or not(MidtermGrade and str(MidtermGrade).strip()) or not(FinalsGrade and str(FinalsGrade).strip()):
        messagebox.showwarning("","Please fill up all entries")
        return
    if(selectedItemId!=StudentID):
        messagebox.showwarning("","You can't change Student ID")
        return
    try:
        cursor.connection.ping()
        sql_update = f"UPDATE student_data SET `first_name` = '{FirstName}', `middle_name` = '{MiddleName}', `surname` = '{Surname}', `year_level` = '{YearLevel}', `program` = '{Program}', `midterm_grade` = '{MidtermGrade}', `finals_grade` = '{FinalsGrade}' WHERE `student_id` = '{StudentID}' "
        cursor.execute(sql_update)
        conn.commit()

        cursor.connection.ping()
        sql_gwa = f"SELECT (`midterm_grade` + `finals_grade`) / 2 AS `GWA` FROM student_data WHERE `student_id` = '{StudentID}'"
        cursor.execute(sql_gwa)
        existing_gwa = cursor.fetchone()

        if existing_gwa:
            new_gwa = (MidtermGrade + FinalsGrade) / 2
            cursor.connection.ping()
            sql_update_gwa = f"UPDATE student_data SET `GWA` = '{new_gwa}' WHERE `student_id` = '{StudentID}'"
            cursor.execute(sql_update_gwa)
            conn.commit()
        else:
            new_gwa = (MidtermGrade + FinalsGrade) / 2
            cursor.connection.ping()
            sql_insert_gwa = f"INSERT INTO student_data (`GWA`) VALUES ('{new_gwa}') WHERE `student_id` = '{StudentID}'"
            cursor.execute(sql_insert_gwa)
            conn.commit()
            
        updated_values = (FirstName, MiddleName, Surname, YearLevel, Program, StudentID, MidtermGrade, FinalsGrade, new_gwa)
        my_tree.item(selectedItem, values=updated_values)

    except Exception as err:
        messagebox.showwarning("", "Error occurred ref: " + str(err))
        return
    
def delete():
    try:
        if(my_tree.selection()[0]):
            decision = messagebox.askquestion("", "Delete the selected data?")
            if(decision != 'yes'):
                return
            else:
                selectedItem = my_tree.selection()[0]
                StudentID = str(my_tree.item(selectedItem)['values'][5])
                try:
                    cursor.connection.ping()
                    sql=f"DELETE FROM student_data WHERE `student_id` = '{StudentID}' "
                    cursor.execute(sql)
                    conn.commit()
                    conn.close()
                    messagebox.showinfo("","Data has been successfully deleted")
                except:
                    messagebox.showinfo("","Sorry, an error occured")
                refreshTable()
    except:
        messagebox.showwarning("", "Please select a data row")

def select():
    try:
        selectedItem = my_tree.selection()[0]
        FirstName = str(my_tree.item(selectedItem)['values'][0])
        MiddleName = str(my_tree.item(selectedItem)['values'][1])
        Surname = str(my_tree.item(selectedItem)['values'][2])
        YearLevel = str(my_tree.item(selectedItem)['values'][3])
        Program = str(my_tree.item(selectedItem)['values'][4])
        StudentID = str(my_tree.item(selectedItem)['values'][5])
        MidtermGrade = str(my_tree.item(selectedItem)['values'][6])
        FinalsGrade = str(my_tree.item(selectedItem)['values'][7])
        setph(FirstName,0)
        setph(MiddleName,1)
        setph(Surname,2)
        setph(YearLevel,3)
        setph(Program,4)
        setph(StudentID,5)
        setph(MidtermGrade,6)
        setph(FinalsGrade,7)
    except:
        messagebox.showwarning("", "Please select a data row")
        
def find():
    FirstName = str(FirstNameEntry.get())
    MiddleName = str(MiddleNameEntry.get())
    Surname = str(SurnameEntry.get())
    YearLevel = str(YearCombo.get())
    Program = str(ProgramCombo.get())
    StudentID = str(StudentIDEntry.get())
    MidtermGrade_str = str(MidtermGradeEntry.get())
    FinalsGrade_str = str(FinalsGradeEntry.get())
    MidtermGrade = float(MidtermGrade_str) if MidtermGrade_str.strip() else 0.0
    FinalsGrade = float(FinalsGrade_str) if FinalsGrade_str.strip() else 0.0
    cursor.connection.ping()


    if FirstName and FirstName.strip():
        sql = f"SELECT * FROM student_data WHERE `first_name` = '{FirstName}'"
    elif MiddleName and MiddleName.strip():
        sql = f"SELECT * FROM student_data WHERE `middle_name` = '{MiddleName}'"
    elif Surname and Surname.strip():
        sql = f"SELECT * FROM student_data WHERE `surname` = '{Surname}'"
    elif YearLevel and YearLevel.strip():
        sql = f"SELECT * FROM student_data WHERE `year_level` = '{YearLevel}'"
    elif Program and Program.strip():
        sql = f"SELECT * FROM student_data WHERE `program` = '{Program}'"
    elif StudentID and StudentID.strip():
        sql = f"SELECT * FROM student_data WHERE `student_id` = '{StudentID}'"
    elif MidtermGrade and str(MidtermGrade).strip():
        sql = f"SELECT * FROM student_data WHERE `midterm_grade` = {MidtermGrade}"
    elif FinalsGrade and str(FinalsGrade).strip():
        sql = f"SELECT * FROM student_data WHERE `finals_grade` = {FinalsGrade}"
    else:
        messagebox.showwarning("", "Please fill up one of the entries")
        return

    cursor.execute(sql)
    try:
        result = cursor.fetchall()
        if result:
            setph(result[0][1], 0)
            setph(result[0][2], 1)
            setph(result[0][3], 2)
            setph(result[0][4], 3)
            setph(result[0][5], 4)
            setph(result[0][6], 5)
            setph(result[0][7], 6)
            setph(result[0][8], 7)
            conn.commit()
            conn.close()
        else:
            messagebox.showwarning("", "No data found")

    except Exception as e:
        messagebox.showwarning("", "Error occurred: " + str(e))
        
        
def clear():
    for num in range(0,8):
        setph('',(num))
        
def exportExcel():
    try:
        cursor.connection.ping()
        sql = "SELECT `first_name`, `middle_name`, `surname`, `year_level`, `program`, `student_id`, `midterm_grade`, `finals_grade`, `GWA` FROM student_data ORDER BY `id` DESC"
        cursor.execute(sql)
        dataraw = cursor.fetchall()

        if not dataraw:
            messagebox.showinfo("", "No data to export.")
            return

        date = str(datetime.now())
        date = date.replace(' ', '_')
        date = date.replace(':', '-')
        dateFinal = date[0:16]

        workbook = xlsxwriter.Workbook(f"student_data_{dateFinal}.xlsx")
        worksheet = workbook.add_worksheet()

        # Write header with formatting
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        header = ["First Name", "Middle Name", "Surname", "Year Level", "Program", "Student ID", "Midterm Grade", "Finals Grade", "GWA"]
        for col_num, value in enumerate(header):
            worksheet.write(0, col_num, value, header_format)

        # Write data rows with formatting
        data_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
        for row_num, record in enumerate(dataraw, start=1):
            for col_num, value in enumerate(record):
                worksheet.write(row_num, col_num, value, data_format)

        workbook.close()

        messagebox.showinfo("", f"Excel file saved as student_data_{dateFinal}.xlsx")
    except Exception as e:
        print(e)
        messagebox.showwarning("", f"Error exporting to Excel: {str(e)}")
    finally:
        conn.commit()
        conn.close()
    

btnColor = "#630000"
mainFrame = tkinter.Frame(window)
mainFrame.pack(side="top", anchor="nw")


manageFrame = tkinter.LabelFrame(mainFrame, text="Manage", bg="#EEEBDD")
manageFrame.grid(row=0, column=1)

saveBtn = Button(manageFrame, text="SAVE", width=14, borderwidth=4, bg=btnColor, fg='white', command=save_button_click)
updateBtn = Button(manageFrame, text="UPDATE", width=14, borderwidth=4, bg=btnColor, fg='white', command=update_button_click)
deleteBtn = Button(manageFrame, text="DELETE", width=14, borderwidth=4, bg=btnColor, fg='white', command=delete_button_click)
selectBtn = Button(manageFrame, text="SELECT", width=14, borderwidth=4, bg=btnColor, fg='white', command=select_button_click)
findBtn = Button(manageFrame, text="FIND", width=14, borderwidth=4, bg=btnColor, fg='white', command=find_button_click)
clearBtn = Button(manageFrame, text="CLEAR", width=14, borderwidth=4, bg=btnColor, fg='white', command=clear_button_click)
exportBtn = Button(manageFrame, text="EXPORT EXCEL", width=14, borderwidth=4, bg=btnColor, fg='white', command=export_button_click)

saveBtn.grid(row=0, column=0, padx=40, pady=5)
updateBtn.grid(row=1, column=0, padx=10, pady=5)
deleteBtn.grid(row=2, column=0, padx=10, pady=5)
selectBtn.grid(row=3, column=0, padx=10, pady=5)
findBtn.grid(row=4, column=0, padx=10, pady=5)
clearBtn.grid(row=5, column=0, padx=10, pady=5)
exportBtn.grid(row=6, column=0, padx=10, pady=5)


entriesFrame = tkinter.LabelFrame(mainFrame, text="Form", bg="#EEEBDD")
entriesFrame.grid(row=0, column=0)

FirstNameLabel = Label(entriesFrame, text="FIRST NAME", bg="#EEEBDD", anchor="e", width=30)
MiddleNameLabel = Label(entriesFrame, text="MIDDLE NAME", bg="#EEEBDD", anchor="e", width=30)
SurnameLabel = Label(entriesFrame, text="SURNAME", bg="#EEEBDD", anchor="e", width=30)
YearLabel = Label(entriesFrame, text="YEAR LEVEL", bg="#EEEBDD", anchor="e", width=30)
ProgramLabel = Label(entriesFrame, text="PROGRAM", bg="#EEEBDD", anchor="e", width=30)
StudentIDLabel = Label(entriesFrame, text="STUDENT ID", bg="#EEEBDD", anchor="e", width=30)
MidtermGradeLabel = Label(entriesFrame, text="MIDTERM GRADE", bg="#EEEBDD", anchor="e", width=30)
FinalsGradeLabel = Label(entriesFrame, text="FINALS GRADE", bg="#EEEBDD", anchor="e", width=30)

FirstNameLabel.grid(row=0, column=0, padx=15, pady=7)
MiddleNameLabel.grid(row=1, column=0, padx=10, pady=7)
SurnameLabel.grid(row=2, column=0, padx=10, pady=7)
YearLabel.grid(row=3, column=0, padx=10, pady=7)
ProgramLabel.grid(row=4, column=0, padx=10, pady=7)
StudentIDLabel.grid(row=5, column=0, padx=10, pady=7)
MidtermGradeLabel.grid(row=6, column=0, padx=10, pady=7)
FinalsGradeLabel.grid(row=7, column=0, padx=10, pady=7)

yearArray=['First Year','Second Year','Third Year','Fourth Year',]
programArray=['Computer Science','Information Technology']

FirstNameEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[0])
MiddleNameEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[1])
SurnameEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[2])
YearCombo=ttk.Combobox(entriesFrame,width=56,textvariable=placeholderArray[3],values=yearArray)
ProgramCombo=ttk.Combobox(entriesFrame,width=56,textvariable=placeholderArray[4],values=programArray)
StudentIDEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[5])
MidtermGradeEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[6])
FinalsGradeEntry=Entry(entriesFrame,width=59,textvariable=placeholderArray[7])

FirstNameEntry.grid(row=0, column=2, padx=8, pady=7)
MiddleNameEntry.grid(row=1, column=2, padx=8, pady=7)
SurnameEntry.grid(row=2, column=2, padx=8, pady=7)
YearCombo.grid(row=3, column=2, padx=8, pady=7)
ProgramCombo.grid(row=4, column=2, padx=8, pady=7)
StudentIDEntry.grid(row=5, column=2, padx=8, pady=7)
MidtermGradeEntry.grid(row=6, column=2, padx=8, pady=7)
FinalsGradeEntry.grid(row=7, column=2, padx=8, pady=7)

style.configure(window)
my_tree['columns']=("First Name","Middle Name","Surname","Year Level","Program","Student ID","Midterm Grade","Finals Grade","GWA")
my_tree.column("#0",width=0,stretch=NO)
my_tree.column("First Name",anchor=W,width=90)
my_tree.column("Middle Name",anchor=W,width=100)
my_tree.column("Surname",anchor=W,width=90)
my_tree.column("Year Level",anchor=W,width=85)
my_tree.column("Program",anchor=W,width=140)
my_tree.column("Student ID",anchor=W,width=75)
my_tree.column("Midterm Grade",anchor=W,width=100)
my_tree.column("Finals Grade",anchor=W,width=90)
my_tree.column("GWA",anchor=W,width=50)

my_tree.heading("First Name",text="First Name",anchor=W)
my_tree.heading("Middle Name",text="Middle Name",anchor=W)
my_tree.heading("Surname",text="Surname",anchor=W)
my_tree.heading("Year Level",text="Year Level",anchor=W)
my_tree.heading("Program",text="Program",anchor=W)
my_tree.heading("Student ID",text="Student ID",anchor=W)
my_tree.heading("Midterm Grade",text="Midterm Grade",anchor=W)
my_tree.heading("Finals Grade",text="Finals Grade",anchor=W)
my_tree.heading("GWA",text="GWA",anchor=W)
my_tree.tag_configure('orow', background="#FFCCCC")
my_tree.pack()

refreshTable()

window.resizable(False, False)
window.mainloop()
