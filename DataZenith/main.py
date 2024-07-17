import tkinter
from tkinter import ttk
import tkinter.messagebox
import tkinter.ttk
from tkinter import messagebox
import os
import openpyxl

def enter_data():

    accepted = accept_var.get()

    if accepted == "Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname:
            title = title_combox.get()
            age = age_spinbox.get()
            nationality = nationality_combox.get()

            # Course info
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()
            registration_status = reg_status_var.get() 


            print("First name: ", firstname , "\nLast name: ", lastname)
            print("Title: ", title, "\nAge: ", age, "\nNationality:", nationality)

            print("No. of Courses: ", numcourses, "\nNo. of Semesters: ", numsemesters)
            print("Registration Status: ", registration_status)

            print("--------------------------------------------")

            filepath = r"C:\Users\Sakthi\Desktop\py project\data.xlsx"



            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Firts Name", "Last Name", "Title","Age", "Nationality", "No. of Courses", "No. of Semesters", "Registration Status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality, numcourses, numsemesters, registration_status])
            workbook.save(filepath)

            clear_fields()

        else:
            tkinter.messagebox.showwarning(title="error", message="First name and Last name is required")

    else:
        tkinter.messagebox.showwarning(title= "Error", message= "You have not accepted the terms")


def clear_fields():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    title_combox.set('')
    age_spinbox.delete(0, tkinter.END)
    nationality_combox.set('')
    numcourses_spinbox.delete(0, tkinter.END)
    numsemesters_spinbox.delete(0, tkinter.END)
    reg_status_var.set("Not registered")
    accept_var.set("Not Accepted")



window = tkinter.Tk()
window.title("DataZenith")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row= 0, column= 0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row= 0, column= 0)

last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)

first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms,", "Mrs"])
title_label.grid(row=0, column=2)
title_combox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=1, to=100)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combox = ttk.Combobox(user_info_frame, values=["Africa", "Antartica", "Asia", "Europe", "North America", "Oceania", "South America"])
nationality_label.grid(row=2, column=1)
nationality_combox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
course_frame = tkinter.LabelFrame(frame)
course_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(course_frame, text="Registration Status")

reg_status_var = tkinter.StringVar(value="Not registered")
registered_check = tkinter.Checkbutton(course_frame, text="Currently Registered", 
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")


registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(course_frame, text="No. of Completed Courses")
numcourses_spinbox = tkinter.Spinbox(course_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(course_frame, text="No. of Semesters")
numsemesters_spinbox = tkinter.Spinbox(course_frame, from_=0, to='infinity')
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in course_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms and Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)





window.mainloop()