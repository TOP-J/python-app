import tkinter as tk
from tkinter import scrolledtext
import subprocess
from tkinter import filedialog
from tableframe import TableFrameWidget
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# Register the Times New Roman font
pdfmetrics.registerFont(TTFont("Times New Roman", "times.ttf"))


from pymongo import MongoClient
from pymongo import errors
from pymongo.errors import ServerSelectionTimeoutError

try:
    client = MongoClient("mongodb://localhost:27017/")
    db = client["School"]
    students = db["Students"] #students collection
    courses = db["courses"] #courses collection
    scale_docs = db["grade_scale"] #grade_scale collection
    #retrieve default student info from db
    defaultStudent =  students.find({"course title":"SWE01A22"})
except ServerSelectionTimeoutError as e:
    print("connection failed")
    tk.Text("Unable to communicate with database servers. Please turn on MongoDB servers and try restarting application.",foreground="red",padx=20,pady=20)


#list of table widgets
grades_Widgets = []
GP_widgets = []
HM_widgets = []
CP_widgets = []
CV_widgets = []
CA_widgets = []
Exam_widgets = []
count = 0
for widget in grades_Widgets:
    if widget.cget("text") == "F":
        count = count + 1



def savefile():
    file_path = filedialog.asksaveasfilename(defaultextension=".txt",filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        with open(file_path, "w") as file:
            file.write("This is the content to be saved.")


def createtable_row(title,cv,code):
    """function creates table row"""

    table_row_Frame = tk.Frame(table_frame,bg="white") 
    coursetitle = tk.Label(table_row_Frame,text=title,font=("Times New Roman",14),bg="#6DB9EF",foreground="white")
    cvl = tk.Label(table_row_Frame,text=cv,font=("Times New Roman",14),bg="#6DB9EF",foreground="white",width=3)
    Cp = tk.Entry(table_row_Frame,font=("Times New Roman",14),bg="white",foreground="black",width=10)
    Hw = tk.Entry(table_row_Frame,font=("Times New Roman",14),bg="white",foreground="black",width=10)
    grade = tk.Label(table_row_Frame,text="0",font=("Times New Roman",14),bg="#6DB9EF",foreground="white",width=3)
    CA_mark = tk.Entry(table_row_Frame,font=("Times New Roman",14),bg="white",foreground="black",width=10)
    Exam_mark = tk.Entry(table_row_Frame,font=("Times New Roman",14),bg="white",foreground="black",width=10)
    Gp = tk.Label(table_row_Frame,text="0",font=("Times New Roman",14),bg="#6DB9EF",foreground="white",width=4)
 




    #displaying widgets using the Grid layout manager
    table_row_Frame.grid(sticky="nsew")
    table_row_Frame.columnconfigure(0,weight=1)

    coursetitle.grid(row=0,column=0,padx=0,pady=10,sticky="ew")
    cvl.grid(row=0,column=1,padx=20,pady=10,sticky="ew")
    Cp.grid(row=0,column=2,padx=10,pady=10,sticky="ew")
    Hw.grid(row=0,column=3,padx=10,pady=10,sticky="ew")
    CA_mark.grid(row=0,column=4,padx=10,pady=10,sticky="ew")
    Exam_mark.grid(row=0,column=5,padx=10,pady=10,sticky="ew")
    grade.grid(row=0,column=6,padx=20,pady=10,sticky="ew")
    Gp.grid(row=0,column=7,padx=20,pady=10,sticky="ew")

  
    #appending widgets to widget list

    CV_widgets.append(cvl)
    CP_widgets.append(Cp)
    HM_widgets.append(Hw)
    CA_widgets.append(CA_mark)
    Exam_widgets.append(Exam_mark)
    grades_Widgets.append(grade)
    GP_widgets.append(Gp)

    


totalCP_frame,totalHW_frame,totalCA_frame,totalExam_frame,totalAttemptedCredit_frame,totalCreditEarned_frame,gradePointAverage_frame = None,None,None,None,None,None,None


def applyGrades():
    semester = semester_Entry.get()
    year = yearEntry.get()
    department = department_Entry.get()
    program = program_entry.get()
    docs = courses.find({"semester":semester,"year":year,"department":department,"program":program})
    doc_count = 0
    for doc in docs:
        doc_count = doc_count + 1
    percent_count = doc_count * 100
    numb_failedCourses = 0
    failedcreditList = []
    length = len(CP_widgets)
    TotalCP_value =0
    TotalHw_value = 0
    TotalCA_value = 0
    TotalExam_value = 0
    TotalCV_value = 0
    TotalFailedCV_value = 0
    TotalPoint = 0
    index = 0
    for elem in range(0,length):
        try:
            val1 = float(CV_widgets[index].cget("text"))
            val2 = float(CP_widgets[index].get())
            val3 = float(HM_widgets[index].get())
            val4 = float(Exam_widgets[index].get())
        except ValueError as e:
            message = "Entries cannot be left empty or be strings!!! restart the application and try again."
            errLabel = tk.Label(table_frame,text=message,foreground="red",background="white", font=("Times New Roman",14))
            errLabel.grid(pady=30,sticky="ew",padx=30)
            # findcourse()
            
        CV_value = float(CV_widgets[index].cget("text"))
        CP_value = float(CP_widgets[index].get())
        Hw_value = float(HM_widgets[index].get())
        CA_value = float(CA_widgets[index].get())
        Exam_value = float(Exam_widgets[index].get())
        grade_widget = grades_Widgets[index]
        point_widget = GP_widgets[index]
        CV_widget = CV_widgets[index]
        docs = scale_docs.find({})


        TotalCP_value = TotalCP_value + CP_value
        TotalHw_value = TotalHw_value + Hw_value
        TotalCA_value = TotalCA_value + CA_value
        TotalExam_value = TotalExam_value + Exam_value
        TotalCV_value = TotalCV_value + CV_value

        x = ((CP_value + Hw_value)*30)/200 + ((CA_value + Exam_value)*70)/200
        index = index + 1

        for doc in docs:
            _range = doc['range']
            grade = doc['grade']
            point = doc['point']
            

            if eval(_range):
                grade_widget.configure(text=grade)
                point_widget.configure(text=point)
                if grade_widget["text"]=="F":
                    failedcreditList.append(int(CV_widget["text"]))
    TotalFailedCV_value =  sum(failedcreditList)
    for widget in GP_widgets:
        point = float(widget.cget("text"))
        TotalPoint = TotalPoint + point
        


    totalCP_label = tk.Label(totalCP_frame,fg="blue" ,bg="white",text=round((TotalCP_value/percent_count) *100,2),font=("Times New Romans",14))
    totalHW_label = tk.Label(totalHW_frame,fg="blue" ,bg="white",text=round((TotalHw_value/percent_count)*100,2),font=("Times New Romans",14))
    totalCA_label = tk.Label(totalCA_frame,fg="blue" ,bg="white",text=round((TotalCA_value/percent_count)*100,2),font=("Times New Romans",14))
    totalExam_label = tk.Label(totalExam_frame,fg="blue" ,bg="white",text=round((TotalExam_value/percent_count)*100,2),font=("Times New Romans",14))
    totalattempted_label = tk.Label(totalAttemptedCredit_frame,fg="blue" ,bg="white",text=TotalCV_value,font=("Times New Romans",14))
    totalearned_label = tk.Label(totalCreditEarned_frame,fg="blue" ,bg="white",text=TotalCV_value-TotalFailedCV_value,font=("Times New Romans",14))
    PointAverage_Label = tk.Label(gradePointAverage_frame,fg="blue" ,bg="white",text=round(TotalPoint/doc_count,2),font=("Times New Romans",14))

    totalCP_label.pack()
    totalHW_label.pack()
    totalCA_label.pack()
    totalExam_label.pack()
    totalattempted_label.pack()
    totalearned_label.pack()
    PointAverage_Label.pack()



#function to search courses offered for particular semester,year,depart and program
def findcourse():
    for widget in table_frame.winfo_children():
        widget.destroy()
    semester = semester_Entry.get()
    year = yearEntry.get()
    department = department_Entry.get()
    program = program_entry.get()
    foundCourses = courses.find({"semester":semester,"year":year,"department":department,"program":program})

    gradestudentButton = tk.Button(table_frame,command=applyGrades,text="Apply grades",bg="#65B741",font=("Times New Roman",14))
    gradestudentButton.grid(column=7,row=0,sticky="ew")
    table_header = tk.Frame(table_frame,bg="#3081D0",name="table_header")
    headerCoursetileAndCode = tk.Label(table_header,font=("Times New Roman",14),text="Course_Title",foreground="white",bg="#3081D0")
    headerCreditValue = tk.Label(table_header,font=("Times New Roman",14),text="CV",foreground="white",bg="#3081D0")
    headerGrade = tk.Label(table_header,font=("Times New Roman",14),text="Grade",foreground="white",bg="#3081D0")
    headerCAMark = tk.Label(table_header,font=("Times New Roman",14),text="CA_Mark/100",foreground="white",bg="#3081D0")
    headerExamMark = tk.Label(table_header,font=("Times New Roman",14),text="Exam Mark/100",foreground="white",bg="#3081D0")
    headerGradePoint = tk.Label(table_header,font=("Times New Roman",14),text="GP",foreground="white",bg="#3081D0")
    graded_assignments = tk.Label(table_header,font=("Times New Roman",14),text="HW/100",foreground="white",bg="#3081D0")
    class_participation = tk.Label(table_header,font=("Times New Roman",14),text="CP/100",foreground="white",bg="#3081D0")
    


    table_header.grid(sticky="nsew")
    table_header.columnconfigure(1,weight=1)
    table_header.columnconfigure(0,weight=0)
    headerCoursetileAndCode.grid(row=0,column=0,padx=40,pady=10,sticky="ew")
    headerCreditValue.grid(row=0,column=1,padx=40,pady=10,sticky="ew")
    class_participation.grid(row=0,column=2,padx=40,pady=10,sticky="ew")
    graded_assignments.grid(row=0,column=3,padx=40,pady=10,sticky="ew")
    headerCAMark.grid(row=0,column=4,padx=40,pady=10,sticky="ew")
    headerExamMark.grid(row=0,column=5,padx=40,pady=10,sticky="ew")
    headerGrade.grid(row=0,column=6,padx=40,pady=10,sticky="ew")
    headerGradePoint.grid(row=0,column=7,padx=40,pady=10,sticky="ew")
   


    for doc in foundCourses:
        id = doc["_id"]
        code=doc["code"]
        # year=doc["year"]
        # semester=doc["semester"]
        title=doc["course title"]
        # department=doc["department"]
        # program=doc["program"]
        cv = doc["cv"]
        createtable_row(title,cv,code)
    global Totals_frame

    Totals_frame = tk.Frame(table_frame,pady=20,bg="white",name="totals_frame")
    Totals_frame.grid(sticky="nsew")

    Text = tk.Label(Totals_frame,fg="red",font=("Times New Romans",14),text="TOTAL/100",padx=10)
    Text.grid(sticky="ew",column=0,row=0)

    global totalCP_frame,totalHW_frame,totalCA_frame,totalExam_frame,totalAttemptedCredit_frame,totalCreditEarned_frame,gradePointAverage_frame

    totalCP_frame = tk.Frame(Totals_frame,bg="white")
    totalHW_frame = tk.Frame(Totals_frame,bg="white")
    totalCA_frame = tk.Frame(Totals_frame,bg="white")
    totalExam_frame = tk.Frame(Totals_frame,bg="white")
    totalAttemptedCredit_frame = tk.Frame(Totals_frame,bg="white")
    totalCreditEarned_frame = tk.Frame(Totals_frame,bg="white")
    gradePointAverage_frame = tk.Frame(Totals_frame,bg="white")

    totalCP_frame.grid(row=0,column=1,sticky="ew",padx=10)
    totalHW_frame.grid(row=0,column=2,sticky="ew",padx=10)
    totalCA_frame.grid(row=0,column=3,sticky="ew",padx=10)
    totalExam_frame.grid(row=0,column=4,sticky="ew",padx=10)
    totalAttemptedCredit_frame.grid(row=0,column=5,sticky="ew",padx=10)
    totalCreditEarned_frame.grid(row=0,column=6,sticky="ew",padx=10)
    gradePointAverage_frame.grid(row=0, column=7,sticky="ew")


    
    CP_label = tk.Label(totalCP_frame,fg="#000000" ,bg="white",text="CP",font=("Times New Romans",14))
    HW_label = tk.Label(totalHW_frame,fg="#000000" ,bg="white",text="HW",font=("Times New Romans",14))
    CA_label = tk.Label(totalCA_frame,fg="#000000", bg="white",text="CA_mark",font=("Times New Romans",14))
    Exam_label = tk.Label(totalExam_frame,fg="#000000", bg="white",text="Exam_mark",font=("Times New Romans",14))
    attemptedCredit_label = tk.Label(totalAttemptedCredit_frame,fg="#000000", bg="white",text="Attempted_Credit",font=("Times New Romans",14))
    creditEarned_label = tk.Label(totalCreditEarned_frame,fg="#000000", bg="white",text="Credit_Earned",font=("Times New Romans",14))
    gradePointAverage_label = tk.Label(gradePointAverage_frame,fg="red", bg="white",text="GPA",font=("Times New Romans",14))

    CP_label.pack()
    HW_label.pack()
    CA_label.pack()
    Exam_label.pack()
    attemptedCredit_label.pack()
    creditEarned_label.pack()
    gradePointAverage_label.pack()
        
def pdfgenerate():
    doc = SimpleDocTemplate("report.pdf", pagesize=landscape(A3))

    # Create the table data as a nested list
    data = [["Course_title", "CV", "CP/100", "HW/100", "Grade", "CA/100", "Exam/100", "GP/100"]]

    for widget in table_frame.winfo_children():
        li = []
        if widget.winfo_name() == "totals_frame":
            totals = widget
            for element in totals.winfo_children():
                if isinstance(element, tk.Label):
                    li.append(element.cget("text"))
                elif isinstance(element, tk.Frame):
                    for item in element.winfo_children():
                        li.append(item.cget("text"))
                
        elif widget.winfo_name() != "table_header":
            for each in widget.winfo_children():
                if isinstance(each, tk.Label):
                    li.append(each.cget("text"))
                elif isinstance(each, tk.Entry):
                    li.append(each.get())


        data.append(li)

    # Define the table style
    table_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.green),  # Header row background color
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),  # Header row text color
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # Center align all cells
        ("FONTNAME", (0, 0), (-1, 0), "Times New Roman"),  # Header row font
        ("FONTSIZE", (0, 0), (-1, 0), 14),  # Header row font size
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),  # Header row bottom padding
        ("BACKGROUND", (0, 1), (-1, -1), colors.beige),  # Data row background color
        ("GRID", (0, 0), (-1, -1), 1, colors.black),  # Table grid color and thickness
    ])

    # Create the table and apply the style
    table = Table(data)
    table.setStyle(table_style)
    space = " "
    title = Paragraph(NameEntry.get()+(space*8)+"__"+idEntry.get(), getSampleStyleSheet()['Title'])
    elements = [title,table]
    doc.build(elements)

    # Save and close the PDF document
    filename = "report.pdf"
    try:
        subprocess.run(["start","", filename],shell=True)
    except FileNotFoundError as e:
        message = "file not found!"
        print(message)



    


def apply_changes():
    """function updates all documents in the scale collection"""

    newRange_list = []
    newGrade_list = []
    newpoint_list = []
    index = 0

    new_docs = []


    for widget in range_col.winfo_children():
        new_range = widget.get()
        newRange_list.append(new_range)
        
    for widget in grade_col.winfo_children():
        new_grade = widget.get()
        newGrade_list.append(new_grade)

    for widget in gradePoint_col.winfo_children():
        new_point = widget.get()
        newpoint_list.append(new_point)

    for elem in newRange_list:
        object = {"range":elem,"grade":newGrade_list[index],"point":newpoint_list[index]}
        new_docs.append(object)
        index = index + 1


    scale_docs.delete_many({})
    scale_docs.insert_many(new_docs)

    for widget in range_col.winfo_children():
        widget.destroy()
    for widget in grade_col.winfo_children():
        widget.destroy()
    for widget in gradePoint_col.winfo_children():
        widget.destroy()
    
    findScale() #reload the scale


def findScale():
    docs = scale_docs.find({})
    index = 0

    for doc in docs:
        range = doc["range"]
        grade = doc["grade"]
        point = doc["point"]

        range_entry = tk.Entry(range_col,font=("Times New Roman",14))
        range_entry.insert(0,range)
        grade_entry = tk.Entry(grade_col,font=("Times New Roman",14))
        grade_entry.insert(0,grade)
        point_entry = tk.Entry(gradePoint_col,font=("Times New Roman",14))
        point_entry.insert(0,point)

        range_entry.grid(row=index,sticky="ew")
        grade_entry.grid(row=index,sticky="ew")
        point_entry.grid(row=index,sticky="ew")

        index = index + 1



            
        
    




# Creating the main window
window = tk.Tk()
window.title("GRADX")
window.configure(bg="#FFFFFF")

window.iconbitmap('images/icon.ico')  # Replace 'path/to/logo.ico' with the actual path to your logo file

# Set the window attributes to remove the title bar and borders
window.overrideredirect(False)
# Get the screen width and height
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

# Set the window geometry to fit the screen
window.geometry(f"{screen_width}x{screen_height}")

container = tk.Canvas(window,highlightthickness=0)
container.pack(side="left",fill="both",expand="true")
yscrollbar = tk.Scrollbar(window,command=container.yview)
yscrollbar.pack(side="right",fill="y")
container.configure(yscrollcommand=yscrollbar.set)
container_frame = tk.Frame(container,bg="white")
container.create_window((0,0),window=container_frame,width=screen_width-10,height=screen_height*2)
container_frame.update_idletasks()
container.configure(scrollregion=container.bbox("all"))





# Create the header with a sky-blue background color

header = tk.Frame(container_frame, bg="white")
header.pack(side="top", fill="x")
navbar_frame = tk.Frame(header, bg="white")
navbar_frame.pack(padx=10,pady=10,fill="x")
schoolLogoFrame = tk.Frame(container_frame,bg="white")
schoolLogoFrame.pack(fill="x")
searchframe = tk.Frame(container_frame,bg="white")
searchframe.pack(pady=30)
studentinfo_frame = tk.Frame(container_frame,bg="white")
studentinfo_frame.pack(pady=30)
student_Name_Matricule = tk.Frame(searchframe,bg="white")
studName = tk.Label(student_Name_Matricule,text="Stud_Name:",font=("Times New Romans",14))
NameEntry = tk.Entry(student_Name_Matricule,font=("Times New Romans",14))
studID = tk.Label(student_Name_Matricule,text="Stud_ID:",font=("Times New Romans",14))
idEntry = tk.Entry(student_Name_Matricule,font=("Times New Romans",14))

student_Name_Matricule.grid(sticky="nsew",row=0,column=4,padx=30,pady=30)
studName.grid(column=0,row=0,sticky="ew")
NameEntry.grid(column=1, row=0,sticky="ew")
studID.grid(column=0,row=1)
idEntry.grid(column=1,row=1)






#create photo icons for widget 
icon = tk.PhotoImage(file="images/resized.png")
folder_icon = tk.PhotoImage(file="images/touchfolder.png")
print_icon = tk.PhotoImage(file="images/printretouch.png")
share_icon = tk.PhotoImage(file="images/shareicon.png")
school_logo = tk.PhotoImage(file="images/citec.png")



# Create the toolbar widgets
save_button = tk.Button(navbar_frame, text="Save",bg="white", image=folder_icon, command=savefile)
share_button = tk.Button(navbar_frame, text="Share", image=share_icon, bg="white")
print_button = tk.Button(navbar_frame, text="Print", image=print_icon, bg="white",command=pdfgenerate)
label1 = tk.Label(schoolLogoFrame,image=school_logo)

#widgets for student information frame
yearlevel_label = tk.Label(searchframe,text="year Level:",font=("Times New Roman",14),bg="white")
yearEntry = tk.Entry(searchframe,font=("Times New Roman",14))
program_label = tk.Label(searchframe,text="Program:",font=("Times New Roman",14),bg="white")
program_entry = tk.Entry(searchframe,font=("Times New Roman",14))
department_label=tk.Label(searchframe,text="Department:",font=("Times New Roman",14),bg="white")
department_Entry = tk.Entry(searchframe ,font=("Times New Roman",14))
semester_label = tk.Label(searchframe,text="semester:",font=("Times New Roman",14),bg="white")
semester_Entry = tk.Entry(searchframe,font=("Times New Roman",14))
searchcontainer = tk.Frame(searchframe,padx=20,bg="white")
searchcontainer.grid(row=0,column=2)
searchcourse_button = tk.Button(searchcontainer,text="search course",font=("Times New Roman",14),command=findcourse)
table_frame = tk.Frame(container_frame,padx=5,pady=5,width=0.7*screen_width)
table_frame.pack()
# table_frame = TableFrameWidget(container_frame,container_frame=container_frame, padx=5, pady=5, width=0.7 *screen_width)






#packing widgets using grid layout manager
yearlevel_label.grid(row=0,column=0)
yearEntry.grid(row=0,column=1)
program_label.grid(row=1,column=0)
program_entry.grid(row=1,column=1)
department_label.grid(row=2,column=0)
department_Entry.grid(row=2,column=1)
searchcourse_button.grid(row=0,column=2)
semester_label.grid(row=3,column=0)
semester_Entry.grid(row=3,column=1)



# Pack the toolbar widgets
save_button.pack(side="left", padx=10, pady=5)
share_button.pack(side="right")
print_button.pack(side="right")
label1.pack()

#grade scale widgets
gradeScale_frame = tk.Frame(container_frame,bg="white",pady=50)
gradeScale_frame.pack()
admin_text = tk.Text(container_frame,font=("Times New Romans",14),foreground="red",height=10)
text = "Please note that this area is strictly for admins only. changing anything might lead to unwanted results during grading process."
admin_text.insert(1.0,text)
range_col = tk.Frame(gradeScale_frame,padx=5,pady=5)
grade_col = tk.Frame(gradeScale_frame,padx=5,pady=5)
gradePoint_col = tk.Frame(gradeScale_frame,padx=5,pady=5)
applyChanges_button = tk.Button(gradeScale_frame,text="update_scale",command=apply_changes,font=("Times New Roman",14))


admin_text.pack()
range_col.grid(sticky="ns",column=0,row=1)
grade_col.grid(sticky="ns",column=1,row=1)
gradePoint_col.grid(sticky="ns",column=2,row=1)
applyChanges_button.grid(sticky="e",row=2,column=3)

findScale()


# Start the main loop
window.mainloop()