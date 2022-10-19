import tkinter as tk
from tkinter import filedialog
import pandas as pd
from fpdf import FPDF
from PIL import ImageTk,Image

#Create an instance of tkinter frame
splash_win= tk.Tk()
splash_win.configure(bg='white')
#splash_win.attributes('-fullscreen', True)
#Set the title of the window

#image=ImageTk.PhotoImage(Image.open("logo.png"))
#logo = tk.Label(splash_win,image =image,borderwidth=0)
#logo.place(x=400,y=200)

#Define the size of the window or frame
splash_win.geometry("1000x500+250+250")

#Remove border of the splash Window
splash_win.overrideredirect(True)

#Define the label of the window
splash_label= tk.Label(splash_win, text= "WISDOM TESTS\n\n\nSTUDENT REPORT CARD GENERATOR",bg="white",fg= "brown",font= ('Times New Roman', 40)).pack(pady=20)

class report_card_generator:
    def __init__(self):
        splash_win.destroy()
        self.win = tk.Tk()
        self.win.geometry("%dx%d" % (self.win.winfo_screenwidth(), self.win.winfo_screenheight()))
        # win.configure(bg='white')
        self.win.update()
        self.frame = tk.Frame(self.win, bg='orange', width=292, height=self.win.winfo_height())
        self.frame.grid(row=0, column=0)
        self.frame.grid_propagate(False)

        # win.configure(bg='#b656e9')
        # win.resizable(False, False)
        self.win.title("STUDENT REPORT CARD GENERATOR - Main Menu")
        self.issue_book = tk.Button(self.win, text="Generate Reports", bg="orange", bd=0, height=3, width=26, font=(0, 15),command=lambda: self.generate_report_callback(self.win))
        self.issue_book.bind("<Enter>", self.on_enter)
        self.issue_book.bind("<Leave>", self.on_leave)
        self.issue_book.place(x=0, y=0)

        self.return_book = tk.Button(self.win, text="Developer Info", bg="orange", bd=0, height=3, width=26, font=(0, 15),command=lambda: self.developer_info_callback(self.win))
        self.return_book.bind("<Enter>", self.on_enter)
        self.return_book.bind("<Leave>", self.on_leave)
        self.return_book.place(x=0, y=75)
        self.win.mainloop()

    def generate_report_callback(self,win):
        resp = filedialog.askopenfilenames(parent=win,initialdir='/', initialfile='tmp',filetypes=[("xlsx", "*.xlsx")])
        file_loc = str(resp[0])
        data = pd.read_excel(file_loc)
        names=data['Full Name '].unique()
        #number_of_studs = len(data['Full Name '].unique())
        for name in names:
            student_data=data[data['Full Name ']==name]
            registration_number=student_data['Registration Number'].unique()
            grade = student_data['Grade '].unique()
            school= student_data['Name of School '].unique()
            gender=student_data['Gender'].unique()
            dob=student_data['Date of Birth '].unique()
            residence = student_data['City of Residence'].unique()
            test_round=student_data['Round'].unique()
            test_date=student_data['Date and time of test'].unique()
            country = student_data['Country of Residence'].unique()
            correct_ans = student_data['Outcome (Correct/Incorrect/Not Attempted)'].value_counts()['Correct']
            incorrect_ans = student_data['Outcome (Correct/Incorrect/Not Attempted)'].value_counts()['Incorrect']
            total_score = student_data['Your score'].sum()
            final_result = student_data['Final result'].unique()
            self.generate_pdf(name,registration_number,grade,school,gender,dob,residence,test_round,test_date,country,correct_ans,incorrect_ans,total_score,final_result)
        self.status_lable = tk.Label(win, text="REPORTS GENERATED SUCCESSFULLY")
        self.status_lable.place(x=750, y=400)


    def generate_pdf(self,name,registration_number,grade,school,gender,dob,residence,test_round,test_date,country,correct_ans,incorrect_ans,total_score,final_result):
        pdf = FPDF('P', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font('Arial', 'B', 12)
        file_name=f"./Pics for assignment/{name}.png"
        pdf.cell(200, 10, "REPORT CARD", ln=1,align = 'C')
        pdf.image(file_name,x=160,y=10,w=30,h=30)
        pdf.cell(10,10, f"NAME : {name}",ln=True)
        pdf.cell(10,10, f"REGISTRATION NUMBER : {registration_number[0]}", ln=True)
        pdf.cell(10,10, f"GRADE : {grade[0]} ", ln=True)
        pdf.cell(10,10, f"SCHOOL  : {school[0]}", ln=True)
        pdf.cell(10,10, f"GENDER : {gender[0]}", ln=True)
        pdf.cell(10,10, f"DATE OF BIRTH  : {dob[0]}", ln=True)
        pdf.cell(10,10, f"RESIDENCE : {residence[0]}", ln=True)
        pdf.cell(10,10, f"COUNTRY :  : {country[0]}", ln=True)
        pdf.cell(10,10, f"TEST ROUND  : {test_round[0]}", ln=True)
        pdf.cell(10,10, f"TEST DATE  : {test_date[0]}", ln=True)
        pdf.set_text_color(0, 0, 255)
        pdf.cell(10,10, f"CORRECT ANSWERS : {correct_ans}", ln=True)
        pdf.cell(10,10, f"INCORRECT ANSWERS : {incorrect_ans}", ln=True)
        pdf.cell(10,10, f"TOTAL SCORE  : {total_score}", ln=True)
        pdf.cell(10,10, f"FINAL RESULT : {final_result[0]}", ln=True)
        pdf.output(f'{name}.pdf', 'F')

    def developer_info_callback(self,win):
        image = Image.open("./Pics for assignment/Himmalay_Devulapalli.png")
        image = image.resize((670, 810), Image.ANTIALIAS)
        image=ImageTk.PhotoImage(image)
        res = tk.Label(win,image =image)
        res.photo=image
        res.place(x=600,y=20)

    def on_enter(self,e):
        e.widget['background'] = 'ivory1'

    def on_leave(self,e):
        e.widget['background'] = 'orange'

splash_win.after(2000,report_card_generator)
splash_win.mainloop()
