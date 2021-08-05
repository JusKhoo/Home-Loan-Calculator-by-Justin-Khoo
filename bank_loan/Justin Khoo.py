import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
from PIL import Image

def disableEvent(event):
    if my_tree.identify_region(event.x, event.y) == 'separator':
        return "break"
def instruction():
    instructions.deiconify()
def Next_Ok():
    global y
    y=y+1
    if y<5:
        instructions_lbl.config(state= NORMAL)
        img = Image.open(list8[y])
        width, height = img.size
        instructions_lbl.delete('1.0',END)
        instructions_lbl.insert(END, list6[y])
        instruction_lbl.config(image= list7[y])
        instruction_lbl.place(x= ((579-width)/2), y=315)
        instructions_lbl.config(state= DISABLED)
    else:
        y=0
        instructions_lbl.config(state= NORMAL)
        instructions_lbl.delete('1.0',END)
        instructions_lbl.insert(END, list6[0])
        instruction_lbl.config(image= list7[0])
        instructions_lbl.config(state= DISABLED)
        instruction_lbl.place(x=147, y= 315)
        next_ok.config(image=nextimg)
        instructions.withdraw()
    if y == 4:
        next_ok.config(image=okimg)
def close_Window_B():
    WindowB.withdraw()
def close_Window_C():
    WindowC.withdraw()
def close_instructions():
    global y
    y=0
    instructions_lbl.config(state= NORMAL)
    instructions_lbl.delete('1.0',END)
    instructions_lbl.insert(END, list6[0])
    instruction_lbl.config(image= list7[0])
    instructions_lbl.config(state= DISABLED)
    instruction_lbl.place(x=147, y= 315)
    next_ok.config(image=nextimg)
    instructions.withdraw()
def enter(event,num):
    if "Please re-enter." in l[num].get():
        l[num].set("")
        l2[num].config(fg="black", font = ("Times New Roman",14,'bold'))

def isfloat(x):
    try:
        float(x)
        return True
    except ValueError:
        return False
    
def calculations(a,b,c,d,e,f):
    mi = c*b
    principal = d-mi
    e = e + principal
    f = f + mi
    c = abs(c - principal)
    list1.append(a+1)
    list2.append("%.2f" % (principal))
    list3.append("%.2f" % (mi))
    list4.append("%.2f" % (c))
    return c, e, f
    
def check():
    WindowB.withdraw()
    WindowC.withdraw()
    my_tree.delete(*my_tree.get_children())
    status = True
    WindowA.focus_force()
    for i in range(4):
        if isfloat((l[i]).get()) == False:
            l[i].set("Non numerical value inputed. Please re-enter.")
            l2[i].config(fg = "red", font= ("Times New Roman",10,'bold'))
            status = False
        else:
            if float((l[i]).get()) < 0:
                l[i].set("Value below 0. Please re-enter.")
                l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
                status = False
            elif i == 1 or i == 2:
                if float((l[i]).get()) > 100:
                    l[i].set("Value exceeded 100. Please re-enter.")
                    l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
                    status = False

    if status == False:
        return None
    else:
        my_tree.yview_moveto(0)
        list1.clear()
        list2.clear()
        list3.clear()
        list4.clear()
        list5.clear()
        loan = float(str_loan.get())
        down_payment = float(str_down_payment.get())
        interest = float(str_interest.get())
        years = float(str_year.get())
        loan = loan - (loan*down_payment)/100
        mr = interest/12/100
        months = int(years*12)
        totalprincipal = 0
        totalinterest = 0
        if float(str_loan.get()) > 100000000:
            l[0].set("Value too high. Please re-enter.")
            l2[0].config(fg = "red", font= ("Times New Roman",11,'bold'))
            try:
                monthly_repayment = loan/(((((1+mr)**months)-1))/(mr*(1+mr)**months))
            except OverflowError:
                l[i].set("Value too high. Please re-enter.")
                l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
                return None
            except ZeroDivisionError:
                l[i].set("0 is entered. Please re-enter.")
                l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
                return None
            return None
        try:
            monthly_repayment = loan/(((((1+mr)**months)-1))/(mr*(1+mr)**months))
        except OverflowError:
            l[i].set("Value too high. Please re-enter.")
            l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
            return None
        except ZeroDivisionError:
            l[i].set("0 is entered. Please re-enter.")
            l2[i].config(fg = "red", font= ("Times New Roman",11,'bold'))
            return None

        for i in range(months):
            loan,totalprincipal,totalinterest = calculations(i,mr,loan,monthly_repayment,totalprincipal, totalinterest)
        lbl_monthly_repayment.config(text = format(monthly_repayment,".02f"))
        lbl_total_principal.config(text = format(totalprincipal,".02f"))
        lbl_total_interest.config(text = format(totalinterest,".02f"))
        list5.append("%.2f" % (monthly_repayment))
        list5.append("%.2f" % (totalprincipal))
        list5.append("%.2f" % (totalinterest))
        WindowB.deiconify()
        WindowB.protocol("WM_DELETE_WINDOW", close_Window_B)
        
def analysis():
    WindowC.deiconify()
    WindowC.protocol("WM_DELETE_WINDOW", close_Window_C)
    for i in range(list1[-1]):
        my_tree.insert(parent='', index='end', iid=i, text='', value=(list1[i],list2[i],list3[i],list4[i]))

def download():
    list1.append("")
    list1.append("Monthly repayment:")
    list1.append("Total principal:")
    list1.append("Total interest:")
    list2.append("")
    list2.append(list5[0])
    list2.append(list5[1])
    list2.append(list5[2])
    for i in range(4):
        list3.append("")
        list4.append("")
    data = {'Period': list1,
            'Principal': list2,
            'Interest rate': list3,
            'Balance': list4
            }
    df = pd.DataFrame(data, columns = ['Period', 'Principal','Interest rate','Balance'])
    current_directory = filedialog.askdirectory()
    file_name = str(str_loan.get()+'_'+str_down_payment.get()+'_'+str_interest.get()+'_'+str_year.get()+'.xlsx')
    df.to_excel(os.path.join(current_directory,file_name), index = False, header=True)
    if current_directory != '':
        messagebox.showinfo("Information", "Analysis downloaded successfully.")
    current_directory = ''
        
WindowA= Tk()
WindowA.title("HOME LOAN CALCULATOR")
WindowA.geometry("867x650")
WindowA.minsize(867,650)
WindowA.maxsize(867,650)

icon= PhotoImage(file= "circle piano.png")
WindowA.iconphoto(False, icon)

WindowAbg = PhotoImage(file="money tree.png")
background = Label(WindowA, image=WindowAbg)
background.place(x=0,y=0)

frame = PhotoImage(file="frame.png")
framelbl = Label(WindowA, image=frame,border=0)
framelbl.place(x=34, y=54)

str_loan=StringVar()
e_loan=Entry(WindowA,textvariable=str_loan,bg="#e5e5e5",border=0,font=("Times New Roman",14,'bold'))
e_loan.place(x=75,y=220,width=350,height=30)
e_loan.config(state=NORMAL)
e_loan.bind("<FocusIn>", lambda event: enter(event, 0))

str_down_payment=StringVar()
e_down_payment=Entry(WindowA,textvariable=str_down_payment,bg="#e5e5e5",border=0,font=("Times New Roman",14,'bold'))
e_down_payment.place(x=75,y=303,width=350,height=30)
e_down_payment.config(state=NORMAL)
e_down_payment.bind("<FocusIn>", lambda event: enter(event, 1) )

str_interest=StringVar()
e_interest=Entry(WindowA,textvariable=str_interest,bg="#e5e5e5",border=0,font=("Times New Roman",14,'bold'))
e_interest.place(x=75,y=386,width=350,height=30)
e_interest.config(state=NORMAL)
e_interest.bind("<FocusIn>", lambda event: enter(event, 2))

str_year=StringVar()
e_year=Entry(WindowA,textvariable=str_year,bg="#e5e5e5",border=0,font=("Times New Roman",14,'bold'))
e_year.place(x=75,y=467,width=350,height=30)
e_year.config(state=NORMAL)
e_year.bind("<FocusIn>", lambda event: enter(event, 3))

y=0
l = [str_loan,str_down_payment,str_interest,str_year]
l2 = [e_loan,e_down_payment,e_interest,e_year]

list1 = []
list2 = []
list3 = []
list4 = []
list5 = []

confirm = PhotoImage(file="confirm button.png")
confirmbtn = Button(WindowA, image=confirm, highlightthickness = 0, bd=0, activebackground = "#919193",command= check)
confirmbtn.place(x=70, y=545)

instructionbtnimg = PhotoImage(file="instruction.png")
instructionbtn = Button(WindowA, image=instructionbtnimg, highlightthickness = 0, bd=0, activebackground = "#767c79",command= instruction)
instructionbtn.place(x=760, y=32)

instructions = Toplevel(WindowA, width=579, height=650)
instructions.minsize(579,650)
instructions.maxsize(579,650)
instructions.iconphoto(False, icon)
instructions['background'] = '#F0F0F0'
instructions.protocol("WM_DELETE_WINDOW", close_instructions)

instructionbgimg = PhotoImage(file="bg1.png")
instructionbg = Label(instructions, image=instructionbgimg)
instructionbg.place(x=0, y=0)

instructions_lbl = Text(instructions, border = 0, bg = '#FFFFFF', font=("Times New Roman",16,'bold'))
text_1 = """Here is a brief tutorial to show you how to use
my system.

Step 1
Please fill in the information in the text boxes
for each of the following, which are your loan
amount, down payment, interest rate and year.
Please ensure you fill in the correct and
suitable numerical inputs in this system. Once
you are sure your data is correct, click the
CONFIRM button."""

text_2 = """
Step 2
After clicking the CONFIRM button, you will
be brought to the second page. Your results
for monthly repayment, total principal and
total interest will be displayed. If you wish
for a more detailed analysis on your loan, you
may click the Analysis button. If not, you may
close the tab.
"""

text_3 = """
Step 3
You will reach the third page after clicking
the Analysis button. Your detailed analysis
for your loan will be displayed there. You may
scroll through the analysis that shows the
principal, interest rate and balance after
every month. You may also choose to download
the analysis in Excel format. If you wish to
do so, click the DOWNLOAD button. If not,
close the tab. 
"""

text_4 = """
Step 4
When you click the Download Button, you will
be requested to select a folder to store your
excel file. Select a folder. You will be notified
when your analysis downloaded succesfully.
"""

text_5 = """
Step 5
Lastly, just go back to the folder you selected
earlier to find your Excel file. Click into the
Excel file and you will see the full analysis
succesfully downloaded.

Hope you enjoy using this program!
"""
instructions_lbl.insert(END, text_1)
instructions_lbl.config(state = DISABLED)

instructions_lbl.place(x=80, y=40, width= 450, height= 280)
instructions.withdraw()

instructionimg1 = PhotoImage(file=r"instruction1.png")
instructionimg2 = PhotoImage(file=r"instruction2.png")
instructionimg3 = PhotoImage(file=r"instruction3.png")
instructionimg4 = PhotoImage(file=r"instruction4.png")
instructionimg5 = PhotoImage(file=r"instruction5.png")

list6 = [text_1,text_2,text_3,text_4,text_5]
list7 = [instructionimg1,instructionimg2,instructionimg3,instructionimg4,instructionimg5]
list8 = ['instruction1.png','instruction2.png','instruction3.png','instruction4.png','instruction5.png']

instruction_lbl = Label(instructions, image= instructionimg1)
instruction_lbl.place(x=147, y= 315)

nextimg = PhotoImage(file= r"next.png")
okimg = PhotoImage(file= r"OK.png")
next_ok = Button(instructions, image=nextimg, border=0, background= "#FFFFFF", activebackground="#FFFFFF", command = Next_Ok)
next_ok.place(x=415, y= 560)

WindowB = Toplevel(WindowA,width=900, height=650)
WindowB.minsize(900,650)
WindowB.maxsize(900,650)
WindowB.iconphoto(False,icon)
WindowBbg = PhotoImage(file= "house.png")
background2 = Label(WindowB, image=WindowBbg)
background2.place(x=0,y=0)

Resultsimg = PhotoImage(file = "results.png")
Results = Label(WindowB, image=Resultsimg, bd=0 )
Results.place(x=14,y=18)

Totalimg = PhotoImage(file = "total.png")
Total = Label(WindowB, image=Totalimg, bd=0 )
Total.place(x=33,y=141)

analysisbtnimg = PhotoImage(file = "analysis button.png")
analysisbtn = Button(WindowB, image=analysisbtnimg, bd=0, highlightthickness = 0, activebackground = "#94cfd5", command = analysis )
analysisbtn.place(x=125,y=530)

lbl_monthly_repayment = Label(WindowB, bg = "#f8f8ff", text = "",font=("Times New Roman",14,'bold'))
lbl_monthly_repayment.place(x=70, y=212, width = 353, height = 40)

lbl_total_principal = Label(WindowB, bg = "#f8f8ff", text = "",font=("Times New Roman",14,'bold'))
lbl_total_principal.place(x=70, y=322, width = 353, height = 40)

lbl_total_interest = Label(WindowB, bg = "#f8f8ff", text = "",font=("Times New Roman",14,'bold'))
lbl_total_interest.place(x=70, y=432, width = 353, height = 40)

WindowB.withdraw()

WindowC = Toplevel(WindowA,width=609, height=650)
WindowC.minsize(609,650)
WindowC.maxsize(609,650)
WindowC.iconphoto(False,icon)
WindowCbg = PhotoImage(file= "WindowC.png")
background3 = Label(WindowC, image=WindowCbg)
background3.place(x=0,y=0)

style = ttk.Style()
style.theme_use("clam")
style.configure("mystyle.Treeview", background= "#A9A9A9", foreground="black", fieldbackground= "#A9A9A9", font= ("Times New Roman", 13, "bold"), rowheight = 33)
style.configure("mystyle.Treeview.Heading", background = '#404040', foreground = 'white', font = ("Times New Roman", 11, "bold"), rowheight = 40)
style.map("mystyle.Treeview",background=[('selected','#404040')],foreground=[('selected','white')])
style.map("mystyle.Treeview.Heading",background=[('disabled','#404040')],foreground=[('disabled','white')])

tree_frame = Frame(WindowC)
tree_frame.place(x=20, y=205, width=568, height= 311)

tree_scroll = Scrollbar(tree_frame, orient=VERTICAL)
tree_scroll.pack(side=RIGHT, fill = Y)

my_tree = ttk.Treeview(tree_frame,style='mystyle.Treeview', yscrollcommand = tree_scroll.set)
my_tree['columns'] = ("Period","Principal","Interest rate","Balance")
my_tree.bind("<Button-1>", disableEvent)
tree_scroll.config(command = my_tree.yview)

my_tree.column("#0", width=0, stretch=NO)
my_tree.column("Period", anchor= CENTER, width=108, minwidth=108)
my_tree.column("Principal", anchor= CENTER, width=143,minwidth=143)
my_tree.column("Interest rate", anchor= CENTER, width=143,minwidth=143)
my_tree.column("Balance", anchor= CENTER, width=143,minwidth=143)

my_tree.heading("#0", text="", anchor= CENTER)
my_tree.heading("Period", text="Period", anchor= CENTER)
my_tree.heading("Principal", text="Principal", anchor= CENTER)
my_tree.heading("Interest rate", text="Interest rate", anchor= CENTER)
my_tree.heading("Balance", text="Balance", anchor= CENTER)

my_tree.place(x=0, y=0, height= 311)

downloadbtnimg = PhotoImage(file = "download.png")
downloadbtn = Button(WindowC, image= downloadbtnimg, border=0 ,highlightthickness = 0, activebackground = "#ffcccb", command = download)
downloadbtn.place(x=133, y=535)

WindowC.withdraw()
WindowA.mainloop()


