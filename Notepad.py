from tkinter import *
import tkinter.messagebox as tkmsg
from tkinter.filedialog import askopenfilename, asksaveasfilename
import os
import threading
from win32com.client import Dispatch
import pythoncom


root = Tk()
root.geometry("900x500")
root.title("Untitled - NotePad by Aditya")
root.wm_iconbitmap('2.ico')

def Newfile():
    global file
    global status_var
    status_var.set('New File')
    root.title('Untitled - NotePad by Aditya')
    file = None
    work_space.delete(1.0, END)


def Openfile():
    global file
    global status_var
    status_var.set('Reading...')
    
    file = askopenfilename(defaultextension = ".txt", filetypes = [("All Files", "*.*"), ("Text Documents", "*.txt")])
    if file == "":
        
        file = None

    else:
        root.title(os.path.basename(file) + "- NotePad by Aditya")
        work_space.delete(1.0, END)
        global f
        f = open(file, 'r')
        work_space.insert(1.0, f.read())
        status_var.set('Showing Old File')
        f.close()

stopper = False
def stop():
    pass


def speaker():
    global stopper
    global work_space
    global x1_id
    data = work_space.get(1.0, END)


    speak = Dispatch("SAPI.SpVoice")
    x1_id = pythoncom.CoMarshalInterThredInterFaceInStream(pythoncom.IID_IDispatch, speak)
    t2 = threading.Thread

    speak.Speak(data)


def start():
    t1 = threading.Thread(target=speaker)
    t1.start()


def Savefile():
    global file
    global status_var

    if file == None:
        file = asksaveasfilename(initialfile = 'Untitled.txt', defaultextension = ".txt", filetypes = [("All Files", "*.*"), ("Text Documents", "*.txt")])
        if file == "":
            status_var.set('Saving...')
            file = None

        else:
            f = open(file, 'w')
            f.write(work_space.get(1.0, END))
            status_var.set('Saved !')
            f.close()

            root.title(os.path.basename(file) + "-NotePad by Aditya")


def Cut():
    work_space.event_generate(("<<Cut>>"))


def Copy():
    work_space.event_generate(("<<Copy>>"))


def Paste():
    work_space.event_generate(("<<Paste>>"))


def about():
    tkmsg.showinfo('About', 'This Notepad is developed by Aditya...')


def contact():
    response = tkmsg.askyesno('Contact Us', 'Do you want to contact the developer ?')
    if response:
        tkmsg.showinfo('Contact Details', 'Contact Us on adityayadav1743@gmail.com')
    
    else:
        tkmsg.showinfo('From Developer', 'Thankyou for using our NotePad please rate us on appstore')


def dark():
    color_var.set('gray26')
    statusbar_clr_var.set('gray26')
    fg_color.set('white')
    work_space.config(bg = color_var.get(), fg = fg_color.get())
    caption.config(bg = statusbar_clr_var.get(), fg = fg_color.get())
    status.config(bg = statusbar_clr_var.get(), fg = fg_color.get())





def light():
    color_var.set('white')
    statusbar_clr_var.set('white')
    fg_color.set('black')
    work_space.config(bg=color_var.get(), fg=fg_color.get())
    caption.config(bg=statusbar_clr_var.get(), fg=fg_color.get())
    status.config(bg=statusbar_clr_var.get(), fg=fg_color.get())


menubar = Menu(root)
filemenu = Menu(menubar)
filemenu.add_command(label = 'New', command = Newfile)
filemenu.add_command(label = 'Open', command = Openfile)
filemenu.add_command(label = 'Save', command = Savefile)


menubar.add_cascade(label = 'File', menu = filemenu)

editmenu = Menu(menubar)
editmenu.add_command(label = 'Cut', command = Cut)
editmenu.add_command(label = 'Copy', command = Copy)
editmenu.add_command(label = 'Paste', command = Paste)

menubar.add_cascade(label = 'Edit', menu = editmenu)

theme_menu = Menu(menubar)
theme_menu.add_command(label = 'Dark Mode', command = dark)
theme_menu.add_command(label = 'Light Mode', command = light)

menubar.add_cascade(label = 'Modes', menu = theme_menu)

helpmenu = Menu(menubar)
helpmenu.add_command(label = 'About', command = about)
helpmenu.add_command(label = 'ContactUs', command = contact)

menubar.add_cascade(label = 'Help', menu = helpmenu)

speak_menu = Menu(menubar)
speak_menu.add_cascade(label = "Text to Speech", command = start)
speak_menu.add_cascade(label = "Stop", command = stop)

menubar.add_cascade(label = 'Read', menu = speak_menu)

root.config(menu = menubar)


scroll = Scrollbar(root )
scroll.pack(fill = Y, side = RIGHT)

color_var = StringVar()
color_var.set('white')

fg_color = StringVar()
fg_color.set('black')

work_space = Text(root, font = 'lucida 10', yscrollcommand = scroll.set, bg= color_var.get(), fg = fg_color.get(), padx = 10)
work_space.pack(fill = BOTH, expand = True)
file = None

scroll.config(command = work_space.yview, bg = color_var.get())

statusbar_clr_var = StringVar()
statusbar_clr_var.set('white')

statusbar = Frame(root, borderwidth = 2, relief = GROOVE)
statusbar.pack(side = BOTTOM, fill = X)


caption = Label(statusbar, text = 'NotePad by Aditya ', font = 'ubantu 10 bold', padx = 680, bg = statusbar_clr_var.get(), fg = fg_color.get())
caption.grid(row = 0, column = 10)

status_var = StringVar()
status_var.set('Ready...')
status = Label(statusbar, textvariable = status_var, bg = statusbar_clr_var.get(), fg = fg_color.get())
status.grid(row = 0, column = 0)



root.mainloop()