from tkinter import *
import os
import Control_Flow as main_program
# Design GUI

def buttonFunc():
    path_value= str(folder.get()).replace('\\','/')+'/'
    print ('Working ...')
    try:
        main_program.main(path_value)
    except Exception as ex:
        print ('Error: '+ex)

# Create a window
window = Tk()
window.title('Extracting BEC Files')
window.geometry('400x100')

# Create Label
theLabel = Label(window, text='Enter the folder containing input data').pack()
folder = StringVar()
entry_box = Entry(window,textvariable=folder).pack()
button = Button(window,text='Execute',command=buttonFunc).pack()


window.mainloop()