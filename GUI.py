from tkinter import *
# import os
import Control_Flow as main_program
# import subprocess as sub
# p = sub.Popen('./script',stdout=sub.PIPE,stderr=sub.PIPE)
# output,errors=p.communicate()

folder = ''

# Function to execute when pressing the button
def buttonFunc():
    global folder
    # Convert input to the right format
    path_value= str(folder.get()).replace('\\','/')+'/'
    print ('Working ...')
    try:
        # Running the main function
        main_program.main(path_value)
    except Exception as ex:
        print ('Error: '+ex)

def main():
    global folder
    # Create a window
    window = Tk()
    window.title('Extracting BEC Files')
    window.geometry('400x100')

    # Create Label
    theLabel = Label(window, text='Enter the folder containing input data').pack()

    # Create Input
    folder = StringVar()
    entry_box = Entry(window,textvariable=folder).pack()

    # Terminal
    # text=Text(window)
    # text.pack()
    # text.insert(END,output)

    # Execution button
    button = Button(window,text='Execute',command=buttonFunc).pack()
    window.mainloop()

if __name__=='__main__':
    main()