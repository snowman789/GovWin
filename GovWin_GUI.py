import Manipulate_Data




from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter as tk
from tkinter import filedialog
from tkinter import *

class File_path:
    def __init__(self):
        self.file_path = r''
    def set_path(self, new_path):
        self.file_path = new_path

old_file_path = File_path()
save_file_path = File_path()

def Capture_old_path(old_file_path : File_path):
    user_path = filedialog.askopenfilename(initialdir = "/",title = "Select Settings File",filetypes = (("Comma Separated Values","*.xls"),("all files","*.*")))
    old_file_path.set_path( user_path)
    tk.Label(master,
             text= user_path).grid(row=0, column=1)




def Capture_save_path(save_file_path : File_path):
    user_path = filedialog.asksaveasfilename(initialdir = "/",title = "Select file",filetypes = (("Excel Files","*.xlsx"),("all files","*.*")))
    save_file_path.set_path(user_path + '.xlsx')
    tk.Label(master,
             text= user_path + '.xlsx').grid(row=1, column=1)

def Generate_report():
     # old_file = e1.get()
     # new_file = e2.get()
     # report_locatoin = e3.get()

     results = Manipulate_Data.Manipulate_Data(old_file_path.file_path, save_file_path.file_path)
     tk.Label(master,
              text=results).grid(row=3, column=1)



master = tk.Tk()


# tk.Label(master,
#          text="Enter date to start search: [MM/DD/YYYY] ").grid(row=0)
#
# e1 = tk.Entry(master)
# e1.grid(row=0, column=1)

tk.Label(master,
         text="").grid(row=0,
                       column = 1)
# tk.Label(master,
#          text="").grid(row=1,
#                        column = 1
#                        )
tk.Label(master,
         text=".xlsx").grid(row=1,
                       column = 1)

# e1 = tk.Entry(master)
# e2 = tk.Entry(master)
# e3 = tk.Entry(master)
#
# e1.grid(row=0, column=1)
# e2.grid(row=1, column=1)
# e3.grid(row=2, column=1)


tk.Button(master,
          text='Select Data',
          command= lambda: Capture_old_path(old_file_path)).grid(row=0,
                                    column=0,
                                    sticky=tk.W,
                                    pady=4)

tk.Button(master,
          text='Select Save Location',
          command=lambda: Capture_save_path(save_file_path)).grid(row=1,
                                    column=0,
                                    sticky=tk.W,
                                    pady=4)

tk.Button(master,
          text='Generate Report', command=Generate_report).grid(row=2,
                                                       column=1,
                                                       sticky=tk.W,
                                                       pady=4)


tk.mainloop()