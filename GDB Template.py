from __future__ import division
import sys
import os

#Add Module folder to path.
    #Module folder contains openpyxl instead of installing it from pip to make it more user friendly
openpyxl_module = "Modules"
module_path = os.path.join(os.path.dirname(__file__), openpyxl_module)
sys.path.append(module_path)

import openpyxl
import arcpy

from datetime import datetime

#Location of user selected Excel file
excelfile_loc = ""

#Imports the specified excel sheet into the specified file geodatabase
def importSheet(excel_loc, output_loc, sheet_name):
    arcpy.management.Delete(output_loc)
    arcpy.ExcelToTable_conversion(excel_loc, output_loc, sheet_name)

#Runs when user clicks "Update" button.
    #arcpy.env.workspace is updated to user selected gdb
def updateFunction():
    #importSheet(temp_excel_loc, "Import", "Sheet1")
    #arcpy.management.JoinField("FeatureClass", "PrimaryKey", "Import", "PrimaryKey", ["Field1", "Field2"])
    #arcpy.management.Delete("Import")
    pass
    
##################################################################################################
#
#              GUI Code
#
##################################################################################################


from Tkinter import *
import tkFileDialog as filedialog
import tkMessageBox as messagebox

root = Tk()
root.title("Update GDB from Excel")
scrollbar = Scrollbar(orient="horizontal")
fileentry = Entry(root, xscrollcommand=scrollbar.set, width=100)
fileentry.focus()
fileentry.grid(row=0, column=1, sticky="w")
scrollbar.grid(row=1, column=1, ipadx=275, sticky="w")
scrollbar.config(command=fileentry.xview)
fileentry.config()

def selectexcel():
    x = filedialog.askopenfilename(title="Select File")
    datafile = x
    fileentry.delete('0', END)
    fileentry.insert(END, datafile)

dataselect = Button(root, text="Select Excel Sheet", padx=50, command=selectexcel)
dataselect.grid(row=0, column=0)


scrollbar2 = Scrollbar(orient="horizontal")
foldentry = Entry(root, xscrollcommand=scrollbar2.set, width=100)
foldentry.focus()
foldentry.grid(row=2, column=1, sticky="w")
scrollbar2.grid(row=3, column=1, ipadx=275, sticky="w")
scrollbar2.config(command=fileentry.xview)
foldentry.config()

def selectgdb():
    x = filedialog.askdirectory(title="Select GDB")
    datafile = x
    foldentry.delete('0', END)
    foldentry.insert(END, datafile)

dataselect = Button(root, text="Select GDB folder", padx=50, command=selectgdb)
dataselect.grid(row=2, column=0)

def mainrun():
    global excelfile_loc
    global output_loc
    arcpy.env.workspace = foldentry.get()
    excelfile_loc = fileentry.get()
    updateFunction()
    messagebox.showinfo("Completed", "GDB has been updated!")
    

runbutton = Button(root, text="Update", command=mainrun)
runbutton.grid(row=4, column=0, ipadx=50, sticky="w")

root.mainloop()

