# Manual Weigh Logger - To log manual weights and export to CSV files
# Copyright (C) 2024  TLundComputing
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

# Imports
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import tkinter.messagebox as box
import datetime
import csv
import os
import sys

def resource_path(relative_path):
    absolute_path = os.path.abspath(__file__)
    root_path = os.path.dirname(absolute_path)
    base_path = getattr(sys, '_MEIPASS', root_path)
    return os.path.join(base_path, relative_path)

# Window
window = Tk()
icon = PhotoImage(file=resource_path('icon.png'))
window.iconphoto(False, icon)

# Widgets
titleLbl = Label(window, font=("Arial", 32))
addWeightBtn = Button(window, font=("Arial", 12))
exportWeightsBtn = Button(window, font=("Arial", 12))
dayLbl = Label(window, font=("Arial", 12))
sessionLbl = Label(window, font=("Arial", 12))
firstnameLbl = Label(window, font=("Arial", 12))
surnameLbl = Label(window, font=("Arial", 12))
stonesLbl = Label(window, font=("Arial", 12))
poundsLbl = Label(window, font=("Arial", 12))
stonesEntry = Entry(window, width=2, font=("Arial", 12))
poundsEntry = Entry(window, width=4, font=("Arial", 12))
firstnameEntry = Entry(window, width=20, font=("Arial", 12))
surnameEntry = Entry(window, width=20, font=("Arial", 12))
dayVar = StringVar()
dayCombo = ttk.Combobox(window, font=("Arial", 12))
sessionVar = StringVar()
sessionCombo = ttk.Combobox(window, font=("Arial", 12))
memberCountLbl = Label(window, font=("Arial", 12))
memberNumberLbl = Label(window, relief="groove", width=4, font=("Arial", 12))
visitorVar = StringVar()
visitorCheck = Checkbutton(window, text="Visitor?", variable=visitorVar, onvalue="Yes", offvalue="No", font=("Arial", 12))

# Geometry
titleLbl.grid(row=1, column=1, columnspan=4, padx=10, pady=10)
dayLbl.grid(row=2, column=1, columnspan=2)
dayCombo.grid(row=2, column=3, columnspan=2)
sessionLbl.grid(row=3, column=1, columnspan=2)
sessionCombo.grid(row=3, column=3, columnspan=2)
firstnameLbl.grid(row=4, column=1, padx=10)
firstnameEntry.grid(row=4, column=2)
surnameLbl.grid(row=4, column=3)
surnameEntry.grid(row=4, column=4, padx=10)
stonesLbl.grid(row=5, column=1)
stonesEntry.grid(row=5, column=2)
poundsLbl.grid(row=5, column=3)
poundsEntry.grid(row=5, column=4)
visitorCheck.grid(row=6, column=1, columnspan=4)
addWeightBtn.grid(row=7, column=1, columnspan=4)
exportWeightsBtn.grid(row=8, column=1, columnspan=2, pady=10)
memberCountLbl.grid(row=8, column=3)
memberNumberLbl.grid(row=8, column=4)

# Static Properties
window.title("Manual Weigh")
window.resizable(0, 0)
dayCombo['values'] = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
sessionCombo['values'] = ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
titleLbl.configure(text="Manual Weigh")
dayLbl.configure(text="Session Day:")
sessionLbl.configure(text="Session Number:")
stonesLbl.configure(text="Stones:")
poundsLbl.configure(text="Pounds:")
addWeightBtn.configure(text="Add Weight")
exportWeightsBtn.configure(text="Export")
memberCountLbl.configure(text="Members Through:")
firstnameLbl.configure(text="Firstname:")
surnameLbl.configure(text="Surname:")

# Initial Properties
membersWeightsToExport = []
day = datetime.datetime.now()
unsavedChanges = False
dayIndex = 0
for index in range(0, len(dayCombo['values'])):
    if day.strftime("%A") == dayCombo['values'][index]:
        dayIndex = index
dayCombo.current(dayIndex)
sessionCombo.current(0)
memberNumberLbl.configure(text=str(len(membersWeightsToExport)))
exportWeightsBtn.configure(state = DISABLED)
visitorCheck.deselect()

# Dynamic Properties
def add_weight():
    global unsavedChanges
    fname = firstnameEntry.get()
    sname = surnameEntry.get()
    st = stonesEntry.get()
    lbs = poundsEntry.get()
    if len(fname.strip()) == 0 or len(sname.strip()) == 0 or len(st.strip()) == 0 or len(lbs.strip()) == 0:
        box.showerror("Missing Information", "Some entry fields are empty. Resolve before continuing")
        return
    try:
        int(st)
        float(lbs)
    except:
        box.showerror("Wrong Format", "Stones or pounds are not entered as valid numbers. Resove before continuing")
        return
    membersWeightsToExport.append([dayCombo.get(), sessionCombo.get(), fname, sname, st, lbs, visitorVar.get()])
    memberNumberLbl.configure(text=str(len(membersWeightsToExport)))
    firstnameEntry.delete(0, len(fname))
    surnameEntry.delete(0, len(sname))
    stonesEntry.delete(0, len(st))
    poundsEntry.delete(0, len(lbs))
    visitorCheck.deselect()
    exportWeightsBtn.configure(state = NORMAL)
    unsavedChanges = True

def export_weights():
    global unsavedChanges
    writeFile = box.askyesno("Confirm?", f"Are you sure you want to commit {len(membersWeightsToExport)} members?")
    if writeFile == 1:
        fields = ["Day", "Session", "Firstname", "Surname", "Stones", "Pounds", "visitor?"]
        path = filedialog.asksaveasfilename(initialfile=dayCombo.get() + " weights.csv", defaultextension=".csv", filetypes=[("Comma Seperated Files", "*.csv")])
        with open(path, "w") as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(fields)
            csvwriter.writerows(membersWeightsToExport)
        unsavedChanges = False

addWeightBtn.configure(command=add_weight)
exportWeightsBtn.configure(command=export_weights)

# On close
def on_close():
    if unsavedChanges:
        e = box.askyesno("Confirm Exit", "You have unexported changes, do you want to exit?")
        if e == 1:
            window.destroy()
    else:
        window.destroy()

# Sustain Window
window.protocol("WM_DELETE_WINDOW", on_close)
window.mainloop()