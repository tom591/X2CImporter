import tkinter as tk
from tkinter import filedialog
from pandas import read_excel, concat
from datetime import date
from tkinter import *

# Function to select an XLSX file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    if file_path:
        selected_file_label.config(text="Selected file: " + file_path)

# Function to perform the conversion and save the output CSV
def perform_action_and_save():
    # Check if a file has been selected
    if selected_file_label.cget("text") == "":
        # Show error if no file selected
        tk.messagebox.showerror("Error", "No file has been selected.")
        return

    file_path = selected_file_label.cget("text").split(": ")[1]

    # Load all sheets using specific columns (third column is Integer)
    df1 = read_excel(file_path, sheet_name=0, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8],
                     dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
    df2 = read_excel(file_path, sheet_name=1, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8],
                     dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
    df3 = read_excel(file_path, sheet_name=2, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8],
                     dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
    df4 = read_excel(file_path, sheet_name=3, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8],
                     dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)

    # Insert row number (ID) into the first sheet
    df1.insert(0, 'Row Number', range(1, len(df1) + 1))

    # Continue row numbering for the second sheet
    next_number1 = len(df1) + 1
    df2.insert(0, 'Row Number', range(next_number1, next_number1 + len(df2)))

    # Continue row numbering for the third sheet
    next_number2 = len(df1) + len(df2) + 1
    df3.insert(0, 'Row Number', range(next_number2, next_number2 + len(df3)))

    # Continue row numbering for the fourth sheet
    next_number3 = len(df1) + len(df2) + len(df3) + 1
    df4.insert(0, 'Row Number', range(next_number3, next_number3 + len(df4)))

    # Concatenate all sheets vertically into one dataframe
    result_df = concat([df1, df2, df3, df4], axis=0)

    # Add today's date to the output filename (format: YYYY-MM-DD)
    creation_date = date.today()
    creation_date_str = str(creation_date)

    # Ask user to choose location to save the CSV file
    save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             initialfile="x2c_import_" + creation_date_str + ".csv")
    if save_path:
        result_df.to_csv(save_path, sep=';', index=False, header=False, encoding='utf-8')

# "About" dialog window
def about():
    tk.messagebox.showinfo(
        title=None,
        message='The "X2CImporter" program selects data from an XLSX spreadsheet and converts it into a CSV file suitable for MySQL database import.\n\nAfter loading the data using the first button, the selected file name will appear below. Then simply click the second button to generate the CSV and save it.\n\n© Tom Salaj 2024'
    )

# Create main window and menu bar
root = tk.Tk()
menubar = Menu(root)
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="About", command=about)
menubar.add_cascade(label="Help", menu=helpmenu)

# Set window title and position
root.title("X2CImporter v1.1")
root.geometry('+%d+%d' % (350, 300))

# Button to select file
img1 = PhotoImage(file="_internal/icons/xlsx_icon.png")
select_file_button = tk.Button(
    root, height=70, width=250, image=img1,
    text="1. Select file to convert", compound="left", command=select_file)
select_file_button.pack(pady=20, padx=20)

# Label showing the selected file name
selected_file_label = tk.Label(root, text="", wraplength=300)
selected_file_label.pack()

# Button to convert and save as CSV
img2 = PhotoImage(file="_internal/icons/csv_icon.png")
perform_action_button = tk.Button(
    root, height=70, width=250, image=img2,
    text="2. Convert data to CSV and save", compound="left", command=perform_action_and_save)
perform_action_button.pack(pady=10, padx=20)

# Footer label
labelfooter = tk.Label(root, fg="#666666", text="© Tom Salaj 2024", anchor="center", justify="right")
labelfooter.pack(pady=10, padx=10)

# Load menu
root.config(menu=menubar)
root.mainloop()