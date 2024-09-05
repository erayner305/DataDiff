import pandas
import tkinter
from tkinter import filedialog
from csv_diff import load_csv, compare, human_text

# Function to format date columns
def format_date_column(df, date_column_name):
    df[date_column_name] = pandas.to_datetime(df[date_column_name], errors='coerce').dt.strftime('%m/%d/%Y')
    return df

# Global variables to hold the file paths
file1_path = ""
file2_path = ""
OUTPUT_PATH = "diff_output.txt"

# Function to trigger the file selection dialog
def file_select(file_num, label):
    global file1_path, file2_path, confirmationLabel, errorLabel
    errorLabel.pack_forget()
    confirmationLabel.pack_forget()
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path != "":
        label.config(text=f"File {file_num} Selected: {file_path}")
        if file_num == 1:
            file1_path = file_path  # Store file1 path
        elif file_num == 2:
            file2_path = file_path  # Store file2 path
    print(f"File {file_num} selected: {file_path}")  # Print the selected file path for debugging

# Function to perform the diff check
def check_Diff(file1, file2):
    global confirmationLabel, errorLabel
    if file1 != "" or file2 != "":
        # Load the CSV files into dataframes
        dataframe1 = pandas.read_excel(file1)
        dataframe2 = pandas.read_excel(file2)

        # Format date columns and create a "key" column
        dataframe1 = format_date_column(dataframe1, 'Date')
        dataframe1['key'] = dataframe1['Well'].astype(str) + '-' + dataframe1['Constituent'].astype(str) + '-' + dataframe1['Date'].astype(str)
        dataframe1.to_csv("file1.csv", index=False)

        dataframe2 = format_date_column(dataframe2, 'Date')
        dataframe2['key'] = dataframe2['Well'].astype(str) + '-' + dataframe2['Constituent'].astype(str) + '-' + dataframe2['Date'].astype(str)
        dataframe2.to_csv("file2.csv", index=False)

        # Compare the two CSV files and generate the diff
        diff = compare(load_csv(open("file1.csv"), "key"), load_csv(open("file2.csv"), "key"))
        global OUTPUT_PATH
        # Save the diff output to a file
        with open(OUTPUT_PATH, 'w') as output_file:
            output_file.write(human_text(diff, show_unchanged=True))

        confirmationLabel.pack()
    else:
        errorLabel.pack()


# Create GUI
root = tkinter.Tk(className="Diff Checker")

# Button to select file 1
importLabel1 = tkinter.Label(root, text=f"No File Selected")
importLabel1.pack(pady=2)
importButton1 = tkinter.Button(root, text="Import File 1", width=15, height=2, command=lambda: file_select(1, importLabel1))
importButton1.pack(pady=2)

# Button to select file 2
importLabel2 = tkinter.Label(root, text=f"No File Selected")
importLabel2.pack(pady=2)
importButton2 = tkinter.Button(root, text="Import File 2", width=15, height=2, command=lambda: file_select(2, importLabel2))
importButton2.pack(pady=2)

# Button to run the diff checker
confirmationLabel = tkinter.Label(root, text=f"Output saved to {OUTPUT_PATH}")
errorLabel = tkinter.Label(root, text="Please select a file for each input")
runButton = tkinter.Button(root, text="Run Diff Checker", width=20, height=2, command=lambda: check_Diff(file1_path, file2_path))
runButton.pack(pady=20, padx=150)

# Start the GUI event loop
root.mainloop()
