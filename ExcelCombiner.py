import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog


# Define a function that will be called when the user clicks the "Extract Data"
def extract_data():
    # Prompt the user to select a folder where the Excel files are located
    directory = filedialog.askdirectory()

    # Get the rows and columns to extract from the user
    rows_to_extract_str = rows_entry.get()
    if ":" in rows_to_extract_str:
        # Input format: "1:3"
        start, end = rows_to_extract_str.split(":")
        rows_to_extract = list(range(int(start), int(end)))
    else:
        # Input format: "1,2,3"
        rows_to_extract = list(map(int, rows_to_extract_str.split(',')))

    cols_to_extract_str = cols_entry.get()
    # Input format: "1,2,3"
    cols_to_extract = list(map(int, cols_to_extract_str.split(',')))
    cols_to_extract = [col - 1 for col in cols_to_extract]  # Shift the index by 1

    # Create an empty dataframe to hold the extracted data
    data = pd.DataFrame()

    # Loop through each file in the selected folder
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):  # Only consider Excel files
            filepath = os.path.join(directory, filename)
            # Read the Excel file into a pandas dataframe
            df = pd.read_excel(filepath)
            # Extract the selected rows and columns from the dataframe
            subset = df.iloc[rows_to_extract, cols_to_extract]
            # Append the subset to the data dataframe
            data = data.append(subset)

    # Prompt the user to select a filename to save the extracted data to
    output_path = filedialog.asksaveasfilename(defaultextension='.xlsx')

    # Save the data dataframe to the selected filename as an Excel file
    data.to_excel(output_path, index=False)


# Create a GUI window
root = tk.Tk()
root.title("Excel Ext.")
root.geometry("300x200")

# Create labels and entry boxes for the rows and columns to extract
rows_label = tk.Label(root, text="Rows (Use commas or colons):")
rows_entry = tk.Entry(root)
cols_label = tk.Label(root, text="Columns (Use commas):")
cols_entry = tk.Entry(root)

# Create a button to extract the data as an Excel file
extract_button = tk.Button(root, text="Extract Data (Excel)", command=extract_data)

# Add the labels, entry boxes, and buttons to the GUI window
rows_label.pack()
rows_entry.pack()
cols_label.pack()
cols_entry.pack()
extract_button.pack()

# Run the GUI loop
root.mainloop()

# Made by: MrYety
# Version 1.0
