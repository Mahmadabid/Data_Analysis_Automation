import re
import numpy as np
import pandas as pd
from tkinter import Frame, StringVar, Tk, Button, Label, ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror, showinfo

def select_input_file():
    return askopenfilename(
        title="Select the Input Excel File",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )

def select_output_file(default_name="mapping.xlsx"):
    file_path = asksaveasfilename(
        title="Save the Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,
    )
    return file_path if file_path else default_name

def process_file(input_file, output_file, cols_to_skip, sheet_name):
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        try:
            cols_to_skip = [int(col.strip()) - 1 for col in cols_to_skip if col.strip()]
        except ValueError:
            showerror("Error", "Columns to skip must be integers.")
            return

        # Function to check if a column contains only numbers (integers or floats)
        def is_numeric_column(column):
            try:
                pd.to_numeric(column, errors='raise')
                return True
            except ValueError:
                return False

        # Identify columns that contain only numbers
        numeric_columns = [i for i, col in enumerate(df.columns) if is_numeric_column(df.iloc[:, i])]
        
        # Update cols_to_skip to include numeric columns
        cols_to_skip.extend(numeric_columns)
        
        # Ensure we do not have duplicates and sort
        cols_to_skip = list(sorted(set(cols_to_skip)))
        
        max_index = len(df.columns) - 1
        cols_to_skip = [i for i in cols_to_skip if 0 <= i <= max_index]

        df = df.drop(df.columns[cols_to_skip], axis=1, errors='ignore')

        def create_mapping(column):
            stripped_responses = column.astype(str).str.strip()

            # Handle 'nan' as actual NaN values
            stripped_responses = stripped_responses.replace('nan', np.nan)
            
            # Get unique non-null responses
            unique_responses = pd.unique(stripped_responses.dropna())

            # Start with predefined mappings
            response_mapping = {'Yes': 1, 'No': 0}

            # Enumerate starting from 1 for new responses
            current_index = 1
            for response in unique_responses:
                # Skip if response is empty or already in the mapping
                if response not in response_mapping and response != '':
                    response_mapping[response] = current_index
                    current_index += 1

            return response_mapping
        
        def convert_column(column, mapping):
            stripped_column = column.astype(str).str.strip()
            return stripped_column.map(lambda x: mapping.get(x, x) if x != '' else x)

        all_mappings = {}

        for column in df.columns:
            # Create mapping for each column
            mapping = create_mapping(df[column])
            df[column] = convert_column(df[column], mapping)
            
            # Add column-specific mappings to all_mappings
            all_mappings.update(mapping)

        # Create DataFrame for mappings
        mapping_df = pd.DataFrame(list(all_mappings.items()), columns=["Value", "Number"])

        try:
            mapping_df.to_excel(output_file, header=False, index=False)
            showinfo("Success", f"Mappings saved to {output_file}")
        except PermissionError:
            showerror("Error", "The output file is currently open. Please close it before saving.")

    except Exception as e:
        showerror("Error", f"An error occurred: {e}")

def main():
    root = Tk()
    root.title("Excel Response Standardizer")

    input_file = None
    output_file = None
    sheet_name = StringVar()
    cols_to_skip = StringVar()

    def load_input_file():
        nonlocal input_file
        input_file = select_input_file()
        if input_file:
            input_label.config(text=f"Input File: {input_file}")
            sheet_names = pd.ExcelFile(input_file).sheet_names
            sheet_dropdown["values"] = sheet_names
            if sheet_names:
                sheet_dropdown.set(sheet_names[0])

    def load_output_file():
        nonlocal output_file
        output_file = select_output_file()
        if output_file:
            output_label.config(text=f"Output File: {output_file}")

    def process_file_button():
        if not input_file:
            showerror("Error", "No input file selected.")
            return
        if not output_file:
            showerror("Error", "No output file selected.")
            return
        if not sheet_name.get():
            showerror("Error", "No sheet name selected.")
            return
        
        cols_to_skip_list = [col.strip() for col in cols_to_skip.get().split(",") if col.strip()]
        process_file(input_file, output_file, cols_to_skip_list, sheet_name.get())

    main_frame = Frame(root, padx=20, pady=20)
    main_frame.pack(padx=10, pady=10)

    Button(
        main_frame,
        text="Select Input File",
        command=load_input_file,
        background="#444d5c",
        foreground="white",
        border=3,
    ).pack(pady=5)
    input_label = Label(main_frame, text="Input File: None")
    input_label.pack(pady=5)

    Button(
        main_frame,
        text="Select Output File",
        command=load_output_file,
        background="#444d5c",
        foreground="white",
        border=3,
    ).pack(pady=5)
    output_label = Label(main_frame, text="Output File: None")
    output_label.pack(pady=5)

    Label(main_frame, text="Sheet Name:").pack(pady=5)
    sheet_dropdown = ttk.Combobox(main_frame, textvariable=sheet_name)
    sheet_dropdown.pack(pady=5)

    Label(main_frame, text="Columns to Skip (comma-separated):").pack(pady=5)
    cols_to_skip_entry = ttk.Entry(main_frame, textvariable=cols_to_skip)
    cols_to_skip_entry.pack(pady=5)

    Button(
        main_frame,
        text="Process File",
        command=process_file_button,
        background="#444d5c",
        foreground="white",
        border=3,
    ).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
