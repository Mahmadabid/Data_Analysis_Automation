import re
import pandas as pd
from tkinter import Frame, StringVar, Tk, Button, Label, ttk, BooleanVar, Checkbutton
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror, showinfo


# Function to open a file dialog for selecting the input file
def select_input_file():
    file_path = askopenfilename(
        title="Select the Input Excel File",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    return file_path

def select_output_file(default_name="output.xlsx"):
    file_path = asksaveasfilename(
        title="Save the Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,  # Set default file name
    )
    return (
        file_path if file_path else default_name
    )  # Return the default if no file selected


def process_file(input_file, output_file, keep_original_columns, sheet_name):
    try:
        # Load the Excel files
        df = pd.read_excel(input_file, sheet_name=sheet_name)
 
        # Function to split by commas outside of brackets
        def split_outside_brackets(text):
            if isinstance(text, str):
                parts = re.split(r";(?![^\(\[]*[\)\]])", text)
                return [part.strip() for part in parts]
            return []

        # Identify columns with multiple responses
        columns_with_multiple_responses = [
            col
            for col in df.columns
            if df[col]
            .apply(lambda x: isinstance(x, str) and len(split_outside_brackets(x)) > 1)
            .any()
        ]

        # Store the original columns to drop them later
        original_columns = []

        # Expand each column with multiple responses
        for col in columns_with_multiple_responses:
            # Get unique responses across all rows, including None values
            unique_responses = []
            df[col].dropna().apply(
                lambda x: [
                    unique_responses.append(resp)
                    for resp in split_outside_brackets(x)
                    if resp not in unique_responses
                ]
            )

            # Create new columns for each unique response
            new_columns = {}
            for response in unique_responses:
                new_col_name = response.strip()

                # Ensure the new column name is unique
                if new_col_name in df.columns:
                    suffix = 1
                    while f"{new_col_name}_{suffix}" in df.columns:
                        suffix += 1
                    new_col_name = f"{new_col_name}_{suffix}"

                # Create the new column data
                new_columns[new_col_name] = df[col].apply(
                    lambda x: (
                        "Yes"
                        if isinstance(x, str) and response in split_outside_brackets(x)
                        else "No"
                    )
                )

            # Insert new columns right after the original column
            col_index = df.columns.get_loc(col)
            for new_col_name, new_col_data in new_columns.items():
                col_index += 1
                df.insert(col_index, new_col_name, new_col_data)

            # Add the original column to the list to be dropped
            original_columns.append(col)

        # Drop the original columns if the option is selected
        if not keep_original_columns:
            df.drop(columns=original_columns, inplace=True)

        # Try to save the standardized DataFrame to a new Excel file
        try:
            df.to_excel(output_file, index=False)
            showinfo("Success", f"Updated DataFrame saved to {output_file}")
        except PermissionError:
            showerror(
                "Error",
                "The output file is currently open. Please close it before saving.",
            )

    except Exception as e:
        showerror("Error", f"An error occurred: {e}")


# GUI Application
def main():
    root = Tk()
    root.title("Excel Response Standardizer")

    input_file = None
    output_file = None
    sheet_name = StringVar()
    keep_columns_var = BooleanVar()

    def load_input_file():
        nonlocal input_file
        input_file = select_input_file()
        if input_file:
            input_label.config(text=f"Input File: {input_file}")
            # Populate sheet dropdown
            sheet_names = pd.ExcelFile(input_file).sheet_names
            sheet_dropdown["values"] = sheet_names
            if sheet_names:
                sheet_dropdown.set(sheet_names[0])  # Set default to the first sheet

    def load_output_file():
        nonlocal output_file
        output_file = select_output_file()
        if output_file:
            output_label.config(text=f"Output File: {output_file}")

    def process_file_button():
        if not input_file:
            showinfo("Error", "No input file selected.")
            return
        if not output_file:
            showinfo("Error", "No output file selected.")
            return
        process_file(input_file, output_file, keep_columns_var.get(), sheet_name.get())

    # Create a frame to contain all widgets and add padding to the frame
    main_frame = Frame(root, padx=20, pady=20)
    main_frame.pack(padx=10, pady=10)  # Add padding around the frame itself

    # Create GUI elements within the frame
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

    # Checkbox for dropping original columns
    keep_columns_checkbox = Checkbutton(
        main_frame, text="Keep Original Columns", variable=keep_columns_var
    )
    keep_columns_checkbox.pack(pady=10)

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
