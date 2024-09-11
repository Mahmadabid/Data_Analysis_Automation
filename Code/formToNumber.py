import re
import pandas as pd
from tkinter import BooleanVar, Frame, StringVar, Tk, Button, Label, ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror, showinfo

# Function to open a file dialog for selecting the input file
def select_input_file():
    file_path = askopenfilename(
        title="Select the Input Excel File",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    return file_path

def select_output_file(default_name="FinalOutput.xlsx"):
    file_path = asksaveasfilename(
        title="Save the Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,  # Set default file name
    )
    return (
        file_path if file_path else default_name
    )  # Return the default if no file selected

def select_mapping_file():
    file_path = askopenfilename(
        title="Select the Mapping Excel File",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    return file_path

def load_mapping(file_path):
    """Load mapping from the Excel file into a dictionary."""
    mapping_df = pd.read_excel(file_path, header=None)
    if mapping_df.shape[1] != 2:
        raise ValueError("Mapping file must have exactly two columns.")
    return dict(zip(mapping_df[0], mapping_df[1]))

def process_file(
    input_file,
    output_file,
    mapping_file,
    cols_to_skip,
    delete_first_column,
    sheet_name
):
    try:
        # Load the mapping data from the Excel file
        mapping = load_mapping(mapping_file)
        
        # Load the Excel file with the specified sheet
        responses_df = pd.read_excel(input_file, sheet_name=sheet_name)

        # Strip whitespace from all cells in the DataFrame
        responses_df = responses_df.map(lambda x: x.strip() if isinstance(x, str) else x)

        def replace_responses(cell):
            """Replace text in a cell based on the mapping dictionary."""
            if isinstance(cell, str):
                # Directly match the entire cell value against the keys in the mapping dictionary
                cell = mapping.get(cell, cell)  # Replace cell value with the mapped value if it exists, otherwise keep the original cell value
            return cell

        # Copy the DataFrame and apply the replacement function to the desired rows and columns
        standardized_df = responses_df.copy()

        # Ensure cols_to_skip is a list of integers
        cols_to_skip = [
            int(col.strip()) - 1 for col in cols_to_skip if col.strip().isdigit()
        ]

        # Identify the columns to skip
        cols_to_process = [
            i for i in range(standardized_df.shape[1]) if i not in cols_to_skip
        ]

        # Apply the replacement function, skipping the specified columns
        standardized_df.iloc[:, cols_to_process] = standardized_df.iloc[
            :, cols_to_process
        ].map(replace_responses)

        # Remove the first column
        if delete_first_column:
            standardized_df = standardized_df.iloc[:, 1:]

        # Try to save the standardized DataFrame to a new Excel file
        try:
            standardized_df.to_excel(output_file, index=False)
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
    mapping_file = None
    sheet_name = StringVar()
    cols_to_skip = StringVar()
    delete_first_column = BooleanVar()

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

    def load_mapping_file():
        nonlocal mapping_file
        mapping_file = select_mapping_file()
        if mapping_file:
            mapping_label.config(text=f"Mapping File: {mapping_file}")

    def process_file_button():
        if not input_file:
            showinfo("Error", "No input file selected.")
            return
        if not output_file:
            showinfo("Error", "No output file selected.")
            return
        if not mapping_file:
            showinfo("Error", "No mapping file selected.")
            return
        if not sheet_name.get():
            showinfo("Error", "No sheet name selected.")
            return
        # Convert the column names from the entry widget to a list
        cols_to_skip_list = [
            col.strip() for col in cols_to_skip.get().split(",") if col.strip()
        ]
        process_file(
            input_file, output_file, mapping_file, cols_to_skip_list, delete_first_column.get(), sheet_name.get()
        )

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

    Button(
        main_frame,
        text="Select Mapping File",
        command=load_mapping_file,
        background="#444d5c",
        foreground="white",
        border=3,
    ).pack(pady=5)
    mapping_label = Label(main_frame, text="Mapping File: None")
    mapping_label.pack(pady=5)

    Label(main_frame, text="Sheet Name:").pack(pady=5)
    sheet_dropdown = ttk.Combobox(main_frame, textvariable=sheet_name)
    sheet_dropdown.pack(pady=5)

    Label(main_frame, text="Columns to Skip (comma-separated):").pack(pady=5)
    cols_to_skip_entry = ttk.Entry(main_frame, textvariable=cols_to_skip)
    cols_to_skip_entry.pack(pady=5)

    delete_first_column_check = ttk.Checkbutton(
        main_frame, text="Delete First Column", variable=delete_first_column
    )
    delete_first_column_check.pack(pady=5)

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
