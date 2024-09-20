import ast
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

def select_output_file(default_name="Output.sps"):
    file_path = asksaveasfilename(
        title="Save the Output Spss Syntax",
        defaultextension=".sps",
        filetypes=[("Spss Syntax", "*.sps"), ("All files", "*.*")],
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
    delete_first_column,
    cols_to_convert,
    sheet_name
):
    try:
        # Load the mapping data from the Excel file
        mapping = load_mapping(mapping_file)
        
        # Load the Excel file with the specified sheet
        responses_df = pd.read_excel(input_file, sheet_name=sheet_name)

        # Apply strip() only to string-type columns
        responses_df = responses_df.map(lambda x: x.strip() if isinstance(x, str) else x)

        # Copy the DataFrame and apply the replacement function to the desired rows and columns
        standardized_df = responses_df.copy()

        # Remove the first column if specified
        if delete_first_column:
            standardized_df = standardized_df.iloc[:, 1:]

             # Adjust column indices for conversion
            cols_to_convert = [int(col) - 2 for col in cols_to_convert if col.isdigit()]
        else:
            # Adjust column indices for conversion
            cols_to_convert = [int(col) - 1 for col in cols_to_convert if col.isdigit()]
            

        # Get the headers from the filtered DataFrame
        original_headers = standardized_df.columns.tolist()

        # Number of columns in the filtered DataFrame
        num_columns = len(original_headers)

        # Generate new headers
        new_headers = [f'Q{i+1}' for i in range(num_columns)]

        # Create dictionary, numeric, number, and other columns
        dict_column = []
        numeric_column = []
        number_column = []
        zero_column = [0] * num_columns  # Column with 0 for all rows
        eight_column = [8] * num_columns  # Column 8 with number 8 for all rows
        empty_column_7 = [None] * num_columns  # Column 7 (1-based) empty
        left_column = ['Left'] * num_columns  # Column 9 (1-based)
        nominal_column = ['Nominal'] * num_columns  # Column 10 (1-based)
        input_column = ['Input'] * num_columns  # New column with "Input"

        # Initialize max_mapping_number based on mapping file content
        max_mapping_number = max(mapping.values())

        for idx, col in enumerate(standardized_df.columns):
            if idx in cols_to_convert:
                dict_column.append(None)  # Set to None if excluded
                numeric_column.append("String")  # Set "String" for excluded columns
                number_column.append(256)  # Assign 256 for "String"
            else:
                col_dict = {}
                for value in standardized_df[col].dropna().unique():
                    if value in mapping:
                        col_dict[value] = int(mapping[value])  # Convert to int
                    else:
                        max_mapping_number += 1
                        col_dict[value] = int(max_mapping_number)  # Convert to int
                
                dict_column.append(str(col_dict))  # Convert dict to string format
                numeric_column.append("Numeric")  # Set "Numeric" for included columns
                number_column.append(8)  # Assign 8 for "Numeric"

        # Create DataFrame with the required columns and rearrange the order
        combined_df = pd.DataFrame({
            'Name': new_headers,
            'Type': numeric_column,
            'Width': number_column,
            'Decimals': zero_column,
            'Label': original_headers,
            'Values': dict_column,
            'Missing': empty_column_7,
            'Columns': eight_column,
            'Align': left_column,
            'Measure': nominal_column,
            'Role': input_column,
        })

        # Generate SPSS syntax
        spss_syntax = []

        # Start with DATA LIST command (adapt based on your actual data structure)
        data_list_command = "DATA LIST FREE / " + " ".join(
            f"{name} ({'A' if type_ == 'String' else 'F'}{width})" 
            for name, type_, width in zip(combined_df['Name'], combined_df['Type'], combined_df['Width'])
        )
        spss_syntax.append(data_list_command + ".")

        # Create VARIABLE LABELS command with each label on a separate line
        for name, label in zip(combined_df['Name'], combined_df['Label']):
            # Replace newline characters with a space
            clean_label = label.replace('\n', ' ').replace('\r', '')
            variable_labels_command = f"VARIABLE LABELS {name} \"{clean_label.replace('\'', '\'\'')}\"."
            spss_syntax.append(variable_labels_command)

        # Add VARIABLE LEVEL command
        variable_level_command = "VARIABLE LEVEL " + " ".join(
            f"{name} (NOMINAL)"
            for name in combined_df['Name']
        )
        spss_syntax.append(variable_level_command + ".")

        # Add ALIGNMENT command for all variables
        alignment_command = "VARIABLE ALIGNMENT " + " ".join(
            f"{name} (LEFT)"
            for name in combined_df['Name']
        )
        spss_syntax.append(alignment_command + ".")

        # Add VALUE LABELS command with error handling
        for name, values in zip(combined_df['Name'], combined_df['Values']):
            if values and values != 'None':  # Ensure the variable name exists
                value_labels = []
                try:
                    value_dict = ast.literal_eval(values)  # Try converting string to dictionary
                    for label, value in value_dict.items():  # Swap key and value
                        value_labels.append(f"{value} '{label}'")
                    value_labels_command = f"VALUE LABELS {name}\n" + "\n".join(f"    {line}" for line in value_labels) + "."
                    spss_syntax.append(value_labels_command)
                except (ValueError, SyntaxError) as e:
                    # Skip malformed strings or handle them
                    showerror("Error", f"Skipping value labels for column {name}: because its already in numbers. It cant be converted.")

        try:
            # Save SPSS syntax to a file with .sps extension
            with open(output_file, 'w', encoding='utf-8') as file:
                file.write('\n'.join(spss_syntax))

            showinfo("Success", f"SPSS syntax file saved as {output_file}")
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
    root.title("SPSS Variable View Generator")

    input_file = None
    output_file = None
    mapping_file = None
    sheet_name = StringVar()
    delete_first_column = BooleanVar()
    cols_to_convert = StringVar()

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
        output_file = asksaveasfilename(
            title="Save the Output File",
            defaultextension=".sps",
            filetypes=[("spss syntax", "*.sps"), ("All files", "*.*")],
        )
        if output_file:
            output_label.config(text=f"Output File: {output_file}")

    def load_mapping_file():
        nonlocal mapping_file
        mapping_file = askopenfilename(
            title="Select the Mapping Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
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
        cols_to_convert_list = [
            col.strip() for col in cols_to_convert.get().split(",") if col.strip()
        ]
        process_file(
            input_file, output_file, mapping_file, delete_first_column.get(), cols_to_convert_list, sheet_name.get()
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

    delete_first_column_check = ttk.Checkbutton(
        main_frame, text="Delete First Column", variable=delete_first_column
    )
    delete_first_column_check.pack(pady=5)

    Label(main_frame, text="Columns to Convert to String (comma-separated):").pack(pady=5)
    cols_to_convert_entry = ttk.Entry(main_frame, textvariable=cols_to_convert)
    cols_to_convert_entry.pack(pady=5)

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
