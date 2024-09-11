import csv
from openpyxl import Workbook
from tkinter import Frame, Tk, Button, Label
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror, showinfo

# Function to open a file dialog for selecting the input CSV file
def select_input_file():
    """Open file dialog to select input CSV file."""
    file_path = askopenfilename(
        title="Select the Input CSV File",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    return file_path if file_path else None

def select_output_file(default_name="output.xlsx"):
    """Open file dialog to select output Excel file with a default name."""
    file_path = asksaveasfilename(
        title="Save the Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name  # Set default file name
    )
    return file_path if file_path else default_name  # Return the default if no file selected

def convert_csv_to_excel(input_file, output_file):
    """Convert CSV file to Excel file using openpyxl, replacing 'None' with 'none'."""
    try:
        # Create a new Excel workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active

        # Read the CSV file
        with open(input_file, 'r', newline='', encoding='utf-8') as csvfile:
            csvreader = csv.reader(csvfile)

            for row_index, row in enumerate(csvreader):
                # Replace 'None' with 'none' in each cell
                row = [value.replace("None", "none") for value in row]
                
                for col_index, value in enumerate(row):
                    ws.cell(row=row_index + 1, column=col_index + 1, value=value)

        # Save the workbook to the output file
        wb.save(output_file)
        showinfo("Success", f"File converted and saved to {output_file}")

    except PermissionError:
        showerror(
            "Error",
            "The output file is currently open. Please close it before saving.",
        )
    except Exception as e:
        showerror("Error", f"An error occurred: {e}")

# GUI Application
def main():
    """Main function to run the GUI application."""
    root = Tk()
    root.title("CSV to Excel Converter")

    input_file = None
    output_file = None

    def load_input_file():
        nonlocal input_file
        input_file = select_input_file()
        if input_file:
            input_label.config(text=f"Input File: {input_file}")

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
        convert_csv_to_excel(input_file, output_file)

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
        text="Convert File",
        command=process_file_button,
        background="#444d5c",
        foreground="white",
        border=3,
    ).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
