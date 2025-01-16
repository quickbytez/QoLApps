import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime
import os
import subprocess
import sys
import pandas as pd
import openpyxl

class LogSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Log Selector")

        self.selection = tk.StringVar(value="Single")
        self.fields = {}

        self.create_radio_buttons()
        self.create_common_fields()
        self.create_single_log_fields()
        self.create_multiple_log_fields()
        self.create_buttons()
        self.update_fields()

        self.output_folder = self.get_output_folder()
        self.excel_data = None  # Placeholder for Excel data

    def create_radio_buttons(self):
        single_radio = ttk.Radiobutton(self.root, text="Single logs", variable=self.selection, value="Single",
                                       command=self.update_fields)
        single_radio.grid(column=0, row=0, padx=10, pady=10, sticky='w')

        multiple_radio = ttk.Radiobutton(self.root, text="Multiple logs", variable=self.selection, value="Multiple",
                                         command=self.update_fields)
        multiple_radio.grid(column=1, row=0, padx=10, pady=10, sticky='w')

    def create_common_fields(self):
        # Add the new "Reference Number" field above "Main Issue"
        self.fields['Reference Number'] = self.create_label_entry("Reference Number", 1)

        # Adjusted rows for the fields below it
        self.fields['Main Issue'] = self.create_label_text("Main Issue", 2)
        self.fields['Action Taken'] = self.create_label_text("Action Taken", 3)
        self.fields['PC#'] = self.create_label_entry("PC#", 4)
        self.fields['Role'] = self.create_label_combo("Role",
                                                      [" ", "Host", "Connecting User"], 5)
        self.fields['Computer Name'] = self.create_label_entry("Computer Name", 6)
        self.fields['Computer Model/Type'] = self.create_label_entry("Computer Model/Type", 7)
        self.fields['Computer Serial Number'] = self.create_label_entry("Computer Serial Number", 8)
        self.fields['Computer OEM'] = self.create_label_combo("Computer OEM",
                                                              ["HP", "Dell", "Acer", "Lenovo", "Asus", "Apple",
                                                               "Microsoft", "Others"], 9)
        self.fields['Computer OS'] = self.create_label_combo("Computer OS",
                                                             [" ", "XP", "XPVM", "Win7", "Win8", "Win8.1", "Win10",
                                                              "Win11", "MacOS", "OS X", "Linux"], 10)

        # Bind the Computer Name entry to the autofill method
        self.fields['Computer Name'].bind("<FocusOut>", self.autofill_fields)

    def create_single_log_fields(self):
        self.single_log_fields = {
            "Reference Number": self.fields["Reference Number"],
            "Main Issue": self.fields["Main Issue"],
            "Action Taken": self.fields["Action Taken"],
            "Computer Name": self.fields["Computer Name"],
            "Computer Model/Type": self.fields["Computer Model/Type"],
            "Computer Serial Number": self.fields["Computer Serial Number"],
            "Computer OEM": self.fields["Computer OEM"],
            "Computer OS": self.fields["Computer OS"]
        }

    def create_multiple_log_fields(self):
        self.multiple_log_fields = {
            "Reference Number": self.fields["Reference Number"],
            "Main Issue": self.fields["Main Issue"],
            "Action Taken": self.fields["Action Taken"],
            "PC#": self.fields["PC#"],
            "Role": self.fields["Role"],
            "Computer Name": self.fields["Computer Name"],
            "Computer Model/Type": self.fields["Computer Model/Type"],
            "Computer Serial Number": self.fields["Computer Serial Number"],
            "Computer OEM": self.fields["Computer OEM"],
            "Computer OS": self.fields["Computer OS"]
        }

    def create_label_entry(self, label_text, row):
        label = ttk.Label(self.root, text=f"{label_text}:")
        label.grid(column=0, row=row, padx=10, pady=5, sticky='e')
        entry = ttk.Entry(self.root)
        entry.grid(column=1, row=row, padx=10, pady=5, sticky='we')
        self.customize_entry(entry)
        return entry

    def create_label_combo(self, label_text, values, row):
        label = ttk.Label(self.root, text=f"{label_text}:")
        label.grid(column=0, row=row, padx=10, pady=5, sticky='e')
        combo = ttk.Combobox(self.root, values=values)
        combo.grid(column=1, row=row, padx=10, pady=5, sticky='we')
        self.customize_entry(combo)
        return combo

    def create_label_text(self, label_text, row):
        label = ttk.Label(self.root, text=f"{label_text}:")
        label.grid(column=0, row=row, padx=10, pady=5, sticky='ne')
        text = scrolledtext.ScrolledText(self.root, width=40, height=5)
        text.grid(column=1, row=row, padx=10, pady=5, sticky='we')
        return text

    def create_buttons(self):
        clear_button = ttk.Button(self.root, text="Clear", command=self.clear_fields)
        clear_button.grid(column=0, row=15, padx=5, pady=10, sticky='we')

        save_button = ttk.Button(self.root, text="Save", command=self.save_to_html)
        save_button.grid(column=1, row=15, padx=5, pady=10, sticky='we')

        output_button = ttk.Button(self.root, text="Output Folder", command=self.open_output_folder)
        output_button.grid(column=2, row=16, columnspan=2, padx=5, pady=10, sticky='we')

        load_button = ttk.Button(self.root, text="Insert DC Records", command=self.load_excel_data)
        load_button.grid(column=3, row=15, padx=5, pady=10, sticky='we')

    def update_fields(self):
        for widget in self.fields.values():
            widget.grid_remove()

        if self.selection.get() == "Single":
            selected_fields = self.single_log_fields
            self.fields['PC#'].grid_forget()
            self.fields['Role'].grid_forget()
        else:
            selected_fields = self.multiple_log_fields

        for widget in selected_fields.values():
            widget.grid()

    def clear_fields(self):
        for widget in self.fields.values():
            if isinstance(widget, ttk.Entry) or isinstance(widget, ttk.Combobox):
                widget.delete(0, tk.END)
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.delete('1.0', tk.END)

    def customize_entry(self, widget):
        widget.config(font=('Calibri', 11), foreground='black')

    def autofill_fields(self, event):
        """Autofill fields based on computer name, using data from the loaded Excel file."""
        if self.excel_data is None:
            return

        computer_name = self.fields['Computer Name'].get().strip()
        if not computer_name:
            return

        matching_rows = self.excel_data[self.excel_data['Computer Name'] == computer_name]

        if not matching_rows.empty:
            if len(matching_rows) > 1:
                # Display warning message and let user choose which record to use
                messagebox.showwarning("Duplicate Record Found",
                                       "Multiple records found for this Computer Name. Please choose one.")
                # Popup for duplicates
                selection_popup = tk.Toplevel(self.root)
                selection_popup.title("Select Record")
                label = ttk.Label(selection_popup, text="Multiple records found, select one:")
                label.pack(padx=10, pady=10)

                # List of options (only showing computer model/type for simplicity)
                options = [
                    f"{row['Computer Model/Type']} (Serial: {row['Computer Serial Number']}, User: {row['Current Logged User']}, IP: {row['IP Address']})"
                    for _, row in matching_rows.iterrows()
                ]
                selection = tk.StringVar(value=options[0])

                for option in options:
                    ttk.Radiobutton(selection_popup, text=option, variable=selection, value=option).pack(anchor='w')

                def apply_selection():
                    selected_index = options.index(selection.get())
                    selected_row = matching_rows.iloc[selected_index]

                    # Autofill fields with selected row data
                    self.fields['Computer Model/Type'].delete(0, tk.END)
                    self.fields['Computer Model/Type'].insert(0, selected_row['Computer Model/Type'])

                    self.fields['Computer Serial Number'].delete(0, tk.END)
                    self.fields['Computer Serial Number'].insert(0, selected_row['Computer Serial Number'])

                    self.fields['Computer OEM'].set(selected_row['Computer OEM'])
                    self.fields['Computer OS'].set(selected_row['Computer OS'])

                    selection_popup.destroy()

                select_button = ttk.Button(selection_popup, text="Select", command=apply_selection)
                select_button.pack(pady=10)

                selection_popup.transient(self.root)
                selection_popup.grab_set()
                self.root.wait_window(selection_popup)
            else:
                # If only one match is found, autofill fields automatically
                self.fields['Computer Model/Type'].delete(0, tk.END)
                self.fields['Computer Model/Type'].insert(0, matching_rows.iloc[0]['Computer Model/Type'])

                self.fields['Computer Serial Number'].delete(0, tk.END)
                self.fields['Computer Serial Number'].insert(0, matching_rows.iloc[0]['Computer Serial Number'])

                self.fields['Computer OEM'].set(matching_rows.iloc[0]['Computer OEM'])
                self.fields['Computer OS'].set(matching_rows.iloc[0]['Computer OS'])

    def save_to_html(self):
        """Saves the form data to an HTML file."""
        selected_fields = self.single_log_fields if self.selection.get() == "Single" else self.multiple_log_fields

        data = {
            label: widget.get() if isinstance(widget, (ttk.Entry, ttk.Combobox)) else widget.get('1.0', tk.END).strip()
            for label, widget in selected_fields.items()
        }

        if not data['Main Issue']:
            messagebox.showwarning("Warning", "Main Issue field is empty")
            return

        if not data['Action Taken']:
            messagebox.showwarning("Warning", "Action Taken field is empty")
            return

        empty_fields = [label for label, value in data.items() if not value]
        if empty_fields:
            messagebox.showwarning("Warning", "Some fields are empty")
            return

        now = datetime.now().strftime("%m-%d-%Y")
        filename = f"{now}_SCTASK_Daily_close.html"
        filepath = os.path.join(self.output_folder, filename)

        log_entry_html = self.generate_log_entry_html(data)

        if os.path.exists(filepath):
            with open(filepath, "r") as file:
                existing_content = file.read()
            # Remove the closing tags from existing content
            existing_content = existing_content.rsplit("</body>", 1)[0].rsplit("</html>", 1)[0]
        else:
            existing_content = self.generate_html_header()

        existing_content += log_entry_html + "<hr>"

        html_content = existing_content + "</body></html>"

        with open(filepath, "w") as file:
            file.write(html_content)

        print(f"Data saved to {filename}")

        self.ask_add_more()

    def generate_log_entry_html(self, data):
        """Generates an HTML snippet for the log entry."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        main_issue = data.get('Main Issue', '').replace('\n', '<br>')
        action_taken = data.get('Action Taken', '').replace('\n', '<br>')

        log_entry_html = f"""
            <div class="log-entry">
                <p>{timestamp}</p>
                <p class="main-issue"><b>Main Issue:</b><br> {main_issue}</p>
                <p class="action-taken"><b>Action Taken:</b><br> {action_taken}</p>
                <b>Computer Name:</b> {data.get('Computer Name', '')}<br>
                <b>Computer Model/Type:</b> {data.get('Computer Model/Type', '')}<br>
                <b>Computer Serial Number:</b> {data.get('Computer Serial Number', '')}<br>
                <b>Computer OEM:</b> {data.get('Computer OEM', '')}<br>
                <b>Computer OS:</b> {data.get('Computer OS', '')}<br>
            </div>
        """
        return log_entry_html

    def generate_html_header(self):
        """Generates the header for the HTML log file."""
        return """
            <html>
            <head>
                <title>Daily Close Logs</title>
                <style>
                    body { font-family: Calibri, sans-serif; margin: 20px; }
                    .log-entry { margin-bottom: 20px; padding: 10px; border: 1px solid #ccc; border-radius: 5px; background-color: #f9f9f9; }
                    .main-issue { white-space: pre-wrap; }
                    .action-taken { white-space: pre-wrap; }
                </style>
            </head>
            <body>
        """

    def ask_add_more(self):
        """Asks the user if they want to add more logs."""
        if messagebox.askyesno("Add More", "Do you want to add more?"):
            self.clear_fields()
        else:
            self.clear_fields()

    def open_output_folder(self):
        """Opens the output folder for the saved HTML files."""
        folder_path = self.output_folder
        if os.name == 'nt':
            os.startfile(folder_path)
        elif os.name == 'posix':
            subprocess.call(['open', folder_path])

    def get_output_folder(self):
        """Gets the output folder for saving HTML files."""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.abspath(os.path.dirname(__file__))

    def load_excel_data(self):
        """Loads data from an Excel file."""
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.excel_data = pd.read_excel(file_path)
                messagebox.showinfo("Success", "Excel data loaded successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel data: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LogSelectorApp(root)
    root.mainloop()
