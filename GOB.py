#S1-GOB was created by Ran Jacobi, 2023.
# V1.7
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from getapi import APIClient
from getexcel import ExcelProcessor
import os
from datetime import datetime
import sys
from tkinter import *
from tkmacosx import Button
from tkinter import messagebox
import threading



class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")
        self.widget.yview("end")

class GUIInput:
    def __init__(self):
        # Create GUI
        self.process_rawdata = None
        self.root = tk.Tk()
        self.root.title("S1-GOB: Guided On Boarding")

        # Function to open GUI in the middle of the screen
        def open_gui_centered():
            window_width = 600
            window_height = 700
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = int((screen_width / 2) - (window_width / 2))
            y = int((screen_height / 2) - (window_height / 2))
            self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Call the function to open the GUI window centered
        open_gui_centered()

        self.progress_window = None  # Initialize progress_window attribute
        self.cancel_button = None

        # Create notebook with tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(padx=10, pady=10)

        # Create a custom style for the tabs
        style = ttk.Style()
        style.configure('TNotebook.Tab', padding=(35, 10))
        style.map('TNotebook.Tab', background=[('selected', 'purple')])

        # Create tab for S1-HealthCheck tool
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Grab CSV's")

        # Create a custom style for the button
        style = ttk.Style()
        style.configure('Purple.TButton', background='purple', foreground='black')

        # Create the "About" button using tkmacosx:
        about_button = Button(self.tab1, command=self.about, text="About", bg='purple', fg='black')
        about_button.pack(side=tk.TOP, padx=10, pady=10)

        # Create tooltip for the API Token entry box
        self.token_label = ttk.Label(self.tab1, text="API Token:")
        self.token_label.pack(padx=10, pady=2)
        self.token_entry = ttk.Combobox(self.tab1)
        self.token_entry.pack(padx=10, pady=2)
        self.token_entry.insert(0, "")
        self.add_tooltip(self.token_entry, "Enter your API token obtained from the console (User or Service User token)")

        # Create tooltip for the Domain entry box
        self.domain_label = ttk.Label(self.tab1, text="Domain:")
        self.domain_label.pack(padx=10, pady=2)
        self.domain_entry = ttk.Combobox(self.tab1)
        self.domain_entry.pack(padx=10, pady=2)
        self.domain_entry.insert(0, "")
        self.add_tooltip(self.domain_entry, "Enter the domain name in this format '[].sentinelone.net'")

        # Create tooltip for the User entry box
        self.user_label = ttk.Label(self.tab1, text="User: (Optional)")
        self.user_label.pack(padx=10, pady=2)
        self.user_entry = ttk.Combobox(self.tab1)
        self.user_entry.pack(padx=10, pady=2)
        self.add_tooltip(self.user_entry, "Enter your user email. This field is optional")

        # Create the "Start" button using tkmacosx:
        self.create_start_button = Button(self.tab1, text="Start", bg='green', fg='white', font=('Helvetica', 15),
                                          command=self.start_script)
        self.create_start_button.pack(padx=20, pady=20)

        # Set permissions for the RawData folder
        try:
            os.chmod("RawData", 0o700)
        except FileNotFoundError:
            pass

        # Create tab for Excel Processor
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="Excel Processor")

        # Create a custom style for the Refresh button
        style = ttk.Style()
        style.configure('Purple.TButton', foreground='white', background='purple', font=('Helvetica', 15), width=20)
        refresh_button = ttk.Button(self.tab2, text="Refresh folder", command=self.refresh, style='Purple.TButton')
        refresh_button.pack(padx=10, pady=10)

        # Set history for each combobox
        self.token_entry['values'] = self.load_from_history("token")
        self.domain_entry['values'] = self.load_from_history("domain")
        self.user_entry['values'] = self.load_from_history("user")

        # Create a label to show the files in the RawData folder
        label = ttk.Label(self.tab2, text="Files in the 'RawData' folder:")
        label.pack()

        # Create a listbox to show the file names
        self.listbox = tk.Listbox(self.tab2, height=10, width=50)
        self.listbox.pack(padx=10, pady=10)

        # Check if the RawData folder exists under the user's home directory in S1GOB/RawData
        raw_data_folder = os.path.expanduser("~/S1GOB/RawData")
        if os.path.exists(raw_data_folder):
            # Get the list of files in the folder
            files = os.listdir(raw_data_folder)
            # Add each file name to the listbox
            for file in files:
                self.listbox.insert(tk.END, file)
                # Add a binding to the file name to open it when double-clicked
                self.listbox.bind('<Double-Button-1>',
                                  lambda event, f=file: subprocess.run(["open", os.path.join(raw_data_folder, f)]))
        else:
            # Show an error message if the folder doesn't exist
            self.listbox.insert(tk.END, "The 'RawData' folder does not exist.")

        # Create the "create_excel_file" button using tkmacosx:
        create_excel_button = Button(self.tab2, text="Create Excel File", bg='green', fg='white',
                                     font=('Helvetica', 15),
                                     borderless=1, command=self.create_excel_file)
        create_excel_button.pack(padx=20, pady=20)

        # Create console-like area for output
        self.console = tk.Text(self.root, height=30, width=70)
        self.console.pack(padx=10, pady=10)

        # Redirect stdout to the console
        sys.stdout = TextRedirector(self.console, "stdout")
        sys.stderr = TextRedirector(self.console, "stderr")

        # Run the GUI
        self.root.mainloop()

    # Define the "About" function
    def about(self):
        # Create a message box with information about the application
        messagebox.showinfo("About S1-GOB",
                            "SS1-GOB is a tool that provides comprehensive health checks for customer's SentinelOne environment, enabling you to identify and present issues to your customer quickly and efficiently.\n\nVersion: 1.5.0\n\nCopyright 2023, Ran Jacobi")

    def refresh(self):
        # Clear listbox
        self.listbox.delete(0, tk.END)

        # Check if the RawData folder exists under the user's home directory in S1GOB/RawData
        raw_data_folder = os.path.expanduser("~/S1GOB/RawData")
        if os.path.exists(raw_data_folder):
            # Get the list of files in the folder
            files = os.listdir(raw_data_folder)
            # Add each file name to the listbox
            for file in files:
                self.listbox.insert(tk.END, file)
                # Add a binding to the file name to open it when double-clicked
                self.listbox.bind('<Double-Button-1>',
                                  lambda event, f=file: subprocess.run(["open", os.path.join(raw_data_folder, f)]))
        else:
            # Show an error message if the folder doesn't exist
            self.listbox.insert(tk.END, "The 'RawData' folder does not exist.")

    def create_excel_file(self):
        # Check if the raw data folder exists and isn't empty
        raw_data_folder = os.path.expanduser("~/S1GOB/RawData")
        if os.path.exists(raw_data_folder) and os.listdir(raw_data_folder):
            # Check if the raw data folder contains any CSV files
            csv_files = [f for f in os.listdir(raw_data_folder) if f.endswith(".csv")]
            if csv_files:
                # Create a new thread to run the ExcelProcessor
                thread = threading.Thread(target=self.run_excel_processor)
                thread.start()
            else:
                print("No CSV files found in the 'RawData' folder.")
        else:
            print("The 'RawData' folder either doesn't exist or is empty.")

    def run_excel_processor(self):
        # Create an ExcelProcessor object and process the CSV files
        self.excel_processor = ExcelProcessor("RawData")
        self.excel_processor.process_csv_files()

        # Destroy the progress window and reset the instance variables
        if self.progress_window:
            self.progress_window.destroy()
        self.progress_window = None
        self.cancel_button = None

    def run_excel_processor(self):
        # Create an ExcelProcessor object and process the CSV files
        self.excel_processor = ExcelProcessor("RawData")
        self.excel_processor.process_csv_files()

        # Destroy the progress window and reset the instance variables
        if self.progress_window:
            self.progress_window.destroy()
        self.progress_window = None
        self.cancel_button = None

    def start_script(self):
        self.create_start_button.config(state=DISABLED)  # Disable the button after it is clicked
        self.token = self.token_entry.get()
        self.domain = self.domain_entry.get()
        self.user = self.user_entry.get()

        # Validate that required fields are not empty
        if not self.token or not self.domain:
            messagebox.showerror("Error", "Please enter both token and domain fields.")
            self.create_start_button.config(state=NORMAL)
            return

        # Save inputs to history
        self.save_to_history("token", self.token)
        self.save_to_history("domain", self.domain)
        self.save_to_history("user", self.user)

        # Create progress bar
        progress_bar = ttk.Progressbar(self.root, orient="horizontal", mode="indeterminate", length=200)
        progress_bar.pack(fill="x", padx=10, pady=10)
        progress_bar.start()

        # Start the script in a separate thread
        script_thread = threading.Thread(target=self.run_script, args=(progress_bar,))
        script_thread.start()

    def save_to_history(self, field, value):
        # Get the existing history for the field
        history = self.load_from_history(field)

        # Remove the value if it already exists
        if value in history:
            history.remove(value)

        # Add the value at the beginning of the history list
        history.insert(0, value)

        # Limit the history to 10 items
        history = history[:10]

        # Save the updated history to a file
        with open(f"{field}_history.txt", "w") as file:
            file.write("\n".join(history))

        # Update the combobox values with the latest history
        if field == "token":
            self.token_entry['values'] = history
        elif field == "domain":
            self.domain_entry['values'] = history
        elif field == "user":
            self.user_entry['values'] = history

    def load_from_history(self, field):
        try:
            with open(f"{field}_history.txt", "r") as file:
                history = file.read().splitlines()
        except FileNotFoundError:
            history = []

        return history

    def add_tooltip(self, widget, text):
        widget.bind("<Enter>", lambda event: self.show_tooltip(text))
        widget.bind("<Leave>", lambda event: self.hide_tooltip())

    def show_tooltip(self, text):
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{self.root.winfo_pointerx()}+{self.root.winfo_pointery() + 20}")

        label = ttk.Label(self.tooltip, text=text, background="lightyellow", relief="solid", borderwidth=1,
                          font=("Arial", "10", "normal"))
        label.pack()

    def hide_tooltip(self):
        self.tooltip.destroy()

    def run_script(self, progress_bar):
        # Get token, domain and user
        token = self.token
        domain = self.domain
        user = self.user

        token_header = 'APIToken ' + token
        self.user = user
        customer_endpoint = "https://" + domain

        print(token_header, file=sys.stdout)
        print(self.user, file=sys.stdout)

        print(customer_endpoint, file=sys.stdout)

        try:
            # Create an object of the class
            my_object = APIClient(token, domain, user)

            # Call the createLevelsDF() function on the object
            level_account_df, level_site_df, level_group_df = my_object.createLevelsDF()

            agents_df = my_object.httpGetPagination("/web/api/v2.1/agents")
            policy_df = my_object.getAllPolicies(level_account_df, level_site_df, level_group_df)
            users_df = my_object.httpGetPagination("/web/api/v2.1/users")
            exclusions_df = my_object.getAllLevels("/web/api/v2.1/exclusions", level_account_df, level_site_df,
                                                   level_group_df)
            print("\n\nStarting to Group Items...")

        except KeyError as e:
            print("An error occurred: Data key not found. Please check your input.")
            print("Possible reasons:")
            print("- Invalid API token")
            print("- Incorrect domain name format")
            print("- VPN connection issue")
            # Destroy progress bar
            progress_bar.stop()
            progress_bar.destroy()
            self.create_start_button.config(state=NORMAL)

        except UnboundLocalError as e:
            # Suppress the specific error message
            pass

        except Exception as e:
            # Suppress any other error messages
            sys.stderr = open("/dev/null", "w")

        # Restore the standard error output
        sys.stderr = sys.__stderr__

        # installed_apps_df = getAllLevels("/web/api/v2.1/installed-applications",level_account_df,level_site_df, level_group_df)
        # print(installed_apps_df.columns.values.tolist())
        # print(installed_apps_df)

        try:
            agent_counts = agents_df.groupby(['agentVersion', 'osType']).size()
        except KeyError:
            print("Unable to group by 'agentVersion' and 'osType'")
            agent_counts = None

        if agent_counts is not None:
            try:
                agent_counts.to_csv(os.path.expanduser("~/agent_counts-{}.csv".format(domain)))
            except AttributeError:
                print("Error: 'agent_counts' is None")

        agents_df.to_csv(os.path.expanduser("~/agents-" + domain + ".csv"))
        users_df.to_csv(os.path.expanduser("~/users-" + domain + ".csv"))
        policy_df.to_csv(os.path.expanduser("~/policies-" + domain + ".csv"))
        exclusions_df.to_csv(os.path.expanduser("~/exclusions-" + domain + ".csv"))
        # installed_apps_df.to_csv(os.path.expanduser("~/installed_apps-" + domain + ".csv"))

        # Create S1GOB folder if it does not exist
        s1gob_folder = os.path.expanduser("~/S1GOB")
        if not os.path.exists(s1gob_folder):
            os.makedirs(s1gob_folder)

        # Create RawData folder inside S1GOB folder if it does not exist
        raw_data_folder = os.path.join(s1gob_folder, "RawData")
        if not os.path.exists(raw_data_folder):
            os.makedirs(raw_data_folder)

        # Check if there are any files in the RawData folder
        if len(os.listdir(raw_data_folder)) > 0:
            # Create a new folder with the current date and time
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_folder_name = f"RawData_{timestamp}"
            new_folder_path = os.path.join(s1gob_folder, new_folder_name)

            # Ensure the new folder name is unique
            counter = 1
            while os.path.exists(new_folder_path):
                new_folder_name = f"RawData_{timestamp}_{counter}"
                new_folder_path = os.path.join(s1gob_folder, new_folder_name)
                counter += 1

            os.makedirs(new_folder_path)

            # Move all files from RawData to the new folder
            files = os.listdir(raw_data_folder)
            for file in files:
                file_path = os.path.join(raw_data_folder, file)
                shutil.move(file_path, new_folder_path)

            # Move the new folder one level below S1GOB
            try:
                shutil.move(new_folder_path, s1gob_folder)
            except shutil.Error as e:
                print(f"Error moving '{new_folder_path}': {str(e)}")

        # Move all CSV files from the user's home directory to RawData folder
        user_home_directory = os.path.expanduser("~")
        csv_files = [file for file in os.listdir(user_home_directory) if file.endswith(".csv")]
        failed_files = []
        for file in csv_files:
            file_path = os.path.join(user_home_directory, file)
            destination_path = os.path.join(raw_data_folder, file)
            try:
                shutil.move(file_path, destination_path)
                print(f"Moved '{file}' to '{destination_path}'")
            except shutil.Error as e:
                failed_files.append(file)
                print(f"Error moving '{file}': {str(e)}")

        if failed_files:
            print("Error: Failed to move the following files to the RawData folder:")
            for file in failed_files:
                print(file)

        print("\n\n")
        print("The Script is Done! The RawData folder is ready")
        print("\n")
        print("Location of RawData folder:", os.path.abspath(raw_data_folder))
        self.create_start_button.config(state=NORMAL)

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None

    def show_tooltip(self, event):
        x, y, _, _ = event.widget.bbox("insert")
        x += event.widget.winfo_rootx() - 65  # Adjust the x-coordinate to display to the left of the widget
        y += event.widget.winfo_rooty() - 25  # Adjust the y-coordinate to display above the widget

        self.tooltip = tk.Toplevel(event.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        self.tooltip.configure(background="#9b59b6")  # Set the background color to purple

        if sys.platform == "darwin":
            self.tooltip.attributes("-type", "tooltip")  # Set the window type to tooltip on macOS

        label = ttk.Label(self.tooltip, text=self.text, background="#9b59b6", foreground="#ffffff", relief="solid",
                          borderwidth=1)
        label.pack()
    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()

gui = GUIInput()
gui.root.mainloop()
exit()
