import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import shutil
import os
import sys
from tkinter import ttk, filedialog, messagebox
from getapi import APIClient
from getexcel import ExcelProcessor
import Pmw


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
        self.root.iconbitmap("./icons/icon.ico")

        self.progress_window = None  # Initialize progress_window attribute
        self.cancel_button = None

        # Set application icon for different platforms
        if os.name == 'nt':  # For Windows
            icon_path = "./icons/icon.ico"
            self.root.iconbitmap(default=icon_path)
        elif os.name == 'posix':  # For macOS and Linux
            icon_path = "./icons/icon.png"
            self.root.iconphoto(True, tk.PhotoImage(file=icon_path))

        # Create notebook with tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(padx=10, pady=10)

        # Create a custom style for the tabs
        style = ttk.Style()
        style.configure('TNotebook.Tab', padding=(30, 10))

        # Create tab for S1-HealthCheck tool
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Grab CSV's")

        # Create the "About" button
        about_icon = tk.PhotoImage(file="./icons/info.png")
        about_button = ttk.Button(self.tab1, command=self.about, image=about_icon)
        about_button.pack(side=tk.TOP, padx=10, pady=10)


        # Create balloon help for the API Token entry box
        self.token_label = ttk.Label(self.tab1, text="API Token:")
        self.token_label.pack(padx=10, pady=2)
        self.token_entry = ttk.Entry(self.tab1)
        self.token_entry.pack(padx=10, pady=2)
        self.token_entry.insert(0, "dACAAjqKaQzosO4Zmw8dzGXAWLVA5xDP21ZnOtgwLlFztLsGXqfe8pyBQcuU55HMtDSzwjTAZOyTWrRT")
        self.token_balloon = Pmw.Balloon(self.tab1)
        self.token_balloon.bind(self.token_entry, "Enter your API token here.")

        # Create balloon help for the Domain entry box
        self.domain_label = ttk.Label(self.tab1, text="Domain:")
        self.domain_label.pack(padx=10, pady=2)
        self.domain_entry = ttk.Entry(self.tab1)
        self.domain_entry.pack(padx=10, pady=2)
        self.domain_entry.insert(0, "usea1-007.sentinelone.net")
        self.domain_balloon = Pmw.Balloon(self.tab1)
        self.domain_balloon.bind(self.domain_entry, "Enter your domain here.")

        # Create balloon help for the User entry box
        self.user_label = ttk.Label(self.tab1, text="User: (Optional)")
        self.user_label.pack(padx=10, pady=2)
        self.user_entry = ttk.Entry(self.tab1)
        self.user_entry.pack(padx=10, pady=2)
        self.user_balloon = Pmw.Balloon(self.tab1)
        self.user_balloon.bind(self.user_entry, "Enter your username here.")


        # Create a custom style for the button
        style = ttk.Style()
        style.configure('Purple.TButton', foreground='white', background='purple', font=('Helvetica', 15), width=20)
        create_start_button = ttk.Button(self.tab1, text="Start", style='Purple.TButton',
                                         command=self.start_script)
        create_start_button.pack(padx=10, pady=10)

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
        refresh_button.pack(padx=10, pady=10, side=tk.TOP)



        # Create a label to show the files in the RawData folder
        label = ttk.Label(self.tab2, text="Files in the 'RawData' folder:")
        label.pack()

        # Add a balloon to the label
        label_balloon = Pmw.Balloon(self.tab2)
        label_balloon.bind(label, "This box displays the list of files in the 'RawData' folder.\nYou can click to open the files")

        # Create a listbox to show the file names
        self.listbox = tk.Listbox(self.tab2, height=10, width=50)
        self.listbox.pack(padx=10, pady=10)

        # Check if the RawData folder exists
        raw_data_folder = "RawData"
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

        # Create a custom style for the button
        style = ttk.Style()
        style.configure('Purple.TButton', foreground='white', background='purple', font=('Helvetica', 15), width=20)
        create_excel_button = ttk.Button(self.tab2, text="Create Excel File", style='Purple.TButton',
                                         command=self.create_excel_file)
        create_excel_button.pack(padx=10, pady=10)

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
            messagebox.showinfo("About S1-GOB", "SS1-GOB is a tool that provides comprehensive health checks for customer's SentinelOne environment, enabling you to identify and present issues to your cusomter quickly and efficiently.\n\nVersion: 1.5.0\n\nCopyright 2023, Ran Jacobi")
    def refresh(self):
        # Clear listbox
        self.listbox.delete(0, tk.END)

        # Check if the RawData folder exists
        raw_data_folder = "RawData"
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
        if os.path.exists("RawData") and os.listdir("RawData"):
            # Check if the raw data folder contains any CSV files
            csv_files = [f for f in os.listdir("RawData") if f.endswith(".csv")]
            if csv_files:
                # Create a new thread to run the ExcelProcessor
                thread = threading.Thread(target=self.run_excel_processor)
                thread.start()
            else:
                print("No CSV files found in the raw data folder.")
        else:
            print("The raw data folder either doesn't exist or is empty.")

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
        self.token = self.token_entry.get()
        self.domain = self.domain_entry.get()
        self.user = self.user_entry.get()

        # Validate that required fields are not empty
        if not self.token or not self.domain:
            messagebox.showerror("Error", "Please enter both token and domain fields.")
            return

        # Display message that script is starting to run
        messagebox.showinfo("S1-GOB", "The script is starting to run...")

        # Create progress bar
        progress_bar = ttk.Progressbar(self.root, mode='indeterminate')
        progress_bar.pack(fill='x', padx=10, pady=10)
        progress_bar.start()

        # Start the script in a separate thread
        script_thread = threading.Thread(target=self.run_script, args=(progress_bar,))
        script_thread.start()

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

        # create an object of the class
        my_object = APIClient(token, domain, user)
        # call the createLevelsDF() function on the object
        level_account_df, level_site_df, level_group_df = my_object.createLevelsDF()
        agents_df = my_object.httpGetPagination("/web/api/v2.1/agents")
        policy_df = my_object.getAllPolicies(level_account_df, level_site_df, level_group_df)
        users_df = my_object.httpGetPagination("/web/api/v2.1/users")
        exclusions_df = my_object.getAllLevels("/web/api/v2.1/exclusions", level_account_df, level_site_df, level_group_df)
        print(exclusions_df.columns.values.tolist())
        print(exclusions_df)

        # installed_apps_df = getAllLevels("/web/api/v2.1/installed-applications",level_account_df,level_site_df, level_group_df)
        # print(installed_apps_df.columns.values.tolist())
        # print(installed_apps_df)

        try:
            agent_counts = agents_df.groupby(['agentVersion', 'osType']).size()
        except KeyError:
            print("Unable to group by 'agentVersion' and 'osType'")
            agent_counts = None
        try:
            agent_counts.to_csv("agent_counts-" + domain + ".csv")
        except AttributeError:
            print("Error: 'agent_counts' is None")

        agents_df.to_csv("agents-" + domain + ".csv")
        users_df.to_csv("users-" + domain + ".csv")
        policy_df.to_csv("policies-" + domain + ".csv")
        exclusions_df.to_csv("exclusions-" + domain + ".csv")
        # installed_apps_df.to_csv("installed_apps-"+domain+".csv")

        # Create RawData folder if it does not exist
        if not os.path.exists("RawData"):
            os.makedirs("RawData")

        # Move all CSV files to RawData folder
        for file in os.listdir():
            if file.endswith(".csv"):
                shutil.move(file, os.path.join("RawData", file))

        print("\n\n\n\n\n\n\n\n\n\n\n\n\n\n")
        print("Agents DF:\n", agents_df)
        print("Agents Summary:\n", agent_counts)
        print("Users DF:\n", users_df)
        print("Policy DF:\n", policy_df)
        # print("Installed Apps\n",installed_apps_df)
        print("Exclusions DF:\n", exclusions_df)
        print("The Script is Done! The RawData folder is ready")

        # Destroy progress bar
        progress_bar.stop()
        progress_bar.destroy()

        # Use after() method to periodically update GUI
        def update_gui():

            # Display message box
            messagebox.showinfo("Script Status", "Script finished successfully")

        # Update GUI every 500 milliseconds
        self.root.after(500, update_gui)

gui = GUIInput()
gui.root.mainloop()
exit()
