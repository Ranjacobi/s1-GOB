# With GUI!!!

import pandas as pd
import requests
import json
import argparse
import time
import tkinter as tk
from tkinter import messagebox
import tkinter.scrolledtext as tkst
import sys
from tkinter import messagebox
from io import StringIO
import os
import shutil
import time
import threading
from tkinter import ttk


class TerminalWindow:
    def __init__(self, parent):
        self.parent = parent
        self.text_widget = tk.Text(parent)
        self.text_widget.pack(fill="both", expand=True)
        self.stdout_backup = sys.stdout
        self.stderr_backup = sys.stderr
        sys.stdout = self
        sys.stderr = self

    def write(self, message):
        self.text_widget.insert("end", message)
        self.text_widget.see("end")  # scroll to the bottom
        self.flush()

    def flush(self):
        pass

    def close(self):
        sys.stdout = self.stdout_backup
        sys.stderr = self.stderr_backup


class RawDataTab:
    def __init__(self, parent):
        self.parent = parent
        self.frame = tk.Frame(parent)
        self.frame.pack(fill="both", expand=True)

        self.refresh_button = tk.Button(self.frame, text="Refresh", command=self.refresh)
        self.refresh_button.pack()

        self.listbox = tk.Listbox(self.frame)
        self.listbox.pack(fill="both", expand=True)

        self.refresh()

    def refresh(self):
        self.listbox.delete(0, tk.END)
        if not os.path.exists("RawData"):
            self.listbox.insert(tk.END, "RawData folder does not exist.")
            return
        files = os.listdir("RawData")
        if not files:
            self.listbox.insert(tk.END, "RawData folder is empty.")
            return
        for file in files:
            self.listbox.insert(tk.END, file)


class GUIInput:
    def __init__(self):
        self.token = ""
        self.domain = ""
        self.user = ""

        self.root = tk.Tk()
        self.root.title("API Scraper v1.0")

        # Notebook (tabbed layout)
        self.notebook = ttk.Notebook(self.root, style='My.TNotebook')
        self.notebook.pack(fill='both', expand=True)
        self.notebook.bind("<Button-3>", lambda e: self.popup_menu(e, self.notebook))

        style = ttk.Style()
        style.configure('My.TNotebook', tabposition='n')
        style.configure('My.TNotebook.Tab', padding=[30, 10])

        # API input tab
        api_tab = tk.Frame(self.notebook)
        self.notebook.add(api_tab, text="API Input")

        # Token input
        tk.Label(api_tab, text="Token without the word 'Token'").pack()
        self.token_entry = tk.Entry(api_tab)
        self.token_entry.pack()

        # Domain input
        tk.Label(api_tab, text="Domain").pack()
        self.domain_entry = tk.Entry(api_tab)
        self.domain_entry.pack()

        # User input
        tk.Label(api_tab, text="User (optional)").pack()
        self.user_entry = tk.Entry(api_tab)
        self.user_entry.pack()

        # Info button with tooltip
        info_icon = tk.PhotoImage(file="info.png")
        info_button = tk.Button(api_tab, image=info_icon, command=self.show_tooltip)
        info_button.image = info_icon
        info_button.pack()

        # Submit button
        tk.Button(api_tab, text="Submit", command=self.start_script).pack()

        # RawData tab
        rawdata_tab = RawDataTab(self.notebook)
        self.notebook.add(rawdata_tab.frame, text="RawData")

        # Add Terminal tab
        terminal_tab = ttk.Frame(self.root, name="terminal_tab")
        ttk.Label(terminal_tab, text="Terminal").pack()
        self.terminal = TerminalWindow(terminal_tab)
        terminal_tab.pack()

        self.root.mainloop()

    def show_tooltip(self):
        tooltip_text = "Enter your API token without the word 'Token'"
        messagebox.showinfo("API Token", tooltip_text)

    def start_script(self):
        self.token = self.token_entry.get()
        self.domain = self.domain_entry.get()
        self.user = self.user_entry.get()

        # Validate that required fields are not empty
        if not self.token or not self.domain:
            messagebox.showerror("Error", "Please enter both token and domain fields.")
            return

        # Display message that script is starting to run
        messagebox.showinfo("API Scraper v1.0", "The script is starting to run...")

        # Start the script in a separate thread
        script_thread = threading.Thread(target=self.run_script)
        script_thread.start()

    def run_script(self):
        # Create progress bar
        progress_bar = ttk.Progressbar(self.root, mode='indeterminate')
        progress_bar.pack(fill='x', padx=10, pady=10)
        progress_bar.start()

        # Get token, domain and user
        token = self.token
        domain = self.domain
        user = self.user

        token_header = 'APIToken ' + token
        real_user = user
        customer_endpoint = "https://" + domain

        print(token_header, file=sys.stdout)
        print(real_user, file=sys.stdout)
        print(customer_endpoint, file=sys.stdout)

        # Rest of the code remains the same

        def httpGet(endpoint):
            url = customer_endpoint + endpoint
            headers = {'Authorization': token_header}
            resp = requests.get(url, headers=headers)
            print(resp)
            print(resp.text)
            return resp

        def httpGetPagination(endpoint):
            query_params = "?limit=100"
            nextPage = True
            df_list = []
            FirstRun = True
            sleepcount = 0
            while nextPage:
                print(endpoint + query_params)
                response_json = json.loads(httpGet(endpoint + query_params).text)
                data = response_json['data']
                df_list.append(pd.DataFrame.from_records(data))
                pagination = response_json['pagination']
                if FirstRun:
                    total_left = pagination['totalItems']
                    FirstRun = False
                try:
                    cursor = pagination['nextCursor']
                except:
                    cursor = None
                if cursor:
                    query_params = "?limit=100&cursor=" + cursor
                else:
                    nextPage = False

                total_left -= 100
                if total_left < 0:
                    total_left = 0

                sleepcount += 1
                if sleepcount >= 5:
                    time.sleep(1)
                    sleepcount = 0
                print("Items Remaining: ", total_left)
            raw_df = pd.concat(df_list)
            print(endpoint + "\n", raw_df)
            return raw_df

        def httpGetPaginationIds(endpoint, level, level_id_df):
            param = "siteIds"

            if level == "accounts":
                param = "accountIds"
            elif level == "sites":
                param = "siteIds"
            elif level == "groups":
                param = "groupIds"

            df_list = []

            print(level)
            print(param)

            for level_id in level_id_df["id"].to_list():
                query_params_base = "?limit=100" + "&" + param + "=" + level_id
                query_params = "?limit=100" + "&" + param + "=" + level_id
                nextPage = True

                FirstRun = True
                sleepcount = 0
                while nextPage:
                    print(endpoint + query_params)
                    response_json = json.loads(httpGet(endpoint + query_params).text)
                    data = response_json['data']
                    dataframe_tmp = pd.DataFrame.from_records(data)
                    dataframe_tmp["level"] = level
                    dataframe_tmp["level_id"] = level_id
                    df_list.append(dataframe_tmp)
                    pagination = response_json['pagination']
                    if FirstRun:
                        total_left = pagination['totalItems']
                        FirstRun = False
                    try:
                        cursor = pagination['nextCursor']
                    except:
                        cursor = None
                    if cursor:
                        query_params = query_params_base + "&cursor=" + cursor
                    else:
                        nextPage = False

                    total_left -= 100
                    if total_left < 0:
                        total_left = 0

                    sleepcount += 1
                    if sleepcount >= 5:
                        time.sleep(1)
                        sleepcount = 0
                    print("Items Remaining: ", total_left)
            raw_df = pd.concat(df_list)
            try:
                tmp_policy_df = pd.merge(raw_df, level_df, how='left', left_on='level_id', right_on='id')
                print(tmp_policy_df)
                raw_df = tmp_policy_df
            except:
                print("GlobalPolicy")

            print(endpoint + "\n", raw_df)
            return raw_df

        def getGlobalPolicies():
            tmp_global_list = []
            endpoint = "/web/api/v2.1/tenant/policy"
            query_params = "?limit=100"
            response_json = json.loads(httpGet(endpoint + query_params).text)
            data = response_json['data']
            response_fw = httpGet("/web/api/v2.1/firewall-control/configuration")

            if response_fw.status_code == 200:
                response_fwjson = json.loads(response_fw.text)
                fw_data = response_fwjson['data']
                data['fw_enabled'] = fw_data['enabled']
                data['fw_inheritAllFirewallRules'] = fw_data['inheritAllFirewallRules']
                data['fw_inheritedFrom'] = fw_data['inheritedFrom']
                data['fw_inherits'] = fw_data['inherits']
                data['fw_inheritSettings'] = fw_data['inheritSettings']
                data['fw_locationAware'] = fw_data['locationAware']
                # data['fw_reportBlocked'] = fw_data['reportBlocked']
                data['fw_selectedTags'] = fw_data['selectedTags']
            else:
                data['fw_enabled'] = None
                data['fw_inheritAllFirewallRules'] = None
                data['fw_inheritedFrom'] = None
                data['fw_inherits'] = None
                data['fw_inheritSettings'] = None
                data['fw_locationAware'] = None
                # data['fw_reportBlocked'] = None
                data['fw_selectedTags'] = None

            response_dc = httpGet("/web/api/v2.1/device-control/configuration")
            if response_dc.status_code == 200:
                response_dcjson = json.loads(response_dc.text)
                dc_data = response_dcjson['data']
                data['dc_disableBleCommunication'] = dc_data['disableBleCommunication']
                data['dc_disableRfcomm'] = dc_data['disableRfcomm']
                data['dc_disallowAccessPermissionControl'] = dc_data['disallowAccessPermissionControl']
                data['dc_enabled'] = dc_data['enabled']
                data['dc_inheritedFrom'] = dc_data['inheritedFrom']
                data['dc_inherits'] = dc_data['inherits']
                data['dc_reportApproved'] = dc_data['reportApproved']
                data['dc_reportBlocked'] = dc_data['reportBlocked']
                data['dc_reportReadOnly'] = dc_data['reportReadOnly']
            else:
                data['dc_dc_disableBleCommunication'] = None
                data['dc_disableRfcomm'] = None
                data['dc_disallowAccessPermissionControl'] = None
                data['dc_enabled'] = None
                data['dc_inheritedFrom'] = None
                data['dc_inherits'] = None
                data['dc_reportBlocked'] = None
                data['dc_reportReadOnly'] = None

            tmp_global_list.append(data)
            data_df = pd.DataFrame.from_records(tmp_global_list)
            return data_df

        def createLevelsDF():
            id_url = "/web/api/v2.1/accounts"
            account_df = httpGetPagination(id_url)
            id_url = "/web/api/v2.1/sites"
            query_params = "?limit=100"
            nextPage = True
            df_list = []
            while nextPage:
                print(id_url + query_params)
                response_json = json.loads(httpGet(id_url + query_params).text)
                tmp_data = response_json['data']
                data = tmp_data['sites']
                df_list.append(pd.DataFrame.from_records(data))
                pagination = response_json['pagination']
                try:
                    cursor = pagination['nextCursor']
                except:
                    cursor = None
                if cursor:
                    query_params = "?limit=100&cursor=" + cursor
                else:
                    nextPage = False
            raw_df = pd.concat(df_list)
            site_df = raw_df

            id_url = "/web/api/v2.1/groups"
            group_df = httpGetPagination(id_url)
            id_url = "/web/api/v2.1/accounts"
            account_df = httpGetPagination(id_url)

            id_url = "/web/api/v2.1/sites"
            query_params = "?limit=100"
            nextPage = True
            df_list = []
            sleepcount = 0
            while nextPage:
                print(id_url + query_params)
                response_json = json.loads(httpGet(id_url + query_params).text)
                tmp_data = response_json['data']
                data = tmp_data['sites']
                df_list.append(pd.DataFrame.from_records(data))
                pagination = response_json['pagination']
                try:
                    cursor = pagination['nextCursor']
                except:
                    cursor = None

                sleepcount += 1
                if sleepcount >= 5:
                    time.sleep(1)
                    sleepcount = 0

                if cursor:
                    query_params = "?limit=100&cursor=" + cursor
                else:
                    nextPage = False
            raw_df = pd.concat(df_list)
            site_df = raw_df

            new_account_df = account_df[['id', 'name']].copy()
            site_df = site_df[['id', 'accountId', 'name']].copy()
            group_df = group_df[['id', 'siteId', 'name']].copy()
            print("SiteDF:\n", site_df)
            print("GroupDF:\n", group_df)
            new_site_df = pd.merge(site_df, new_account_df, how='left', left_on='accountId', right_on='id')

            new_site_df['Scope'] = new_site_df["name_y"] + "\\" + new_site_df["name_x"]
            new_site_df = new_site_df.rename(columns={"id_x": "id"})
            del new_site_df['accountId']
            del new_site_df['name_x']
            del new_site_df['id_y']
            del new_site_df['name_y']
            print("PreJoinSite")
            print(new_site_df)

            new_group_df = pd.merge(group_df, new_site_df, how='left', left_on='siteId', right_on='id')

            print("PreNameJoin")
            print(new_group_df)
            new_group_df['Scope_tmp'] = new_group_df['Scope'] + "\\" + new_group_df['name']
            del new_group_df['siteId']
            del new_group_df['name']
            del new_group_df['Scope']
            del new_group_df['id_y']
            new_group_df = new_group_df.rename(columns={"id_x": "id", "Scope_tmp": "Scope"})
            print("NewGroupDF\n", new_group_df)
            print("NewSiteDF\n", new_site_df)
            new_account_df = new_account_df.rename(columns={"name": "Scope"})
            print("NewAccountDF\n", new_account_df)

            return new_account_df, new_site_df, new_group_df

        # Levels = accounts, sites, groups
        def getPolicies(level, level_df):
            if level == "accounts":
                id_url = "/web/api/v2.1/accounts"
                fw_param = "accountIds"
            elif level == "sites":
                id_url = "/web/api/v2.1/sites"
                fw_param = "siteIds"
            elif level == "groups":
                id_url = "/web/api/v2.1/groups"
                fw_param = "groupIds"

            else:
                id_url = "/web/api/v2.1/accounts"

            if level == "global":
                policy_df = getGlobalPolicies()
                policy_df['Scope'] = "Global"
                policy_df['id'] = 0
                policy_df['level_id'] = 0

            else:
                level_df.to_csv("levels-" + level + "-" + domain + ".csv")

                level_id_list = level_df['id'].to_list()
                policy_df_list = []
                iterations_length = len(level_id_list)
                print(iterations_length)
                for level_id in level_id_list:
                    policies_url = id_url + "/" + str(level_id) + "/policy"
                    print(policies_url)
                    response = httpGet(policies_url)
                    if response.status_code == 200:
                        response_json = json.loads(response.text)
                        data = response_json['data']
                        data['level_id'] = level_id
                        fwparam = '?' + fw_param + "=" + level_id
                        response_fw = httpGet("/web/api/v2.1/firewall-control/configuration" + fwparam)
                        if response_fw.status_code == 200:
                            response_fwjson = json.loads(response_fw.text)
                            fw_data = response_fwjson['data']
                            data['fw_enabled'] = fw_data['enabled']
                            data['fw_inheritAllFirewallRules'] = fw_data['inheritAllFirewallRules']
                            data['fw_inheritedFrom'] = fw_data['inheritedFrom']
                            data['fw_inherits'] = fw_data['inherits']
                            data['fw_inheritSettings'] = fw_data['inheritSettings']
                            data['fw_locationAware'] = fw_data['locationAware']
                            # data['fw_reportBlocked'] = fw_data['reportBlocked']
                            data['fw_selectedTags'] = fw_data['selectedTags']
                        else:
                            data['fw_enabled'] = None
                            data['fw_inheritAllFirewallRules'] = None
                            data['fw_inheritedFrom'] = None
                            data['fw_inherits'] = None
                            data['fw_inheritSettings'] = None
                            data['fw_locationAware'] = None
                            # data['fw_reportBlocked'] = None
                            data['fw_selectedTags'] = None

                        response_dc = httpGet("/web/api/v2.1/device-control/configuration" + fwparam)
                        if response_dc.status_code == 200:
                            response_dcjson = json.loads(response_dc.text)
                            dc_data = response_dcjson['data']
                            data['dc_disableBleCommunication'] = dc_data['disableBleCommunication']
                            data['dc_disableRfcomm'] = dc_data['disableRfcomm']
                            data['dc_disallowAccessPermissionControl'] = dc_data['disallowAccessPermissionControl']
                            data['dc_enabled'] = dc_data['enabled']
                            data['dc_inheritedFrom'] = dc_data['inheritedFrom']
                            data['dc_inherits'] = dc_data['inherits']
                            data['dc_reportApproved'] = dc_data['reportApproved']
                            data['dc_reportBlocked'] = dc_data['reportBlocked']
                            data['dc_reportReadOnly'] = dc_data['reportReadOnly']
                        else:
                            data['dc_dc_disableBleCommunication'] = None
                            data['dc_disableRfcomm'] = None
                            data['dc_disallowAccessPermissionControl'] = None
                            data['dc_enabled'] = None
                            data['dc_inheritedFrom'] = None
                            data['dc_inherits'] = None
                            data['dc_reportBlocked'] = None
                            data['dc_reportReadOnly'] = None

                        policy_df_list.append(data)
                    else:
                        print("ERROR Status Code: " + str(response.status_code))
                        print(policies_url)
                        print(response.text)

                    iterations_length -= 1
                    print(level + " Policies left: " + str(iterations_length))
                policy_df = pd.DataFrame.from_records(policy_df_list)

            policy_df['level'] = level
            try:
                tmp_policy_df = pd.merge(policy_df, level_df, how='left', left_on='level_id', right_on='id')
                print(tmp_policy_df)
                policy_df = tmp_policy_df
            except:
                print("GlobalPolicy")

            print("Policy DF " + level + ": ")
            print(policy_df)
            return policy_df

        def getAllPolicies(level_account_df, level_site_df, level_group_df):
            group_policy_df = getPolicies("groups", level_group_df)
            site_policy_df = getPolicies("sites", level_site_df)
            account_policy_df = getPolicies("accounts", level_account_df)
            global_policy_df = getPolicies("global", None)
            policy_frames = [global_policy_df, account_policy_df, site_policy_df, group_policy_df]

            policy_df = pd.concat(policy_frames)
            del policy_df['level_id']
            print(policy_df)

            return (policy_df)

        def getAllLevels(ep, level_account_df, level_site_df, level_group_df):
            global_df = httpGetPagination(ep)
            global_df['level'] = 'Global'
            account_df = httpGetPaginationIds(ep, "account", level_account_df)
            site_df = httpGetPaginationIds(ep, "site", level_site_df)
            group_df = httpGetPaginationIds(ep, "groups", level_group_df)
            policy_frames = [global_df, account_df, site_df, group_df]

            policy_df = pd.concat(policy_frames)
            del policy_df['level_id']
            print(policy_df)
            return policy_df

        level_account_df, level_site_df, level_group_df = createLevelsDF()

        agents_df = httpGetPagination("/web/api/v2.1/agents")
        policy_df = getAllPolicies(level_account_df, level_site_df, level_group_df)
        users_df = httpGetPagination("/web/api/v2.1/users")
        exclusions_df = getAllLevels("/web/api/v2.1/exclusions", level_account_df, level_site_df, level_group_df)
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

        time.sleep(5)  # Example script

        # Destroy progress bar
        progress_bar.stop()
        progress_bar.destroy()

        # Display message box
        messagebox.showinfo("Script Status", "Script finished successfully")

gui_input = GUIInput()
exit()
