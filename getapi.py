import pandas as pd
import requests
import json
import time

class APIClient:
    def __init__(self, token, domain, user):
        self.token_header = 'APIToken ' + token
        self.real_user = user
        self.customer_endpoint = "https://" + domain
        self.domain = domain

    def getDomain(self):
        return self.domain

    def httpGet(self, endpoint):
        url = self.customer_endpoint + endpoint
        headers={'Authorization': self.token_header}
        resp = requests.get(url, headers=headers)
        print(resp)
        print(resp.text)
        return resp

    def httpGetPagination(self, endpoint):
        query_params = "?limit=100"
        nextPage = True
        df_list = []
        FirstRun = True
        sleepcount = 0
        while nextPage:
            print(endpoint+query_params)
            response_json = json.loads(self.httpGet(endpoint+query_params).text)
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
            print("Items Remaining: ",total_left)
        raw_df = pd.concat(df_list)
        print(endpoint+"\n", raw_df)
        return raw_df

    def httpGetPaginationIds(self, endpoint, level, level_id_df):
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
                response_json = json.loads(self.httpGet(endpoint + query_params).text)
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
        print(endpoint + "\n", raw_df)
        return raw_df

    def httpGetPaginationIds(self, endpoint, level, level_id_df):
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
                response_json = json.loads(self.httpGet(endpoint + query_params).text)
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
            tmp_policy_df = pd.merge(raw_df, self.level_df, how='left', left_on='level_id', right_on='id')
            print(tmp_policy_df)
            raw_df = tmp_policy_df
        except:
            print("GlobalPolicy")

        print(endpoint + "\n", raw_df)
        return raw_df

    def getGlobalPolicies(self):
        tmp_global_list = []
        endpoint = "/web/api/v2.1/tenant/policy"
        query_params = "?limit=100"
        response_json = json.loads(self.httpGet(endpoint + query_params).text)
        data = response_json['data']
        response_fw = self.httpGet("/web/api/v2.1/firewall-control/configuration")

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

        response_dc = self.httpGet("/web/api/v2.1/device-control/configuration")
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

    def createLevelsDF(self):
        id_url = "/web/api/v2.1/accounts"
        account_df = self.httpGetPagination(id_url)
        id_url = "/web/api/v2.1/sites"
        query_params = "?limit=100"
        nextPage = True
        df_list = []
        while nextPage:
            print(id_url + query_params)
            response_json = json.loads(self.httpGet(id_url + query_params).text)
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
        group_df = self.httpGetPagination(id_url)
        id_url = "/web/api/v2.1/accounts"
        account_df = self.httpGetPagination(id_url)

        id_url = "/web/api/v2.1/sites"
        query_params = "?limit=100"
        nextPage = True
        df_list = []
        sleepcount = 0
        while nextPage:
            print(id_url + query_params)
            response_json = json.loads(self.httpGet(id_url + query_params).text)
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
    def getPolicies(self, level, level_df):
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
            policy_df = self.getGlobalPolicies()
            policy_df['Scope'] = "Global"
            policy_df['id'] = 0
            policy_df['level_id'] = 0
        else:
            level_df.to_csv("levels-" + level + "-" + self.getDomain() + ".csv")

            level_id_list = level_df['id'].to_list()
            policy_df_list = []
            iterations_length = len(level_id_list)
            print(iterations_length)
            for level_id in level_id_list:
                policies_url = id_url + "/" + str(level_id) + "/policy"
                print(policies_url)
                response = self.httpGet(policies_url)
                if response.status_code == 200:
                    response_json = json.loads(response.text)
                    data = response_json['data']
                    data['level_id'] = level_id
                    fwparam = '?' + fw_param + "=" + level_id
                    response_fw = self.httpGet("/web/api/v2.1/firewall-control/configuration" + fwparam)
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

                    response_dc = self.httpGet("/web/api/v2.1/device-control/configuration" + fwparam)
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

    def getAllPolicies(self, level_account_df, level_site_df, level_group_df):
        group_policy_df = self.getPolicies("groups", level_group_df)
        site_policy_df = self.getPolicies("sites", level_site_df)
        account_policy_df = self.getPolicies("accounts", level_account_df)
        global_policy_df = self.getPolicies("global", None)
        policy_frames = [global_policy_df, account_policy_df, site_policy_df, group_policy_df]

        policy_df = pd.concat(policy_frames)
        del policy_df['level_id']
        print(policy_df)

        return (policy_df)

    def getAllLevels(self, ep, level_account_df, level_site_df, level_group_df):
        global_df = self.httpGetPagination(ep)
        global_df['level'] = 'Global'
        account_df = self.httpGetPaginationIds(ep, "account", level_account_df)
        site_df = self.httpGetPaginationIds(ep, "site", level_site_df)
        group_df = self.httpGetPaginationIds(ep, "groups", level_group_df)
        policy_frames = [global_df, account_df, site_df, group_df]

        policy_df = pd.concat(policy_frames)
        del policy_df['level_id']
        print(policy_df)
        return policy_df
