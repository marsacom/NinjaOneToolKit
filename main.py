# Compare devices in NinjaOne organization to data in XLSX Sheet
# Call NinjaOne API for organization and device information 
# Gather values from XLSX Sheet and compare to NinjaOne
# Update values in XLSX Sheet to reflect devices in/not in NinjaOne
# Brayden Kukla - 2024


import os
import sys
import warnings
import requests
import openpyxl as xl
from dotenv import load_dotenv


endpoint = "https://app.ninjarmm.com/v2/"
oauth_url = "https://app.ninjarmm.com/ws/oauth/token"
api_token = ''

# Be sure to change the path if running script on different machine
load_dotenv() 
path = os.getenv('XL_PATH')

warnings.simplefilter('ignore')
wb = xl.load_workbook(path)
ws = wb['Computers']


# Call api endpoint for bearer token, currently this is just uses a machine-to-machine application using client credentials
def getToken():
    data = {
        "grant_type": "client_credentials",
        "client_id": str(os.getenv('CLIENT_ID')),
        "client_secret": str(os.getenv('CLIENT_SECRET')),
        "scope": "monitoring"
    }

    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    token = requests.post(oauth_url, data, headers).json()

    api_token = token["access_token"]
    print("API Token: " + token["access_token"] + "\n")

    getOrgs(api_token)


# Get organizations assocaited in NinjaOne
def getOrgs(token):
    org_url = endpoint + "organizations/"

    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer " + token,
    }

    organizations = requests.get(org_url, headers=headers).json()

    org = []
    org_id = []

    print("Organizations...\n")
    for i in organizations:
        print(i["id"], i["name"])
        org.append(i["name"])
        org_id.append(i["id"])

    print('\n')

    getDevices(token)


# Gather devices based on users organization selection
def getDevices(token):
    user_sel = input("Please select an organization...")

    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer " + token,
    }

    device_url = endpoint + "/organization/" + user_sel + "/devices/"
    devices = requests.get(device_url, headers=headers).json()

    global ninja_ids
    global ninja_systemNames 
    global ninja_status

    ninja_ids = []
    ninja_systemNames = []
    ninja_status = []

    l = 0

    print("\nDevices in NinjaOne...\n")

    if len(devices) >= 1:
        for k in devices:
            ninja_ids.append(int(k["id"]))
            ninja_systemNames.append(str(k["systemName"]))
            ninja_status.append(str(k["offline"]))

            if ninja_status[l] == 'False':
                status = 'Online'
            else:
                status = "Offline"

            print(f"{'System Name' : <15}{'ID' : ^10}{'Status' : >10}")
            print(f"{'-------------' : <15}{'------' : ^10}{'-------' : >10}")
            print(f"{ninja_systemNames[l] : <15}{ninja_ids[l] : ^10}{status : >10} \n")
            l = l + 1
    else:
        print("\nThere are no devices currently associated with this organization...\n")
        sys.exit()

# Load excel sheet and gather device info
def getExcelData():
    global xl_ids
    global xl_systemNames 
    global xl_rowNum
    global xl_ninja_statuses

    xl_ids = []
    xl_systemNames = []
    xl_rowNum = []
    xl_ninja_statuses = []

    l = 1

    for row in ws.iter_rows(min_row=2, max_row=80, values_only=True):
        if row[0] == None:
            pass
        else:
            l = l + 1
            xl_ids.append(row[1])
            xl_systemNames.append(row[0])
            xl_rowNum.append(l)
            xl_ninja_statuses.append(row[4])

    compareRes()


# Compare results of devices in NinjaOne to the Excel File and update values
def compareRes():
    missing = []

    print("\nComparing results...\n")

    for i in range(len(xl_systemNames)):
        if xl_systemNames[i] not in ninja_systemNames:
            print("Device: ", xl_systemNames[i], " is missing from NinjaOne...\n")
            missing.append([xl_systemNames[i]])
            ws['E'+str(xl_rowNum[i])] = 'N'
        else:
            print("Device: ", xl_systemNames[i], " is already registered in NinjaOne...\n")
            ws['E'+str(xl_rowNum[i])] = 'Y'

    wb.save(path)


# Main
def main():   
    getToken()
    getExcelData()


if __name__=="__main__":
    main()
else:
    sys.exit()













