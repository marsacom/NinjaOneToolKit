#!/user/bin/env python3

# NinjaOneToolKit
# Perform a variety of functions related to devices in NinjaOne/Domain
# Gather values from XLSX Sheet and compare to NinjaOne/Domain
# Update values in XLSX Sheet to reflect devices in/not in NinjaOne/Domain
# List devices in and not in NinjaOne/Domain
# Add devices into NinjaOne/Domain
# Brayden Kukla - 2024


import os
import sys
import csv
import warnings
import requests
import openpyxl as xl
import subprocess
from dotenv import load_dotenv


endpoint = "https://app.ninjarmm.com/v2/"
oauth_url = "https://app.ninjarmm.com/ws/oauth/token"
api_token = 'token'


# Be sure to change the path in .env 
load_dotenv() 
path = os.getenv('XL_PATH')

# This is specifically to ignore the unknown warning that is generated when accessing the worksheet via openpyxl
warnings.simplefilter('ignore')

wb = xl.load_workbook(path)
ws = wb['Computers']

 
def start():
    getToken()
    # getExcelData()


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

    global api_token
    api_token = token["access_token"]

    #print("API Token: " + token["access_token"] + "\n")


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

    print('-'*60)
    print("\nOrganizations\n")
    print('-'*60, '\n')
    for i in organizations:
        print(i["id"], i["name"])
        org.append(i["name"])
        org_id.append(i["id"])

    print('\n')

    getDevices(token)


# Gather devices based on users organization selection
def getDevices(token):
    user_sel = input("Please select an organization... ")

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

    print('-'*60)
    print("\nDevices in NinjaOne...\n")
    print('-'*60, '\n')

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
            print(f"{'-'*12 : <15}{'-'*6 : ^10}{'-'*8 : >10}")
            print(f"{ninja_systemNames[l] : <15}{ninja_ids[l] : ^10}{status : >10} \n")
            l = l + 1
    else:
        print("\nThere are no devices currently associated with this organization...\n")
        sys.exit()


# Parse info from computers.csv to be able to compare in a later function
def checkCSV():
    file = "computers.csv"

    global ad_rows
    global ad_dns
    global ad_ip
    global ad_name

    ad_rows = []
    ad_dns = []
    ad_ip = []
    ad_name = []

    with open(file, 'r') as csvfile:
        reader = csv.reader(csvfile)
        
        for row in reader:
            ad_rows.append(row)

        # This just deletes the first 2 items in the list to get rid of the bullshit info we dont want
        for i in range(2):
            ad_rows.pop(0)

        l = 0

        print('-'*60)
        print("\nDevices in the Domain...\n")
        print('-'*60, '\n')

        for row in ad_rows:
            ad_name.append(row[4])
            ad_dns.append(row[1])

            # Some of the IPs are unknown in the domain for some reason, this is just to check if 
            if row[3] == '':
                ad_ip.append('   UNKNOWN  ')
            else:
                ad_ip.append(row[3])

            print(f"{'System Name' : <20}{'IP' : ^15}{'DNS Name' : >20}")
            print(f"{'-'*12 : <20}{'-'*12 : ^15}{'-'*18 : >25}")
            print(f"{ad_name[l] : <20}{ad_ip[l] : ^2}{ad_dns[l] : >28} \n")
            l = l + 1


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


# Compare results of devices in NinjaOne to the Excel File and update values in the "Computers Sheet"
def compareRes():
    ninja_missing = []
    ad_missing = []
    both = []

    print('-'*60)
    print("\nComparing results... \n")

    for i in range(len(xl_systemNames)):
        if xl_systemNames[i] not in ad_name:
            print("Device: ", xl_systemNames[i], " has NOT yet joined the Domain...")
            ad_missing.append(xl_systemNames[i])
            ws['D'+str(xl_rowNum[i])] == 'N'

            if xl_systemNames[i] not in ninja_systemNames:
                print("Device: ", xl_systemNames[i], " has NOT been registered into NinjaOne")
                ninja_missing.append((xl_systemNames[i]))
                ws['E'+str(xl_rowNum[i])] = 'N'

            else:
                print("Device: ", xl_systemNames[i], " has already been registered in NinjaOne")
                ws['E'+str(xl_rowNum[i])] = 'Y'

        else:
            print("Device: ", xl_systemNames[i], " has already joined the Domain...")
            ws['D'+str(xl_rowNum[i])] == 'Y'

            if xl_systemNames[i] not in ninja_systemNames:
                print("Device: ", xl_systemNames[i], " has NOT been registered into NinjaOne")
                ninja_missing.append(xl_systemNames[i])
                ws['E'+str(xl_rowNum[i])] = 'N'

            else:
                print("Device: ", xl_systemNames[i], " has already been registered in NinjaOne")
                ws['E'+str(xl_rowNum[i])] = 'Y'
    
    # Which devices are missing from NinjaOne & Domain
    for d in range(len(ninja_missing)):
        if ninja_missing[d] not in ad_missing:
            pass
        else:
            both.append(ninja_missing[d])

    # Attempt to write results to results.txt file
    try:
        with open(os.getenv('LOG_PATH'), "w") as f:
            f.write("NinjaOne\n" + "------------\n")
            for i in range(len(ninja_missing)):
                f.write("Device: " + ninja_missing[i] + " has NOT yet joined NinjaOne... \n")
            f.write("\nDomain\n" + "--------\n")
            for d in range(len(ad_missing)):
                f.write("Device: " + ad_missing[d] + " has NOT yet joined the Domain... \n")
            f.write("\nNinjaOne & Domain\n" + "----------------\n")
            for k in range(len(both)):
                f.write("Device: " + both[k] + " has NOT yet joined the Domain or NinjaOne... \n")
        print('-'*60)
        print("\nSUCCESS: Results have been saved in " + os.getenv('LOG_PATH') + '...\n')    
    except:
        print("\nERROR: File not found. Please ensure the path for logs is set correctly in the .env file...\n")

    wb.save(path)


# Main
def main():
    start()   
    print("\nStarting CompTools v.1.0...")
    print('-'*60)
    print(" 1: List all devices in NinjaOne\n", "2: List all devices in the Domain\n", "3: List devices missing from NinjaOne and the Domain\n", "4: Add computer to the Domain\n", "5: Add computer to NinjaOne\n")

    choice = int(input("Please select an option from the list above (1-5)... "))
    
    if choice == 1: # List all devices in NinjaOne
        getOrgs(api_token)
    elif choice == 2: # List all devices in the Domain
        cmd = " Get-ADComputer -Filter * -Properties IPv4Address | Export-Csv C:\\Users\\bkukla\\VSCode\\NinjaOneXLSXCompare\\computers.csv"
        p = subprocess.Popen('powershell -command' + cmd)
        p.communicate()
        checkCSV()
    elif choice == 3: # List devices missing from NinjaOne and the Domain
        getOrgs(api_token)
        cmd = " Get-ADComputer -Filter * -Properties IPv4Address | Export-Csv C:\\Users\\bkukla\\VSCode\\NinjaOneXLSXCompare\\computers.csv"
        p = subprocess.Popen('powershell -command' + cmd)
        p.communicate()
        getExcelData()    
        checkCSV()
        compareRes()
    elif choice == 4: # Add computer to the Domain
        cmd = " Add-Computer -DomainName ssis.local -Credential ssis\\administrator -Restart"
        p = subprocess.Popen('powershell -command' + cmd)
        p.communicate()
    elif choice == 5: # Add computer to NinjaOne
        pass
    else:
        print("\nERROR: Please re-run the script and enter a valid value, 1-5")


if __name__=="__main__":
    main()
else:
    print("ERROR: Unknown error occurred, exiting...\n")
    sys.exit()