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
from datetime import datetime

endpoint = "https://app.ninjarmm.com/v2/"
oauth_url = "https://app.ninjarmm.com/ws/oauth/token"
api_token = 'token'

pwrshl = 'powershell -command'


# Be sure to change the path in .env 
load_dotenv() 
path = os.getenv('XL_PATH')

# This is specifically to ignore the random warning that is generated when accessing the worksheet via openpyxl (does not affect the script)
warnings.simplefilter('ignore')

wb = xl.load_workbook(path)
ws = wb['Computers']

 
def start():
    get_token()
    get_excel_data()


# Call api endpoint for bearer token, currently this is just uses a machine-to-machine application using client credentials
def get_token():
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


# Get organizations assocaited in NinjaOne
def get_orgs(token):
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
    get_devices(token)


# Gather devices based on users organization selection
def get_devices(token):
    user_sel = input("Please select an organization... ")

    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer " + token,
    }

    device_url = endpoint + "/organization/" + user_sel + "/devices/"
    devices = requests.get(device_url, headers=headers).json()

    global ninja_ids
    global ninja_system_names 
    global ninja_status

    ninja_ids = []
    ninja_system_names = []
    ninja_status = []

    l = 0

    print('-'*60)
    print("\nDevices in NinjaOne...\n")
    print('-'*60, '\n')

    if len(devices) >= 1:
        for k in devices:
            ninja_ids.append(int(k["id"]))
            ninja_system_names.append(str(k["systemName"]))
            ninja_status.append(str(k["offline"]))

            if ninja_status[l] == 'False':
                status = 'Online'
            else:
                status = "Offline"

            print(f"{'System Name' : <15}{'ID' : ^10}{'Status' : >10}")
            print(f"{'-'*12 : <15}{'-'*6 : ^10}{'-'*8 : >10}")
            print(f"{ninja_system_names[l] : <15}{ninja_ids[l] : ^10}{status : >10} \n")
            l = l + 1
    else:
        print("\nThere are no devices currently associated with this organization...\n")
        sys.exit()


# Parse info from computers.csv to be able to compare in a later function
def check_csv():
    file = os.getenv('CSV_PATH')

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
        for _ in range(2):
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
def get_excel_data():
    global xl_ids
    global xl_system_names 
    global xl_rowNum
    global xl_ninja_statuses

    xl_ids = []
    xl_system_names = []
    xl_rowNum = []
    xl_ninja_statuses = []

    l = 1

    for row in ws.iter_rows(min_row=2, max_row=80, values_only=True):
        if row[0] != None:
            l = l + 1
            xl_ids.append(row[1])
            xl_system_names.append(row[0])
            xl_rowNum.append(l)
            xl_ninja_statuses.append(row[4])


def in_ninja(device):
    if device in ninja_system_names:
        return True
    else:
        return False

def in_domain(device):
    if device in ad_name:
        return True
    else:
        return False
       

# Compare results of devices in NinjaOne to the Excel File and update values in the "Computers Sheet"
def compare_res():
    ninja_missing = []
    ad_missing = []
    both = []

    global dev_lbl
    dev_lbl = "Device: "
    dom_yes = " : Domain - YES"
    dom_no = " : Domain - NO";

    nin_yes = " : NinjaOne - YES"
    nin_no = " : NinjaOne - NO"

    print('-'*60)
    print("\nComparing results... \n")

    for i in range(len(xl_system_names)):
        if in_domain(xl_system_names[i]) == False:
            print(dev_lbl, xl_system_names[i], dom_no)
            ad_missing.append(xl_system_names[i])
            ws['D'+str(xl_rowNum[i])] = 'N'

            if in_ninja(xl_system_names[i]) == False:
                print(dev_lbl, xl_system_names[i], nin_no)
                ninja_missing.append((xl_system_names[i]))
                ws['E'+str(xl_rowNum[i])] = 'N'

            else:
                print(dev_lbl, xl_system_names[i], nin_yes)
                ws['E'+str(xl_rowNum[i])] = 'Y'

        else:
            print(dev_lbl, xl_system_names[i], dom_yes)
            ws['D'+str(xl_rowNum[i])] = 'Y'

            if in_ninja(xl_system_names[i]) == False:
                print(dev_lbl, xl_system_names[i], nin_no)
                ninja_missing.append(xl_system_names[i])
                ws['E'+str(xl_rowNum[i])] = 'N'

            else:
                print(dev_lbl, xl_system_names[i], nin_yes)
                ws['E'+str(xl_rowNum[i])] = 'Y'
        print('-'*60)
    
    # Which devices are missing from NinjaOne & Domain
    for d in range(len(ninja_missing)):
        if ninja_missing[d] in ad_missing:
            both.append(ninja_missing[d])

    write_to_file(ninja_missing, ad_missing, both)


# Add a computer to the domain remotely
def add_to_domain():
    name = input("Please enter the name of the PC you wish to join to the domain... ex. WKSSSISXX-XX : ")

    # Using same creds for local and domain admin as 
    cmd = " Add-Computer -ComputerName " + str(name) + " -LocalCredential ssis\\administrator -DomainName ssis.local -Credential ssis\\administrator -Restart"
    p = subprocess.Popen(pwrshl + cmd)
    p.communicate()


# Get all computers associated with Active Directory
def get_ad_computers():
    cmd = " Get-ADComputer -Filter * -Properties IPv4Address | Export-Csv C:\\Users\\bkukla\\VSCode\\NinjaOneToolKit\\computers.csv"
    p = subprocess.Popen(pwrshl + cmd)
    p.communicate()


# def add_to_ninja():
#     cmd = " "
#     p = subprocess.Popen(pwrshl + cmd)
#     p.communicate


def write_to_file(ninja_missing, ad_missing, both):
    # Write results to results.txt file
    try:    
        with open(os.getenv('LOG_PATH'), "w") as f:
            f.write("NinjaOne\n" + "------------\n")
            for i in range(len(ninja_missing)):
                f.write(dev_lbl + ninja_missing[i] + " has NOT yet joined NinjaOne... \n")
            f.write("\nDomain\n" + "--------\n")
            for d in range(len(ad_missing)):
                f.write(dev_lbl + ad_missing[d] + " has NOT yet joined the Domain... \n")
            f.write("\nNinjaOne & Domain\n" + "----------------\n")
            for k in range(len(both)):
                f.write(dev_lbl + both[k] + " has NOT yet joined the Domain or NinjaOne... \n")
            f.write("\nSUCCESS: Script completed at - " + str(datetime.now()))

            print('-'*60)
            print("\nSUCCESS: Results have been saved in " + os.getenv('LOG_PATH') + '...\n')    
    except FileNotFoundError:
        print("\nERROR: File not found. Please ensure the path for logs is set correctly in the .env file...\n")


# Main
def main():
    start()   
    print("\nStarting NinjaOneToolKit v.1.1...")
    print('-'*60)
    print(" 1: List all devices in NinjaOne\n", "2: List all devices in the Domain\n", "3: List devices missing from NinjaOne and the Domain\n", "4: Add computer to the Domain\n", "5: Add computer to NinjaOne\n")

    choice = int(input("Please select an option from the list above (1-5)... "))
    
    if choice == 1: # List all devices in NinjaOne
        get_orgs(api_token)
    elif choice == 2: # List all devices in the Domain
        get_ad_computers()
        check_csv()
    elif choice == 3: # List devices missing from NinjaOne and the Domain
        get_ad_computers()
        get_orgs(api_token)   
        check_csv()
        compare_res()
    elif choice == 4: # Add computer to the Domain
        add_to_domain()
    elif choice == 5: # Add computer to NinjaOne
        pass
    else:
        print("\nERROR: Please re-run the script and enter a valid value, 1-5")


# https://app.ninjarmm.com/agent/installer/09e02ed8-75fc-4456-86eb-1c8d14444a63/ssindustrialsurplusmainoffice-5.8.9154-windows-installer.msi


if __name__=="__main__":
    main()
else:
    print("ERROR: Unknown error occurred, exiting...\n")
    sys.exit()