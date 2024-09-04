# NinjaOneToolKit

A CLI Application for IT Technicians to access and manage Organizations & Devices in NinjaOne(NinjaRMM) straight from the Command Line.

- Can additionally be used to compare devices from an XLSX sheet and Active Directory to devices registered in NinjaOne

## Current Features
> 1. List all devices in NinjaOne
> 2. List all devices in the Domain 
> 3. List devices missing from NinjaOne & Domain
> 4. Add computer to Domain
> 5. ~~Add computer to NinjaOne~~  ***WIP*** 

## Installation

Step 1. ``git clone https://github.com/marsacom/NinjaOneToolKit.git``

Step 2. ``pip3 install -r requirements.txt``

## Usage

``python3 main.py``
- Follow prompts in CLI

## Notes

***This script is currently in early stages of development and is being actively updated.*** 

**CURRENTLY** this script will require some customization/tweaking to work out of the box for your organization but if you are here looking to use this tool I am hopefull you can figure it out ;)

**FUTURE** updates will turn this script into a full CLI Application for...
> - Accessing & managing devices in NinjaOne.
>   - Performing a variety of functions related to managing and gathering device information.
> - Compare XLSX spreadsheet to computers in NinjaOne *and/or* Active Directory.
> - Support more customizations.
>   - Turn the script from a tailored solution for 1 organization into a tool that can be used out of the box by anyone for any organization.

The goal of this script is to be a one stop shop for device info gathering & management of devices by IT Technicians via a single CLI Application 


Author : Brayden Kukla - 2024
