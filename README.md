# NinjaOneToolKit

A CLI Application for IT Technicians to access and manage Organizations & Devices in NinjaOne(NinjaRMM) straight from the Command Line.

- Can additionally be used to compare devices from an XLSX sheet and Active Directory to devices registered in NinjaOne


## Current Features
> 1. List all devices in NinjaOne
> 2. List all devices in the Domain 
> 3. List devices in Ninja & AD and compare with XLSX file
> 4. Generate XLSX file of device in Ninja
> 5. ~~Add computer to NinjaOne~~  ***WIP*** 


## Installation

Step 1. ``git clone https://github.com/marsacom/NinjaOneToolKit.git``

Step 2. ``pip3 install -r requirements.txt``


## Prerequisites
This script currently uses **.env** files for storing API Keys and other configurable variables that are used by the script in various places.

Please refer to the **.env** file for configuring these various variables.

- Future updates will most likely transition away from the current **.env** approach to opt for a more user friendly **.ini** file


## Usage

``python3 main.py``
- Follow prompts in CLI


## Notes

***This script is currently in early stages of development and is being actively updated.*** 

**CURRENTLY** this script will require some customization/tweaking to work out of the box for your organization but if you are here looking to use this tool I am hopefull you can figure it out ;)

**FUTURE** updates will turn this script into a full CLI Application for...
> - Accessing & managing devices in NinjaOne.
>   - Performing a variety of functions related to managing and gathering device information.
> - More advanced, in depth comparison reports between Ninja, AD, amd a given list of devices.
> - Support more customizations.
>   - Turn the script from a tailored solution for 1 organization into a tool that can be used out of the box by anyone for any organization.

The goal of this script is to be a one stop shop for device info gathering & management of devices by IT Technicians via a single CLI Application 


Author : Brayden Kukla - 2024
