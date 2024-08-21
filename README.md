# SNMP Printer Management Script

This PowerShell script is designed to retrieve SNMP data from networked printers, compile usage statistics (black/white and color prints), and export this data into a CSV file. It supports a variety of printer models and can be easily extended to include more devices.

## Prerequisites

- **Windows Environment**: This script is designed for Windows and requires PowerShell 5.1 or newer.
- **SNMP Service**: Ensure SNMP services are enabled and accessible on the printers and the machine running the script.
- **PowerShell SNMP Module**: This script uses the PowerShell SNMP module. You must have this module installed, which can be done via PowerShellGet or manually if not available by default.

## Usage

1. **Download the Script**: Clone or download the script files to your local machine or network accessible to the printers.
2. **Configure SNMP on Printers**: Ensure that SNMP is enabled on each printer you wish to monitor and that you have the correct community string.
3. **Edit the Script**: Update the `$devices` array with the printers' details (see the section on adding new printers).
4. **Run the Script**: Open PowerShell, navigate to the directory containing the script, and run it using the command:
   ```powersershell
   .\PrinterSNMPManagement.ps1


### Notes on the README

- **Clear and Concise**: Provides all necessary information without overwhelming the user.
- **Structured**: Organized into sections that a user might refer to independently.
- **Actionable Steps**: Each section contains actionable and specific steps for performing tasks related to the script.

Make sure to update the README with any specific paths, dependency details, or internal procedures relevant to your environment or organization. Adjust the troubleshooting section based on common issues observed or anticipated with the deployment of the script.

