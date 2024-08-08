# Network PC Info Retriever

## Description
This PowerShell script creates a graphical user interface (GUI) to retrieve information about PCs in a network. It uses Active Directory and WMI to gather PC names and hardware information and saves the results to specified files.

## Features
- GUI with a form for easy interaction
- Retrieve PC names from Active Directory
- Retrieve detailed hardware information of PCs using WMI
- Save results to specified directories
- Open the results files directly from the GUI

## Prerequisites
- Windows operating system
- PowerShell installed
- Active Directory module for PowerShell
- Necessary permissions to query Active Directory and WMI

## Setup Instructions
1. **Install the Active Directory Module:**
   - If you don't have the Active Directory module installed, you can install it via PowerShell:
     ```powershell
     Install-WindowsFeature -Name RSAT-AD-PowerShell
     ```

2. **Place the Script:**
   - Place the `Network_PC_Info_Retriever.ps1` script in your desired directory.

## Usage Instructions
1. **Open PowerShell with Administrative Privileges:**
   - Right-click on the PowerShell icon and select `Run as administrator`.

2. **Navigate to the Script Directory:**
   ```powershell
   cd path\to\your\script

3. **Run the Script:**
   ```powershell
   .\Network_PC_Info_Retriever.ps1

4. **Interact with the GUI:**
- **Retrieve PC Names:** Click the "Retrieve PC Names" button to get a list of PC names from Active Directory.
- **Retrieve PC Hardware:** Click the "Retrieve PC Hardware" button to gather detailed hardware information from the PCs.
- **Open Names File:** Click the "Open Names File" button to open the directory containing the saved PC names.
- **Open Hardware File:** Click the "Open Hardware File" button to open the directory containing the saved hardware information.

  ## Directory Structure
- **Input Directory:** Stores the PC names file.
- Default: **'C:\Active Directory Reports\Input'**
- **Output Directory:** Stores the PC hardware information file.
- Default: **'C:\Active Directory Reports\Output'**

  ## Notes
- Modify the directory paths in the script if necessary.
- Ensure you have the necessary permissions to query Active Directory and WMI.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

