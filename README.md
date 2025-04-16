# VM Migration Automation Script

This repository contains a PowerShell automation script for handling virtual machine operations (e.g., migration, configuration, processing) using an Excel-based input file.

## 📁 Files in this Repository

- **partbypartversion2.ps1**  
  Main PowerShell script to automate operations for each VM listed in the Excel sheet.

- **Input_VMList.xlsx**  
  Input file containing details of virtual machines. This file must be edited with valid data before running the script.

---

## ✅ Prerequisites

- Windows OS with **PowerShell 5.1+** or **PowerShell Core**
- Excel installed, or the `ImportExcel` PowerShell module
- Execution policy set to allow script execution:
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser


📝 Editing the Input Excel File (Input_VMList.xlsx)
Before running the script, update the following mandatory columns in the Excel sheet:

Column Name	Description	Example Value
VMIPAddress	IP Address of the virtual machine	192.168.1.10
VMOSType	OS Type (e.g., Windows, Linux)	Windows or Linux
VMUsername	Username used to access the VM	adminuser
VMPassword	Password for the VM	P@ssword123
Make sure to populate these fields with accurate credentials for each VM. Leave no rows blank for active VMs.

▶️ Running the Script
Once the Excel file is updated:

Open PowerShell.

Navigate to the directory where your files are stored.

Run the script using the following command:

powershell
Copy
Edit
.\partbypartversion2.ps1
The script will process each row from Input_VMList.xlsx one by one and perform the defined automation based on the VM details.

🛠️ Customization
You can modify the PowerShell script to:

Add logging functionality

Validate VM credentials before proceeding

Filter VMs based on tags or OS types

Handle errors more robustly

🧾 Example
Here's how your Excel sheet should look before execution:

VMIPAddress	VMOSType	VMUsername	VMPassword
10.10.233.18	Linux	ubuntu	mylinuxpwd
10.10.233.81	Windows	admin	Passw0rd!
10.10.233.27	Linux	root	R00t@123
📌 Notes
Ensure the Excel file is not open when running the script.

Make a backup of your Excel file before bulk changes.

Sensitive credentials should be secured or encrypted for production environments.

📧 Contact
For any queries or improvements, feel free to raise an issue or submit a pull request.

yaml
Copy
Edit

---

Let me know if you'd like me to auto-generate sample data into your Excel file or provide a CSV version for easier version control on GitHub.







