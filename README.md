```markdown
# AppRiver (SecureTide) to Microsoft 365 Migration Tool

A PowerShell automation script to migrate Allow/Block lists (Senders, Domains, and IPs) from AppRiver SecureTide into Microsoft 365 Defender.

## 🚀 Why this exists
Migrating manual filter settings from AppRiver to Microsoft 365 is often painful because:
1.  **Mismatched Policies:** AppRiver puts everything in one place; Microsoft 365 requires Senders/Domains in the *Content Filter* but IPs in the *Connection Filter*.
2.  **IP Subnet Errors:** Microsoft 365 rejects any IP range larger than `/24` (256 addresses). AppRiver allows larger ranges (e.g., `/20`), causing import scripts to fail.
3.  **Data Cleanup:** AppRiver exports often contain messy headers or duplicate entries.

This script solves those problems automatically.

## ✨ Features
* **Auto-Detection:** Automatically finds CSV files in the script directory.
* **Smart IP Splitting:** Detects IP ranges that are too large for M365 (e.g., `96.43.144.0/20`) and automatically splits them into valid `/24` subnets during import.
* **Policy Routing:** intelligently routes:
    * Emails/Domains -> `HostedContentFilterPolicy`
    * IPs -> `HostedConnectionFilterPolicy`
* **Safety Checks:** Warns you if your IP list exceeds the Microsoft 1275 entry limit.
* **Non-Destructive:** Appends to your existing Microsoft 365 lists rather than overwriting them.

## 📋 Prerequisites
* PowerShell 5.1 or PowerShell 7+
* **Exchange Online PowerShell Module** installed:
  `Install-Module -Name ExchangeOnlineManagement`
* Global Administrator or Exchange Administrator permissions.

## ⚙️ Setup & Usage

### 1. Export Data from AppRiver
Log in to the AppRiver Admin Center and export your filters. If there is no export button, copy the tables into Excel and save them as **CSV (Comma delimited)**.

### 2. Prepare the Files
Rename your CSV files to match these exact filenames and place them in the same folder as the script:

| Filename | Required Headers | Description |
| :--- | :--- | :--- |
| `FilteredEmailAddresses.csv` | `Email`, `Type` | Contains email addresses. Type column must contain "Allowed" or "Blocked". |
| `FilteredDomains.csv` | `Domain`, `Type` | Contains domains (e.g., gmail.com). Type column must contain "Allowed" or "Blocked". |
| `FilteredIPs.csv` | `Ip Addresses`, `Type` | Contains IPs or CIDR ranges. Type column must contain "Allowed". |

### 3. Run the Script
1. Open PowerShell as Administrator.
2. Navigate to the folder containing the script and CSVs.
3. Run the script:
   `.\Migrate-AppRiver.ps1`
4. Sign in to your Microsoft 365 account when prompted.

## ⚠️ Known Limitations
* **IP Limit:** Microsoft 365 allows a maximum of **1275 entries** in the Connection Filter. If splitting a large AppRiver range results in exceeding this limit, the script will warn you.
* **Massive Networks:** The script skips network ranges larger than `/16` (65,536 addresses) to prevent flooding your policy with thousands of entries. These should be handled via Transport Rules instead.

## 📄 License
MIT License - feel free to modify and use this in your organization.

## 🤝 Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.