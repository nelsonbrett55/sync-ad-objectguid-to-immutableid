# ActiveDirectory to EntraID (formally AzureAD) ImmutableID Migrator

A GUI-based PowerShell tool for syncing and verifying on-premises Active Directory users with their Azure Active Directory counterparts by comparing and updating their ImmutableID based on the AD objectGUID.

---

## 🔍 Features

- 🎛️ Windows Forms GUI interface for easy interaction
- 🔄 Pulls and compares on-prem AD and Azure AD user data
- 🧮 Auto-generates the expected ImmutableID from on-prem GUID
- 🟢 Color-coded display indicating sync status
- 🔁 Supports refreshing, sorting, and zooming
- 🔐 Optional: Set or update ImmutableID (functionality stubbed, can be extended)
- 📊 Live progress bar and status updates

---

## 📷 Screenshot
![Screenshot](https://github.com/nelsonbrett55/sync-ad-objectguid-to-immutableid/blob/main/Screenshot.png)

---

## 🛠 Requirements

- Windows PowerShell
- Admin privileges (for AD queries and module installations)
- Modules:
  - [`MSOnline`](https://learn.microsoft.com/en-us/powershell/module/msonline/?view=azureadps-1.0)
  - `ActiveDirectory` PowerShell module

---

## 🚀 Getting Started

1. Clone this repo:
   ```powershell
   git clone https://github.com/nelsonbrett55/sync-ad-objectguid-to-immutableid.git
   cd AzureAD-SID-to-ImmutableID-Migrator
