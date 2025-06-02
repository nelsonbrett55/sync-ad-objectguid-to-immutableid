# Active Directory to Entra ID (formerly Azure AD) ImmutableID Migrator

A GUI-based PowerShell tool for syncing on-premises Active Directory users with their Entra ID (formerly Azure AD) counterparts by comparing and updating their `ImmutableID` based on the on-premises `objectGUID`. This ensures that when you migrate and sync, you avoid duplicated or mismatched users in Microsoft 365.

---

## 🔍 Features

- 🎛️ Intuitive Windows Forms GUI for ease of use  
- 🔄 Automatically pulls and compares on-prem AD and Entra ID user data  
- 🧮 Generates expected `ImmutableID` from the on-prem `objectGUID`  
- 🟢 Color-coded display showing match/mismatch status  
- 🔁 Supports sorting, zooming, and refreshing  
- 🔐 Placeholder for setting/updating ImmutableID (stubbed, can be extended)  
- 📊 Live progress bar and status updates during sync  

---

## 📷 Screenshot

![Screenshot](https://github.com/nelsonbrett55/sync-ad-objectguid-to-immutableid/blob/main/Screenshot.png)

---

## 🛠 Requirements

- Windows PowerShell (not PowerShell Core)
- Local admin privileges
- Required PowerShell modules:
  - [`MSOnline`](https://learn.microsoft.com/en-us/powershell/module/msonline/?view=azureadps-1.0)  
  - `ActiveDirectory` module (available via RSAT or Windows Server)

> 💡 Ensure you're running PowerShell as Administrator and that the RSAT tools are installed to access Active Directory cmdlets.

---

## 🚀 Getting Started

1. **Clone this repository:**
   ```powershell
   git clone https://github.com/nelsonbrett55/sync-ad-objectguid-to-immutableid.git
   cd sync-ad-objectguid-to-immutableid
