# DAG-SAR
Scripts here helps in analyzing the JSON file as part of your SharePoint Online Site Access Review to analyze the group permissions and remediate oversharing.

# SharePoint Permission Management Toolkit

This repository contains a set of PowerShell scripts to **analyze**, **summarize**, and **remove** permissions from SharePoint Online sites. These scripts are built to work with **PnP PowerShell** and **Microsoft Graph** modules, supporting modern authentication and robust permission reporting. The toolkit is ideal for IT admins, site collection administrators, and SharePoint security reviewers.

---

## Table of Contents

- [Overview](#overview)
- [Scripts](#scripts)
  - [1. SARPermissionSummary.ps1](#1-sarpermissionsummaryps1)
  - [2. SARPermission1.ps1 (Main Site)](#2-sarpermission1ps1-main-site)
  - [3. SARPermission1subweb.ps1 (Subwebs)](#3-sarpermission1subwebps1-subwebs)
- [Usage Examples](#usage-examples)
- [Authentication](#authentication)
- [Requirements](#requirements)
- [Best Practices & Notes](#best-practices--notes)

---

## Overview

These scripts help you:

- **Summarize** permissions across multiple site collections.
- **Remove** unwanted group or "Everyone" permissions from the main SharePoint site.
- **Remove** permissions from subwebs (subsites) with granular control.
- **Export** sharing summaries for review and compliance.

---

## Scripts

### 1. SARPermissionSummary.ps1

**Purpose:**  
Generate a summary CSV showing which items (Web, File, List, Folder) in each site collection are shared and with whom.

**How it works:**

- Reads a CSV list of site collection URLs.
- Connects to each site via PnP PowerShell with Azure AD App credentials.
- Finds the latest permission report file in the `DO_NOT_DELETE_REVIEW_INSTANCE` list.
- Parses the file (supports JSON and JSONL formats).
- Outputs a CSV showing each item and its sharing (groups, Everyone, Everyone except external users).

**Sample command:**
```powershell
.\SARPermissionSummary.ps1 -CsvPath "C:\Sites.csv" -ClientId "<client-id>"
```

---

### 2. SARPermission1.ps1 (Main Site)

**Purpose:**  
Remove or simulate removal of permissions for the **main SharePoint site** only.  
**Identifies subweb items** and reports them for separate processing.

**How it works:**

- Connects to the main site with the authentication method you specify.
- Retrieves the latest permission report from the specified list.
- Processes only items belonging to the main site (not subwebs).
- Removes permissions for selected groups, "Everyone", or "Everyone except external users".
- Reports subweb items to be handled separately.

**Sample command:**
```powershell
.\SARPermission1.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MainSite" -RemoveEveryone -WhatIf -ClientId "<client-id>"
```

---

### 3. SARPermission1subweb.ps1 (Subwebs)

**Purpose:**  
Remove or simulate removal of permissions for a **specific subweb/subsite**.

**How it works:**

- Connects to the main site to retrieve the permission report.
- Connects to the target subweb.
- Processes and removes permissions for items belonging to that subweb, according to your filters (group, "Everyone", etc.).
- Supports WhatIf/dry-run mode.

**Sample command:**
```powershell
.\SARPermission1subweb.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MainSite/Subweb" -MainSiteUrl "https://tenant.sharepoint.com/sites/MainSite" -RemoveEveryone -WhatIf -ClientId "<client-id>"
```

---

## Usage Examples

#### Summarize Permissions for Many Sites

```powershell
.\SARPermissionSummary.ps1 `
  -CsvPath "C:\Scripts\SiteInput.csv" `
  -ClientId "<your-client-id>"
```

#### Remove "Everyone" from Main Site (Simulation)

```powershell
.\SARPermission1.ps1 `
  -SiteUrl "https://tenant.sharepoint.com/sites/MainSite" `
  -RemoveEveryone `
  -WhatIf `
  -ClientId "<your-client-id>"
```

#### Remove Specific Groups from Subweb (Simulation)

```powershell
.\SARPermission1subweb.ps1 `
  -SiteUrl "https://tenant.sharepoint.com/sites/MainSite/Subweb" `
  -MainSiteUrl "https://tenant.sharepoint.com/sites/MainSite" `
  -GroupsToRemove @("Visitors", "Members") `
  -WhatIf `
  -ClientId "<your-client-id>"
```

---

## Authentication

All scripts support modern authentication using Azure AD App registrations with appropriate permissions.  
**Supported methods:**

- **Interactive** (recommended for most users)
- **Client Secret** (service principal)
- **Certificate** (service principal with certificate)
- **Device Login**

> See parameter descriptions in the script headers for details.

---

## Requirements

- **PowerShell 7.0+** (recommended)
- **Modules:**
  - [PnP.PowerShell](https://pnp.github.io/powershell/)
  - [Microsoft.Graph.Authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview)
  - [Microsoft.Graph.Users]
  - [Microsoft.Graph.Groups]
- **Azure AD App** with permissions:
  - Microsoft Graph: `Sites.Read.All` (minimum), `Sites.Manage.All` (to modify)
  - SharePoint: `Sites.Read.All`
- **SharePoint Online**

---

## Best Practices & Notes

- **Always run with `-WhatIf` first** to simulate and review changes before applying.
- The main site script will not modify subwebs—it will only report them. Use the subweb script for those.
- Ensure the site/subweb URLs you provide match exactly as they appear in the permission report.
- Permission report files must be present in the `DO_NOT_DELETE_REVIEW_INSTANCE` list.
- **Logs:** Each script writes a detailed log file to the working directory.
- **Output:** The summary script produces a CSV suitable for Excel review or audit.
- Scripts are designed for **review, compliance, and risk reduction**—test on a non-production tenant first.

---

## Maintainer

Chintan Pandya
https://github.com/chpandya_microsoft
```
If you have questions or need help, please open an issue in this repository.
```
