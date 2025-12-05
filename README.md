# SharePoint Folder Permission Assignment Script

This PowerShell script assigns permissions to specific folders in a SharePoint site for a registered Azure AD Application.

## Prerequisites

- PowerShell 7+
- [Microsoft.Graph PowerShell module](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation)
- Permissions to connect to Microsoft Graph with `Sites.ReadWrite.All` and `Files.ReadWrite.All` scopes
- Application (client) ID from Azure AD App Registration

## Installation

1. Install the Microsoft Graph module if not already installed:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. Save the script as `sharepoint-permission.ps1`.

## Usage


Run the script in PowerShell with the required parameters:

```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -AppId "YOUR-APP-ID" -folderPaths "/Shared Documents/Folder1,/Shared Documents/Folder2"
```

You can also specify a custom drive name (optional, defaults to "Documents"):

```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -AppId "YOUR-APP-ID" -folderPaths "/Finance/Invoices,/Finance/Reports" -DriveName "FinanceDocs"
```

### Parameters

### Script Options

- `-SiteUrl` (required)
  - The full URL of your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/yoursite`).

- `-AppId` (required)
  - The Application (client) ID from your Azure AD App registration.

- `-folderPaths` (required)
  - Comma-separated list of folder paths (relative to the SharePoint document library root) to assign permissions to.
  - Example: `"/Folder1,/Folder2"`

- `-DriveName` (optional)
  - The name of the SharePoint drive (document library). Defaults to `Documents` if not supplied.
  - Example: `-DriveName "FinanceDocs"`

## What the Script Does

1. Connects to Microsoft Graph with the required scopes.
2. Finds the SharePoint site and drive (document library).
3. Iterates through each folder path provided (comma-separated).
4. Assigns "read" permission to the specified Azure AD Application for each folder.
5. Handles errors for missing sites, drives, or folders and outputs a summary table of the results.

## Example


### Example 1: Basic Usage
```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" -AppId "70d96320-e711-4e0e-94cf-53e43b557b0a" -folderPaths "/Invoices,/Reports"
```

### Example 2: Custom Drive Name
```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/hr" -AppId "70d96320-e711-4e0e-94cf-53e43b557b0a" -folderPaths "/Policies,/Benefits" -DriveName "HRDocs"
```

### Example 3: Multiple Folders and Error Handling
```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projects" -AppId "70d96320-e711-4e0e-94cf-53e43b557b0a" -folderPaths "/2025/Alpha,/2025/Beta,/2025/NonExistentFolder"
```
If a folder does not exist, the script will report the error and continue processing the remaining folders.

## Troubleshooting


- Ensure you have the necessary permissions in Azure AD and SharePoint.
- Double-check folder paths for typos and correct casing.
- If you see "Site not found" or "Drive not found", verify the `-SiteUrl` and `-DriveName` are correct.
- If a folder is not found, check the path and ensure it exists in the specified drive.
