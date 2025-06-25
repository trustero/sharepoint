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

### Parameters

- `-SiteUrl`  
  The full URL of your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/yoursite`).

- `-AppId`  
  The Application (client) ID from your Azure AD App registration.

- `-folderPaths`  
  Comma-separated list of folder paths (relative to the SharePoint document library root) to assign permissions to.  
  Example: `"/Folder1,/Folder2"`

## What the Script Does

1. Connects to Microsoft Graph with the required scopes.
2. Finds the SharePoint site and drive.
3. Iterates through each folder path provided.
4. Assigns "read" permission to the specified Azure AD Application for each folder.
5. Outputs a summary table of the results.

## Example

```powershell
.\sharepoint-permission.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" -AppId "70d96320-e711-4e0e-94cf-53e43b557b0a" -folderPaths "/Invoices,/Reports"
```

## Troubleshooting

- Ensure you have the necessary permissions in Azure AD and SharePoint.
- Double-check folder paths for typos.
- If you see "Site not found" or "Drive not found", verify the `-SiteUrl` is correct
