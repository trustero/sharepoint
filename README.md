# SharePoint Access Setup for Trustero

One PowerShell script that gives a Trustero app registration read access to
specific SharePoint folders. It replaces two manual steps in the Azure portal
and one in SharePoint:

| Step | Portal / manual way | With the script |
| --- | --- | --- |
| Add the `Sites.Selected` and `Files.SelectedOperations.Selected` application permissions | Azure portal → App registrations → API permissions | `-SetupAppPermissions` |
| Grant admin consent | Azure portal → "Grant admin consent for <tenant>" | `-SetupAppPermissions` |
| Grant the app read access to each folder | No portal UI; Graph API call per folder | `-SiteUrl` + `-folderPaths` |

The app registration itself (name, client secret, tenant ID) is still created
in the Azure portal. See the
[App Registration Setup Guide](https://support.trustero.com/knowledge/microsoft-app-registration-setup-guide)
and the
[SharePoint Document Management System Setup Guide](https://support.trustero.com/knowledge/sharepoint-document-management-system-setup-guide).

## Prerequisites

- PowerShell 7 or later (Windows, macOS or Linux).
- The Microsoft Graph PowerShell modules `Microsoft.Graph.Authentication`,
  `Microsoft.Graph.Sites`, `Microsoft.Graph.Files` and, for
  `-SetupAppPermissions`, `Microsoft.Graph.Applications`. The script installs
  any that are missing into the current user's profile; no admin rights on the
  machine are needed. To preinstall them yourself:

  ```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files, Microsoft.Graph.Applications -Scope CurrentUser
  ```

- The Application (client) ID of the Trustero app registration.

## Who runs which part

The script has two independent phases, each with its own Microsoft sign-in
that asks only for the scopes that phase needs.

| Phase | Switch | Who | Graph scopes requested at sign-in |
| --- | --- | --- | --- |
| A. App permissions + admin consent | `-SetupAppPermissions` | Global Administrator or Privileged Role Administrator | `Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All` |
| B. Folder grants | `-SiteUrl` + `-folderPaths` | Anyone with access to the SharePoint site | `Sites.ReadWrite.All`, `Files.ReadWrite.All` |

Phase A runs once per app registration. Phase B runs whenever folders are
added. One person with both roles can run both in a single command.

> Cloud Application Administrator is **not** enough for Phase A. Microsoft
> does not let that role consent to Microsoft Graph application permissions.

## Usage

Phase A, Global Administrator, once:

```powershell
./sharepoint-permission.ps1 -AppId <client-id> -SetupAppPermissions
```

Phase B, site owner, per folder set:

```powershell
./sharepoint-permission.ps1 -AppId <client-id> -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -folderPaths '/Policies,/Evidence'
```

Both phases in one run:

```powershell
./sharepoint-permission.ps1 -AppId <client-id> -SetupAppPermissions -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -folderPaths '/Policies,/Evidence'
```

A document library other than the default `Documents`:

```powershell
./sharepoint-permission.ps1 -AppId <client-id> -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -folderPaths '/Invoices' -DriveName 'FinanceDocs'
```

Each command opens a browser window for Microsoft sign-in. Run the script
again with more folders at any time; grants that already exist are skipped.

## Parameters

| Parameter | Required | Description |
| --- | --- | --- |
| `-AppId` | Always | Application (client) ID of the Trustero app registration. |
| `-SetupAppPermissions` | Phase A | Add `Sites.Selected` and `Files.SelectedOperations.Selected` as application permissions and grant admin consent. Existing permissions on the registration (for example the delegated ones used by the file picker) are kept. |
| `-SiteUrl` | Phase B | Full SharePoint site URL, e.g. `https://contoso.sharepoint.com/sites/finance`. |
| `-folderPaths` | Phase B | Comma-separated folder paths relative to the document library root, e.g. `'/Policies,/Evidence/2026'`. Wrap in single quotes; folder names with spaces are fine. |
| `-DriveName` | No | Document library name. Defaults to `Documents`. |

## What the script does

Phase A (`-SetupAppPermissions`):

1. Signs in with `Application.ReadWrite.All` and `AppRoleAssignment.ReadWrite.All`.
2. Finds the app registration by client ID and the Microsoft Graph service
   principal in your tenant.
3. Adds `Sites.Selected` and `Files.SelectedOperations.Selected` to the
   registration's requested permissions, merging with whatever is already
   there.
4. Creates the enterprise application (service principal) if the app has
   never been consented, then grants admin consent by assigning the two app
   roles to it.
5. Signs out.

Phase B (`-SiteUrl` + `-folderPaths`):

1. Signs in with `Sites.ReadWrite.All` and `Files.ReadWrite.All`.
2. Resolves the site from its URL and the document library from `-DriveName`.
3. For each folder, grants the app registration the `read` role on that
   folder.
4. Prints a summary table and exits non-zero if any folder was not granted.

Both phases are idempotent: rerunning skips anything already in place.

## Troubleshooting

- **`The term 'Get-MgSite' is not recognized`**: the Graph modules are not
  installed. The current script installs them automatically; with an older
  copy run the `Install-Module` line under Prerequisites.
- **`Site not found for '<url>'`**: check the URL. The script resolves the
  site by host and path, so `https://<tenant>.sharepoint.com/sites/<site>`
  must be exact, without a trailing page or library path.
- **`Drive '<name>' not found`**: the document library name differs from
  `Documents`. Pass `-DriveName` with the library's display name.
- **`Error retrieving folder`**: the path is relative to the library root and
  case-sensitive. `/Shared Documents/...` is wrong when the library is already
  selected by `-DriveName`; use `/Folder` instead.
- **`Insufficient privileges` during `-SetupAppPermissions`**: the signed-in
  user is not a Global Administrator or Privileged Role Administrator.
- **Trustero gets 403 on a file**: its folder was not granted. Rerun Phase B
  with that folder, or a parent folder, in `-folderPaths`.
