[CmdletBinding(DefaultParameterSetName = 'Folders')]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Application ID (e.g., 70d96320-e711-4e0e-94cf-53e43b557b0a) from the Azure AD App registration.")]
    [string]$AppId,

    # Folder grants (Phase B). Mandatory on a plain run; optional when
    # -SetupAppPermissions is given, so an admin can do the permission half alone.
    [Parameter(Mandatory = $true, ParameterSetName = 'Folders', HelpMessage = "Enter the SharePoint site URL (e.g., https://yourtenant.sharepoint.com/sites/yoursite)")]
    [Parameter(Mandatory = $false, ParameterSetName = 'AppPermissions')]
    [string]$SiteUrl,
    [Parameter(Mandatory = $true, ParameterSetName = 'Folders', HelpMessage = "Enter list of folder paths to give permission separated by comma")]
    [Parameter(Mandatory = $false, ParameterSetName = 'AppPermissions')]
    [string]$folderPaths,
    [Parameter(Mandatory = $false, HelpMessage = "Enter the Drive name (e.g., Documents). Defaults to 'Documents' if not supplied.")]
    [string]$DriveName = "Documents",

    # App permissions (Phase A). Needs only -AppId. Requires Global Administrator
    # or Privileged Role Administrator.
    [Parameter(Mandatory = $true, ParameterSetName = 'AppPermissions', HelpMessage = "Add the Sites.Selected and Files.SelectedOperations.Selected application permissions to the app registration and grant admin consent.")]
    [switch]$SetupAppPermissions
)

$ErrorActionPreference = 'Stop'

# Get-MgSite / Get-MgSiteDrive live in Microsoft.Graph.Sites; Get-MgDriveItem /
# New-MgDriveItemPermission live in Microsoft.Graph.Files. Install what is
# missing instead of failing later with "Get-MgSite is not recognized".
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites', 'Microsoft.Graph.Files')
if ($SetupAppPermissions) { $requiredModules += 'Microsoft.Graph.Applications' }
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing PowerShell module $m (CurrentUser scope)..." -ForegroundColor Yellow
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m
}

# ---------------------------------------------------------------------------
# Phase A (optional, -SetupAppPermissions): what the Azure portal steps
# "API permissions > Add a permission" and "Grant admin consent" do by hand.
# Runs with its own Graph sign-in and only the directory scopes it needs, so
# a site owner running Phase B alone is never asked to consent to
# Application.ReadWrite.All. The signed-in user must be a Global
# Administrator or Privileged Role Administrator: Cloud Application
# Administrator cannot consent to Microsoft Graph application permissions.
# ---------------------------------------------------------------------------
$GraphResourceAppId = '00000003-0000-0000-c000-000000000000'
$RequiredAppRoles   = @('Sites.Selected', 'Files.SelectedOperations.Selected')

function Set-TrusteroAppPermissions {
    param([string]$AppId)

    Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -NoWelcome

    $app = Get-MgApplication -Filter "appId eq '$AppId'"
    if (-not $app) { throw "No app registration with Application (client) ID $AppId in this tenant." }

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphResourceAppId'"
    if (-not $graphSp) { throw "Microsoft Graph service principal not found in this tenant." }

    $roles = @()
    foreach ($value in $RequiredAppRoles) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' }
        if (-not $role) { throw "Microsoft Graph application permission '$value' is not available in this tenant." }
        $roles += $role
    }

    # Step 1: request the permissions on the registration (portal: "Add a permission").
    # Merge into the existing list; the setup guide also has customers adding
    # delegated Sites.Read.All / User.Read.All for the picker, which an
    # overwrite would silently drop.
    $existing = @($app.RequiredResourceAccess)
    $graphEntry = $existing | Where-Object { $_.ResourceAppId -eq $GraphResourceAppId } | Select-Object -First 1
    $otherEntries = @($existing | Where-Object { $_.ResourceAppId -ne $GraphResourceAppId })
    $access = @()
    if ($graphEntry) { $access = @($graphEntry.ResourceAccess) }
    $added = @()
    foreach ($role in $roles) {
        if (-not ($access | Where-Object { $_.Id -eq $role.Id -and $_.Type -eq 'Role' })) {
            $access += @{ Id = $role.Id; Type = 'Role' }
            $added += $role.Value
        }
    }
    if ($added.Count -gt 0) {
        $newGraphEntry = @{ ResourceAppId = $GraphResourceAppId; ResourceAccess = $access }
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess ($otherEntries + @($newGraphEntry))
        Write-Host "Requested application permissions on '$($app.DisplayName)': $($added -join ', ')" -ForegroundColor Cyan
    }
    else {
        Write-Host "Application permissions already requested on '$($app.DisplayName)'." -ForegroundColor Cyan
    }

    # Step 2: grant admin consent (portal: "Grant admin consent for <tenant>").
    # Consent for application permissions is an app role assignment on the
    # app's own service principal, pointing at the Graph service principal.
    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'"
    if (-not $sp) {
        Write-Host "Creating enterprise application (service principal) for $AppId..." -ForegroundColor Yellow
        $sp = New-MgServicePrincipal -AppId $AppId
    }
    $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id)
    foreach ($role in $roles) {
        if ($assignments | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $graphSp.Id }) {
            Write-Host "Admin consent already granted for $($role.Value)." -ForegroundColor Cyan
            continue
        }
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $role.Id | Out-Null
        Write-Host "Admin consent granted for $($role.Value)." -ForegroundColor Green
    }

    Disconnect-MgGraph | Out-Null
}

if ($SetupAppPermissions) {
    Set-TrusteroAppPermissions -AppId $AppId
    if (-not $folderPaths) {
        Write-Host "App permissions configured. Run the script again with -SiteUrl and -folderPaths (as a site member) to grant folders." -ForegroundColor Green
        exit 0
    }
    if (-not $SiteUrl) {
        Write-Host "-folderPaths requires -SiteUrl." -ForegroundColor Red
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Phase B: grant the app read access to specific folders. The signed-in user
# needs access to the site, not a directory role.
# ---------------------------------------------------------------------------
Connect-MgGraph -Scopes "Sites.ReadWrite.All", "Files.ReadWrite.All" -NoWelcome

function Get-DriveByName {
    param (
        [string]$SiteId,
        [string]$DriveName
    )
    $drives = Get-MgSiteDrive -SiteId $SiteId
    $drive = $drives | Where-Object { $_.Name -eq $DriveName }
    return $drive
}

# Resolve the site by hostname + server-relative path (GET /sites/{host}:{path}).
# This is exact; -Search is full-text and often returns nothing for a URL.
try {
    $uri = [Uri]$SiteUrl
    $site = Get-MgSite -SiteId "$($uri.Host):$($uri.AbsolutePath.TrimEnd('/'))"
}
catch {
    Write-Host "Site not found for '$SiteUrl'. Please check the URL and try again. $_" -ForegroundColor Red
    exit 1
}
$siteId = $site.Id
Write-Host "Sharepoint Site Found" -ForegroundColor Cyan -nonewline
$site | Format-Table -AutoSize | Out-Host

# Get the Drive by name from the Site ID
$drive = Get-DriveByName -SiteId $siteId -DriveName $DriveName
if (-not $drive) {
    Write-Host "Drive '$DriveName' not found. Please check the Drive name and try again." -ForegroundColor Red
    exit
}
Write-Host "Sharepoint Drive '$DriveName' found" -ForegroundColor Cyan -nonewline
$drive | Format-Table -AutoSize | Out-Host

# Split the folder paths into an array
$folderPathArray = $folderPaths -split ',' | ForEach-Object { $_.Trim() }
$itemsResult = @()

foreach ($folderPath in $folderPathArray) {
    try {
        $item = Get-MgDriveItem -DriveId $drive.Id -DriveItemId "root:$folderPath"
        Write-Host "Folder retrieved successfully: $folderPath" -ForegroundColor Cyan
        $item | Format-Table -AutoSize | Out-Host
        # Microsoft documents grantedToV2 as the only accepted grantee shape on
        # driveItem permissions (grantedTo / grantedToIdentities are deprecated).
        $permissionParams = @{
            roles       = @("read")
            grantedToV2 = @{
                application = @{
                    id = $appId
                }
            }
        }
        $permission = New-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $item.Id -BodyParameter $permissionParams
        $grantedApp = $permission.GrantedToV2.Application.Id
        if (-not $grantedApp) { $grantedApp = $permission.GrantedTo.Application.Id }
        if ($grantedApp -ne $appId) {
            Write-Host "Warning: permission $($permission.Id) was created but Graph reports grantee '$grantedApp', not $appId. Verify with Get-MgDriveItemPermission." -ForegroundColor Yellow
        }
        $itemsResult += [PSCustomObject]@{
            FolderPath      = $folderPath
            PermissionId    = $permission.Id
            PermissionRoles = $permission.Roles -join ', '
            GrantedToAppId  = $grantedApp
            PermissionAdded = $true
        }
    }
    catch {
        Write-Host "Error retrieving folder: $folderPath. $_" -ForegroundColor Red
        $itemsResult += [PSCustomObject]@{
            FolderPath      = $folderPath
            PermissionId    = $null
            PermissionRoles = $null
            GrantedToAppId  = $null
            PermissionAdded = $false
        }
    }
}
Write-Host "Permissions added for the following folders:" -ForegroundColor Green
$itemsResult | Format-Table -AutoSize | Out-Host
$failed = @($itemsResult | Where-Object { -not $_.PermissionAdded })
if ($failed.Count -gt 0) {
    Write-Host "Script completed with $($failed.Count) folder(s) NOT granted. Trustero will get 403 on files under those folders." -ForegroundColor Red
    exit 1
}
Write-Host "Script completed." -ForegroundColor Green
