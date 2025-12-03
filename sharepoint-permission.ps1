param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SharePoint site URL (e.g., https://yourtenant.sharepoint.com/sites/yoursite)")]
    [string]$SiteUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the Application ID (e.g., 70d96320-e711-4e0e-94cf-53e43b557b0a) from the Azure AD App registration.")]
    [string]$AppId,
    [Parameter(Mandatory = $true, HelpMessage = "Enter list of folder paths to give permission separated by comma")]
    [string]$folderPaths,
    [Parameter(Mandatory = $false, HelpMessage = "Enter the Drive name (e.g., Documents). Defaults to 'Documents' if not supplied.")]
    [string]$DriveName = "Documents"
)

Import-Module Microsoft.Graph.Sites

Connect-MgGraph -Scopes "Sites.ReadWrite.All", "Files.ReadWrite.All"

function Get-DriveByName {
    param (
        [string]$SiteId,
        [string]$DriveName
    )
    $drives = Get-MgSiteDrive -SiteId $SiteId
    $drive = $drives | Where-Object { $_.Name -eq $DriveName }
    return $drive
}

# Get the Site ID from the provided URL
$site = Get-MgSite -Search $siteUrl
if (-not $site) {
    Write-Host "Site not found. Please check the URL and try again." -ForegroundColor Red
    exit
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
        $permissionParams = @{
            roles     = @("read")
            grantedTo = @{
                application = @{
                    id = $appId
                }
            }
        }
        $permission = New-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $item.Id -BodyParameter $permissionParams
        $itemsResult += [PSCustomObject]@{
            FolderPath      = $folderPath
            PermissionId    = $permission.Id
            PermissionRoles = $permission.Roles -join ', '
            PermissionAdded = $true
        }
    }
    catch {
        Write-Host "Error retrieving folder: $folderPath. $_" -ForegroundColor Red
        $itemsResult += [PSCustomObject]@{
            FolderPath      = $folderPath
            PermissionId    = $null
            PermissionRoles = $null
            PermissionAdded = $false
        }
    }
}
Write-Host "Permissions added for the following folders:" -ForegroundColor Green
$itemsResult | Format-Table -AutoSize | Out-Host
Write-Host "Script completed." -ForegroundColor Green