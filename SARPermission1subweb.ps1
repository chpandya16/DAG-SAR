[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$MainSiteUrl,

    [Parameter(Mandatory=$false)]
    [string]$ConfigFileName = "",

    [Parameter(Mandatory=$false)]
    [string]$ConfigListName = "DO_NOT_DELETE_REVIEW_INSTANCE",

    [Parameter(Mandatory=$false)]
    [string]$LogPath = ".\Subweb-PermissionRemovalLog.txt",

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,

    [Parameter(Mandatory=$false)]
    [string[]]$GroupsToRemove = @(),

    [Parameter(Mandatory=$false)]
    [ValidateSet("Web", "File", "List", "Folder")]
    [string[]]$ItemTypesToProcess = @("Web", "File", "List", "Folder"),

    [Parameter(Mandatory=$false)]
    [switch]$RemoveEveryoneExceptExternal = $false,

    [Parameter(Mandatory=$false)]
    [switch]$RemoveEveryone = $false,

    [Parameter(Mandatory=$false)]
    [switch]$ProcessOnlyUniquePermissions = $true,

    # Authentication Parameters
    [Parameter(Mandatory=$false)]
    [string]$ClientId,

    [Parameter(Mandatory=$false)]
    [string]$ClientIdForSecret,
    [Parameter(Mandatory=$false)]
    [SecureString]$ClientSecret,
    [Parameter(Mandatory=$false)]
    [string]$TenantIdForSecret,

    [Parameter(Mandatory=$false)]
    [string]$ClientIdForCert,
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint = "",
    [Parameter(Mandatory=$false)]
    [string]$CertificatePath = "",
    [Parameter(Mandatory=$false)]
    [SecureString]$CertificatePassword = $null,
    [Parameter(Mandatory=$false)]
    [string]$TenantIdForCert,

    [Parameter(Mandatory=$false)]
    [string]$ClientIdForDevice
)

$Script:Config = @{
    RequiredModules = @("PnP.PowerShell", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups")
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    if ($Message -eq "") {
        Write-Host ""
        if ($LogPath) {
            Add-Content -Path $LogPath -Value "" -ErrorAction Stop
        }
        return
    }
    $logMessage = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        "DEBUG"   { "Cyan" }
        default   { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
    if ($LogPath) {
        try {
            $logDir = Split-Path -Path $LogPath -Parent
            if ($logDir -and -not (Test-Path $logDir)) {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }
}

function Initialize-RequiredModules {
    foreach ($moduleName in $Script:Config.RequiredModules) {
        if (!(Get-Module -ListAvailable -Name $moduleName)) {
            Write-Log "Required module '$moduleName' is not installed" "ERROR"
            throw "Missing required module: $moduleName"
        }
        Import-Module $moduleName -Force -ErrorAction Stop
        Write-Log "Imported module: $moduleName" "SUCCESS"
    }
}

function Get-AuthMethod {
    if ($ClientIdForSecret) { return 'ClientSecret' }
    elseif ($ClientIdForCert) { return 'Certificate' }
    elseif ($ClientIdForDevice) { return 'DeviceLogin' }
    elseif ($ClientId) { return 'Interactive' }
    else { throw "No valid authentication parameters provided." }
}

function Connect-SharePointSite {
    param(
        [string]$SiteUrl,
        [string]$AuthenticationMethod,
        [string]$ClientId,
        [SecureString]$ClientSecret,
        [string]$CertificateThumbprint,
        [string]$CertificatePath,
        [SecureString]$CertificatePassword,
        [string]$TenantId
    )
    try {
        Write-Log "Connecting to SharePoint using $AuthenticationMethod method" "INFO"
        switch ($AuthenticationMethod) {
            'Interactive' {
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ForceAuthentication -ErrorAction Stop
            }
            'ClientSecret' {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientIdForSecret -ClientSecret $ClientSecret -Tenant $TenantIdForSecret -ErrorAction Stop
            }
            'Certificate' {
                if ($CertificateThumbprint) {
                    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientIdForCert -Thumbprint $CertificateThumbprint -Tenant $TenantIdForCert -ErrorAction Stop
                }
                elseif ($CertificatePath) {
                    if ($CertificatePassword) {
                        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientIdForCert -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Tenant $TenantIdForCert -ErrorAction Stop
                    }
                    else {
                        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientIdForCert -CertificatePath $CertificatePath -Tenant $TenantIdForCert -ErrorAction Stop
                    }
                }
            }
            'DeviceLogin' {
                Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientIdForDevice -ErrorAction Stop
            }
        }
        $web = Get-PnPWeb -ErrorAction Stop
        Write-Log "Successfully connected to: $($web.Title) ($($web.Url))" "SUCCESS"
        return $web
    }
    catch {
        Write-Log "Failed to connect to SharePoint: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Get-ConfigurationFromMainSite {
    param(
        [string]$MainSiteUrl,
        [string]$ListName,
        [string]$ConfigFileName,
        [string]$AuthenticationMethod
    )
    $tempConfigPath = $null
    $ConfigFileUrl = $null
    try {
        Write-Log "Connecting to main site for configuration: $MainSiteUrl" "INFO"
        $authParams = @{
            SiteUrl = $MainSiteUrl
            AuthenticationMethod = $AuthenticationMethod
            ClientId = $ClientId
            ClientIdForSecret = $ClientIdForSecret
            ClientSecret = $ClientSecret
            TenantIdForSecret = $TenantIdForSecret
            ClientIdForCert = $ClientIdForCert
            CertificateThumbprint = $CertificateThumbprint
            CertificatePath = $CertificatePath
            CertificatePassword = $CertificatePassword
            TenantIdForCert = $TenantIdForCert
            ClientIdForDevice = $ClientIdForDevice
        }
        Connect-SharePointSite @authParams | Out-Null

        $list = Get-PnPList -Identity $ListName -ErrorAction Stop
        Write-Log "Found configuration list: $($list.Title)" "SUCCESS"
        if (-not $ConfigFileName) {
            $camlQuery = @"
<View>
  <Query>
    <Where>
      <IsNotNull>
        <FieldRef Name='FileLeafRef'/>
      </IsNotNull>
    </Where>
    <OrderBy>
      <FieldRef Name='Created' Ascending='False'/>
    </OrderBy>
  </Query>
  <RowLimit>5</RowLimit>
</View>
"@
            $files = Get-PnPListItem -List $ListName -Query $camlQuery -ErrorAction Stop
            foreach ($file in $files) {
                $fileName = $file.FieldValues.FileLeafRef
                $fileUrl = $file.FieldValues.FileRef
                $fileSize = if ($file.FieldValues.File_x0020_Size) { $file.FieldValues.File_x0020_Size } else { 0 }
                if ($fileSize -gt 100) {
                    try {
                        $testContent = Get-PnPFile -Url $fileUrl -AsString -ErrorAction Stop
                        if ($testContent.Contains('"ItemType"') -or $testContent.Contains('"TenantId"')) {
                            $ConfigFileName = $fileName
                            $ConfigFileUrl = $fileUrl
                            Write-Log "Selected latest permission report: $fileName" "SUCCESS"
                            break
                        }
                    }
                    catch { continue }
                }
            }
        }
        if (-not $ConfigFileName -or -not $ConfigFileUrl) {
            throw "No valid permission report file found in main site"
        }
        $tempConfigPath = Join-Path $env:TEMP "subweb-config-$(Get-Date -Format 'yyyyMMddHHmmss')"
        $configContent = Get-PnPFile -Url $ConfigFileUrl -AsString -ErrorAction Stop
        Set-Content -Path $tempConfigPath -Value $configContent -Encoding UTF8
        Write-Log "Downloaded configuration file: $ConfigFileName" "SUCCESS"
        $permissionData = @()
        $lines = ($configContent -split "`r?`n") | Where-Object { $_.Trim() -ne "" }
        if ($lines.Count -gt 1 -and $lines[0].Trim().StartsWith('{')) {
            foreach ($line in $lines) {
                try {
                    $jsonObject = $line.Trim() | ConvertFrom-Json
                    if ($jsonObject.ItemType -or $jsonObject.TenantId) {
                        $permissionData += $jsonObject
                    }
                }
                catch { Write-Log "Failed to parse config line" "WARNING" }
            }
        }
        else {
            $parsedData = $configContent | ConvertFrom-Json
            $permissionData = if ($parsedData -is [array]) { $parsedData } else { @($parsedData) }
        }
        Write-Log "Loaded $($permissionData.Count) permission records from main site" "SUCCESS"
        return $permissionData
    }
    finally {
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
        if ($tempConfigPath -and (Test-Path $tempConfigPath)) {
            Remove-Item $tempConfigPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SubwebItems {
    param(
        [array]$AllPermissionData,
        [string]$CurrentSiteUrl
    )
    $subwebItems = @()
    foreach ($item in $AllPermissionData) {
        $parentWeb = if ($item.ParentWeb) { $item.ParentWeb } else { "" }
        if ($parentWeb -eq $CurrentSiteUrl) {
            $subwebItems += $item
        }
    }
    Write-Log "Found $($subwebItems.Count) items for current subweb: $CurrentSiteUrl" "INFO"
    return $subwebItems
}

function New-SubwebActions {
    param(
        [array]$SubwebItems,
        [hashtable]$FilterCriteria
    )
    $actions = @()
    foreach ($item in $SubwebItems) {
        if ($FilterCriteria.ItemTypes -and $item.ItemType -notin $FilterCriteria.ItemTypes) { continue }
        $peopleCount = if ($item.PeopleCount) { [int]$item.PeopleCount } else { 0 }
        $groupsToRemove = @()
        if ($FilterCriteria.GroupsToRemove.Count -gt 0) {
            foreach ($group in $item.GroupDetails) {
                if ($group.Name -in $FilterCriteria.GroupsToRemove) {
                    $groupsToRemove += $group.Name
                }
            }
        }
        $hasEveryone = if ($item.HasEveryone) { [bool]$item.HasEveryone } else { $false }
        $hasEveryoneExceptExternal = if ($item.HasEveryoneExceptExternalUser) { [bool]$item.HasEveryoneExceptExternalUser } else { $false }
        $shouldProcess = ($groupsToRemove.Count -gt 0) -or
                         ($FilterCriteria.RemoveEveryone -and $hasEveryone) -or
                         ($FilterCriteria.RemoveEveryoneExceptExternal -and $hasEveryoneExceptExternal)
        if ($shouldProcess) {
            $actions += @{
                ItemType = $item.ItemType
                ItemURL = $item.ItemURL
                ListId = if ($item.ListId) { $item.ListId } else { "00000000-0000-0000-0000-000000000000" }
                ListItemId = if ($item.ListItemId) { $item.ListItemId } else { $null }
                GroupsToRemove = $groupsToRemove
                HasEveryone = $hasEveryone
                HasEveryoneExceptExternal = $hasEveryoneExceptExternal
                RemoveEveryone = $FilterCriteria.RemoveEveryone
                RemoveEveryoneExceptExternal = $FilterCriteria.RemoveEveryoneExceptExternal
                PeopleCount = $peopleCount
            }
        }
    }
    Write-Log "Generated $($actions.Count) subweb permission actions" "SUCCESS"
    return $actions
}

function Remove-SubwebPermissions {
    param(
        [array]$Actions,
        [bool]$WhatIfMode
    )
    $processedCount = 0
    $errorCount = 0
    foreach ($action in $Actions) {
        Write-Log "Processing $($action.ItemType): $($action.ItemURL)" "INFO"
        try {
            if ($action.RemoveEveryone -and $action.HasEveryone) {
                if ($WhatIfMode) {
                    Write-Log "WHAT-IF: Would remove 'Everyone' from $($action.ItemType)" "INFO"
                }
                else {
                    Write-Log "Would remove 'Everyone' from $($action.ItemType)" "INFO"
                }
            }
            if ($action.RemoveEveryoneExceptExternal -and $action.HasEveryoneExceptExternal) {
                if ($WhatIfMode) {
                    Write-Log "WHAT-IF: Would remove 'Everyone except external users' from $($action.ItemType)" "INFO"
                }
                else {
                    Write-Log "Would remove 'Everyone except external users' from $($action.ItemType)" "INFO"
                }
            }
            foreach ($group in $action.GroupsToRemove) {
                if ($WhatIfMode) {
                    Write-Log "WHAT-IF: Would remove group '$group' from $($action.ItemType)" "INFO"
                }
                else {
                    Write-Log "Would remove group '$group' from $($action.ItemType)" "INFO"
                }
            }
            $processedCount++
        }
        catch {
            Write-Log "Error processing $($action.ItemType) '$($action.ItemURL)': $($_.Exception.Message)" "ERROR"
            $errorCount++
        }
    }
    Write-Log "Processing complete: $processedCount successful, $errorCount errors" "INFO"
}

try {
    $startTime = Get-Date
    Write-Log "SharePoint Subweb Permission Removal Script - Starting" "INFO"
    Write-Log "Target Subweb: $SiteUrl" "INFO"
    Write-Log "Main Site: $MainSiteUrl" "INFO"
    Write-Log "Mode: $(if ($WhatIf) { 'SIMULATION (WhatIf)' } else { 'PRODUCTION' })" "INFO"
    Initialize-RequiredModules

    $authMethod = Get-AuthMethod

    Write-Log "Retrieving configuration from main site..." "INFO"
    $allPermissionData = Get-ConfigurationFromMainSite -MainSiteUrl $MainSiteUrl -ListName $ConfigListName -ConfigFileName $ConfigFileName -AuthenticationMethod $authMethod

    Write-Log "Connecting to target subweb..." "INFO"
    $authParams = @{
        SiteUrl = $SiteUrl
        AuthenticationMethod = $authMethod
        ClientId = $ClientId
        ClientIdForSecret = $ClientIdForSecret
        ClientSecret = $ClientSecret
        TenantIdForSecret = $TenantIdForSecret
        ClientIdForCert = $ClientIdForCert
        CertificateThumbprint = $CertificateThumbprint
        CertificatePath = $CertificatePath
        CertificatePassword = $CertificatePassword
        TenantIdForCert = $TenantIdForCert
        ClientIdForDevice = $ClientIdForDevice
    }
    $currentWeb = Connect-SharePointSite @authParams

    $subwebItems = Get-SubwebItems -AllPermissionData $allPermissionData -CurrentSiteUrl $SiteUrl
    if ($subwebItems.Count -eq 0) {
        Write-Log "No items found for this subweb in the permission report" "WARNING"
        Write-Log "Verify that the subweb URL matches exactly: $SiteUrl" "WARNING"
        return
    }

    $filterCriteria = @{
        GroupsToRemove = $GroupsToRemove
        ItemTypes = $ItemTypesToProcess
        RemoveEveryone = $RemoveEveryone
        RemoveEveryoneExceptExternal = $RemoveEveryoneExceptExternal
    }
    $actions = New-SubwebActions -SubwebItems $subwebItems -FilterCriteria $filterCriteria
    if ($actions.Count -eq 0) {
        Write-Log "No actions generated based on filter criteria" "WARNING"
        return
    }

    Write-Log "Processing $($actions.Count) permission actions..." "INFO"
    Remove-SubwebPermissions -Actions $actions -WhatIfMode $WhatIf
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    Write-Log "Subweb script session ended" "INFO"
}


<#
.SYNOPSIS
    Subweb SharePoint Permission Removal Script

.DESCRIPTION
    Processes permission removal for SharePoint subwebs using configuration from the main site.
    Connects to the main site to get the permission report, then connects to the target subweb
    to process only items that belong to that subweb.

.PARAMETER SiteUrl
    Target subweb URL (e.g., https://tenant.sharepoint.com/teams/NBH/DAGSubsite)

.PARAMETER MainSiteUrl
    Main site URL where the configuration file is located (e.g., https://tenant.sharepoint.com/teams/NBH)

.PARAMETER GroupsToRemove
    Groups to remove permissions for (e.g., -GroupsToRemove @("Visitors", "Members"))

.PARAMETER RemoveEveryone
    Remove "Everyone" permissions from items

.PARAMETER RemoveEveryoneExceptExternal
    Remove "Everyone except external users" permissions from items

.PARAMETER WhatIf
    Simulate actions without making changes (recommended for review)

.PARAMETER ClientId
    Azure AD App Client ID for Interactive authentication (recommended for most users)

.PARAMETER ClientIdForSecret, TenantIdForSecret, ClientSecret
    Use these for Client Secret authentication (service principal)

.PARAMETER ClientIdForCert, TenantIdForCert, CertificateThumbprint/CertificatePath, CertificatePassword
    Use these for Certificate authentication

.PARAMETER ClientIdForDevice
    Use this for Device Login authentication

.EXAMPLE
    # Interactive login (recommended)
    .\SARPermission1subweb.ps1 `
        -SiteUrl "https://tenant.sharepoint.com/teams/NBH/DAGSubsite" `
        -MainSiteUrl "https://tenant.sharepoint.com/teams/NBH" `
        -RemoveEveryone `
        -WhatIf `
        -ClientId "<your-client-id>"

.EXAMPLE
    # Remove specific groups (simulation mode)
    .\SARPermission1subweb.ps1 `
        -SiteUrl "https://tenant.sharepoint.com/teams/NBH/DAGSubsite" `
        -MainSiteUrl "https://tenant.sharepoint.com/teams/NBH" `
        -GroupsToRemove @("Visitors", "Members") `
        -WhatIf `
        -ClientId "<your-client-id>"

.EXAMPLE
    # Using Client Secret authentication
    .\SARPermission1subweb.ps1 `
        -SiteUrl "https://tenant.sharepoint.com/teams/NBH/DAGSubsite" `
        -MainSiteUrl "https://tenant.sharepoint.com/teams/NBH" `
        -RemoveEveryone `
        -ClientIdForSecret "<your-client-id>" `
        -TenantIdForSecret "<your-tenant-id>" `
        -ClientSecret (ConvertTo-SecureString "<your-secret>" -AsPlainText -Force)

.NOTES
    - Always run in WhatIf mode first to review actions before making changes.
    - The script auto-detects authentication method based on parameters provided.
    - Ensure your Azure AD App has the necessary permissions for SharePoint access.
    - For subwebs, the URL must match exactly as in the permission report.
    - Log file is written to .\Subweb-PermissionRemovalLog.txt by default.
#>