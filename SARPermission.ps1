# SharePoint Main Site Permission Removal Script
# This script processes permissions for the main site only and identifies subweb items for separate processing

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory=$true, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory=$true, ParameterSetName = 'ClientSecret')]
    [Parameter(Mandatory=$true, ParameterSetName = 'Certificate')]
    [Parameter(Mandatory=$true, ParameterSetName = 'DeviceLogin')]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigFileName = "",
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigListName = "DO_NOT_DELETE_REVIEW_INSTANCE",
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = ".\MainSite-PermissionRemovalLog.txt",
    
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
    
    # Authentication Parameters - Interactive
    [Parameter(Mandatory=$true, ParameterSetName = 'Interactive')]
    [string]$ClientId,
    
    # Authentication Parameters - Client Secret
    [Parameter(Mandatory=$true, ParameterSetName = 'ClientSecret')]
    [string]$ClientIdForSecret,
    
    [Parameter(Mandatory=$true, ParameterSetName = 'ClientSecret')]
    [SecureString]$ClientSecret,
    
    [Parameter(Mandatory=$true, ParameterSetName = 'ClientSecret')]
    [string]$TenantIdForSecret,
    
    # Authentication Parameters - Certificate
    [Parameter(Mandatory=$true, ParameterSetName = 'Certificate')]
    [string]$ClientIdForCert,
    
    [Parameter(Mandatory=$false, ParameterSetName = 'Certificate')]
    [string]$CertificateThumbprint = "",
    
    [Parameter(Mandatory=$false, ParameterSetName = 'Certificate')]
    [string]$CertificatePath = "",
    
    [Parameter(Mandatory=$false, ParameterSetName = 'Certificate')]
    [SecureString]$CertificatePassword = $null,
    
    [Parameter(Mandatory=$true, ParameterSetName = 'Certificate')]
    [string]$TenantIdForCert,
    
    # Authentication Parameters - Device Login
    [Parameter(Mandatory=$true, ParameterSetName = 'DeviceLogin')]
    [string]$ClientIdForDevice
)

# Script Configuration
$Script:Config = @{
    RequiredModules = @("PnP.PowerShell", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups")
    ErrorCodes = @{
        AUTH_FAILED = 1001
        CONFIG_NOT_FOUND = 1002
        VALIDATION_FAILED = 1003
        PERMISSION_DENIED = 1004
        NETWORK_ERROR = 1005
    }
}

# Initialize script variables
$Script:AuthMethod = $PSCmdlet.ParameterSetName
$Script:EffectiveClientId = switch ($Script:AuthMethod) {
    'Interactive' { $ClientId }
    'ClientSecret' { $ClientIdForSecret }
    'Certificate' { $ClientIdForCert }
    'DeviceLogin' { $ClientIdForDevice }
    default { $null }
}
$Script:EffectiveTenantId = switch ($Script:AuthMethod) {
    'ClientSecret' { $TenantIdForSecret }
    'Certificate' { $TenantIdForCert }
    default { $null }
}

# Initialize modules
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

# Logging function - Fixed to handle empty messages
function Write-Log {
    param(
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Message = " ",
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    # Handle empty messages by using a space
    if ([string]::IsNullOrEmpty($Message)) {
        $Message = " "
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        "DEBUG" { "Cyan" }
        default { "White" }
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

# Helper function to write blank lines
function Write-BlankLine {
    Write-Host ""
}

# Authentication function
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
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret -Tenant $TenantId -ErrorAction Stop
            }
            'Certificate' {
                if ($CertificateThumbprint) {
                    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant $TenantId -ErrorAction Stop
                }
                elseif ($CertificatePath) {
                    if ($CertificatePassword) {
                        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Tenant $TenantId -ErrorAction Stop
                    }
                    else {
                        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -Tenant $TenantId -ErrorAction Stop
                    }
                }
            }
            'DeviceLogin' {
                Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientId -ErrorAction Stop
            }
        }
        
        $web = Get-PnPWeb -ErrorAction Stop
        Write-Log "Successfully connected to: $($web.Title) ($($web.Url))" "SUCCESS"
    }
    catch {
        Write-Log "Failed to connect to SharePoint: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Get latest permission report file
function Get-LatestPermissionReport {
    param([string]$ListName)
    
    try {
        Write-Log "Searching for permission report files in list: $ListName" "INFO"
        
        $list = Get-PnPList -Identity $ListName -ErrorAction Stop
        
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
    <RowLimit>10</RowLimit>
</View>
"@
        
        $allFiles = Get-PnPListItem -List $ListName -Query $camlQuery -ErrorAction Stop
        
        if ($allFiles.Count -eq 0) {
            throw "No files found in the list '$ListName'"
        }
        
        # Find the latest valid permission report file
        foreach ($file in $allFiles) {
            $fileName = $file.FieldValues.FileLeafRef
            $fileUrl = $file.FieldValues.FileRef
            $fileSize = if ($file.FieldValues.File_x0020_Size) { $file.FieldValues.File_x0020_Size } else { 0 }
            
            if ($fileSize -lt 100) { continue }
            
            try {
                $sampleContent = Get-PnPFile -Url $fileUrl -AsString -ErrorAction Stop
                
                if ($sampleContent -and ($sampleContent.Contains('"ItemType"') -or $sampleContent.Contains('"TenantId"'))) {
                    Write-Log "Found valid permission report: $fileName" "SUCCESS"
                    
                    $tempPath = Join-Path $env:TEMP "mainsite-permission-report-$(Get-Date -Format 'yyyyMMddHHmmss')"
                    Set-Content -Path $tempPath -Value $sampleContent -Encoding UTF8 -ErrorAction Stop
                    
                    return @{
                        FileName = $fileName
                        FilePath = $tempPath
                        FileUrl = $fileUrl
                        CreatedDate = $file.FieldValues.Created
                        FileSize = $fileSize
                    }
                }
            }
            catch {
                continue
            }
        }
        
        throw "No valid permission report files found"
    }
    catch {
        Write-Log "Failed to get permission report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Parse configuration file
function Get-ConfigurationFromFile {
    param([string]$FilePath)
    
    try {
        Write-Log "Loading permission report from: $FilePath" "INFO"
        $fileContent = Get-Content -Path $FilePath -Raw -ErrorAction Stop
        
        $permissionData = @()
        $lines = ($fileContent -split "`r?`n") | Where-Object { $_.Trim() -ne "" }
        
        # Detect JSONL format
        if ($lines.Count -gt 1 -and $lines[0].Trim().StartsWith('{')) {
            Write-Log "Detected JSONL format, parsing..." "INFO"
            foreach ($line in $lines) {
                try {
                    $jsonObject = $line.Trim() | ConvertFrom-Json
                    if ($jsonObject.ItemType -or $jsonObject.TenantId) {
                        $permissionData += $jsonObject
                    }
                }
                catch {
                    Write-Log "Failed to parse line: $($line.Substring(0, [Math]::Min(50, $line.Length)))" "WARNING"
                }
            }
        }
        else {
            $parsedData = $fileContent | ConvertFrom-Json
            $permissionData = if ($parsedData -is [array]) { $parsedData } else { @($parsedData) }
        }
        
        Write-Log "Successfully loaded $($permissionData.Count) permission records" "SUCCESS"
        return $permissionData
    }
    catch {
        Write-Log "Failed to load configuration: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Separate main site and subweb items
function Separate-PermissionItems {
    param(
        [array]$PermissionData,
        [string]$MainSiteUrl
    )
    
    $mainSiteItems = @()
    $subwebItems = @()
    $subwebs = @{}
    
    # Normalize the main site URL for comparison
    $normalizedMainSiteUrl = $MainSiteUrl.TrimEnd('/')
    
    foreach ($item in $PermissionData) {
        # Check various fields for parent web information
        $parentWeb = $null
        
        if ($item.PSObject.Properties['ParentWeb'] -and $item.ParentWeb) {
            $parentWeb = $item.ParentWeb
        }
        elseif ($item.PSObject.Properties['SiteURL'] -and $item.SiteURL) {
            $parentWeb = $item.SiteURL
        }
        elseif ($item.PSObject.Properties['WebURL'] -and $item.WebURL) {
            $parentWeb = $item.WebURL
        }
        
        # Normalize the parent web URL for comparison
        if ($parentWeb) {
            $parentWeb = $parentWeb.TrimEnd('/')
            # If parent web is relative, make it absolute
            if ($parentWeb.StartsWith('/')) {
                $mainSiteUri = [Uri]$normalizedMainSiteUrl
                $parentWeb = "$($mainSiteUri.Scheme)://$($mainSiteUri.Host)$parentWeb"
            }
        }
        else {
            # If no parent web specified, assume it's the main site
            $parentWeb = $normalizedMainSiteUrl
        }
        
        # Compare normalized URLs
        if ($parentWeb -eq $normalizedMainSiteUrl) {
            $mainSiteItems += $item
        }
        else {
            $subwebItems += $item
            if (-not $subwebs.ContainsKey($parentWeb)) {
                $subwebs[$parentWeb] = @()
            }
            $subwebs[$parentWeb] += $item
        }
    }
    
    Write-Log "Found $($mainSiteItems.Count) main site items and $($subwebItems.Count) subweb items" "INFO"
    
    if ($subwebs.Count -gt 0) {
        Write-Log "Subwebs detected:" "INFO"
        foreach ($subweb in $subwebs.Keys) {
            Write-Log "  - $subweb ($($subwebs[$subweb].Count) items)" "INFO"
        }
    }
    
    return @{
        MainSiteItems = $mainSiteItems
        SubwebItems = $subwebItems
        Subwebs = $subwebs
    }
}

# Generate actions for main site only
function New-MainSiteActions {
    param(
        [array]$MainSiteItems,
        [hashtable]$FilterCriteria
    )
    
    $actions = @()
    
    foreach ($item in $MainSiteItems) {
        if ($FilterCriteria.ItemTypes -and $item.ItemType -notin $FilterCriteria.ItemTypes) {
            continue
        }
        
        $peopleCount = if ($item.PeopleCount) { [int]$item.PeopleCount } else { 0 }
        if ($peopleCount -lt $FilterCriteria.MinPeopleCount) {
            continue
        }
        
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
            }
        }
    }
    
    Write-Log "Generated $($actions.Count) main site permission actions" "SUCCESS"
    return $actions
}

# Simple permission removal (WhatIf mode for main site)
function Remove-MainSitePermissions {
    param(
        [array]$Actions,
        [bool]$WhatIfMode
    )
    
    foreach ($action in $Actions) {
        Write-Log "Processing $($action.ItemType): $($action.ItemURL)" "INFO"
        
        # Process Everyone
        if ($action.RemoveEveryone -and $action.HasEveryone) {
            if ($WhatIfMode) {
                Write-Log "WHAT-IF: Would remove 'Everyone' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
            else {
                Write-Log "Would remove 'Everyone' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
        }
        
        # Process Everyone Except External
        if ($action.RemoveEveryoneExceptExternal -and $action.HasEveryoneExceptExternal) {
            if ($WhatIfMode) {
                Write-Log "WHAT-IF: Would remove 'Everyone except external users' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
            else {
                Write-Log "Would remove 'Everyone except external users' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
        }
        
        # Process specific groups
        foreach ($group in $action.GroupsToRemove) {
            if ($WhatIfMode) {
                Write-Log "WHAT-IF: Would remove group '$group' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
            else {
                Write-Log "Would remove group '$group' from $($action.ItemType) '$($action.ItemURL)'" "INFO"
            }
        }
    }
}

# Report subweb items and provide guidance
function Report-SubwebItems {
    param([hashtable]$Subwebs)
    
    if ($Subwebs.Count -eq 0) {
        Write-Log "No subweb items found - all processing complete!" "SUCCESS"
        return
    }
    
    Write-BlankLine
    Write-Log "=== SUBWEB ITEMS DETECTED ===" "WARNING"
    Write-Log "The following subweb items were SKIPPED and need separate processing:" "WARNING"
    Write-BlankLine
    
    foreach ($subwebUrl in $Subwebs.Keys) {
        $items = $Subwebs[$subwebUrl]
        Write-Log "Subweb: $subwebUrl ($($items.Count) items)" "WARNING"
        
        $itemsByType = $items | Group-Object ItemType
        foreach ($typeGroup in $itemsByType) {
            Write-Log "  - $($typeGroup.Name): $($typeGroup.Count) items" "INFO"
            foreach ($item in $typeGroup.Group | Select-Object -First 3) {
                Write-Log "    * $($item.ItemURL)" "INFO"
            }
            if ($typeGroup.Count -gt 3) {
                Write-Log "    * ... and $($typeGroup.Count - 3) more items" "INFO"
            }
        }
        Write-BlankLine
    }
    
    Write-Log "=== NEXT STEPS ===" "WARNING"
    Write-Log "Run the Subweb Permission Removal Script for each subweb:" "WARNING"
    Write-BlankLine
    
    foreach ($subwebUrl in $Subwebs.Keys) {
        Write-Log "For subweb: $subwebUrl" "WARNING"
        Write-Log ".\Remove-SubwebPermissions.ps1 -SiteUrl `"$subwebUrl`" -MainSiteUrl `"$SiteUrl`" [other parameters]" "WARNING"
        Write-BlankLine
    }
}

# Main execution
try {
    $startTime = Get-Date
    Write-Log "SharePoint Main Site Permission Removal Script - Starting" "INFO"
    Write-Log "Target Site: $SiteUrl" "INFO"
    Write-Log "Mode: $(if ($WhatIf) { 'SIMULATION (WhatIf)' } else { 'PRODUCTION' })" "INFO"
    
    # Initialize and connect
    Initialize-RequiredModules
    
    $authParams = @{
        SiteUrl = $SiteUrl
        AuthenticationMethod = $Script:AuthMethod
        ClientId = $Script:EffectiveClientId
        TenantId = $Script:EffectiveTenantId
        ClientSecret = if ($ClientSecret) { $ClientSecret } else { $null }
        CertificateThumbprint = $CertificateThumbprint
        CertificatePath = $CertificatePath
        CertificatePassword = $CertificatePassword
    }
    
    Connect-SharePointSite @authParams
    
    # Get configuration
    $fileInfo = Get-LatestPermissionReport -ListName $ConfigListName
    $permissionData = Get-ConfigurationFromFile -FilePath $fileInfo.FilePath
    
    # Separate main site and subweb items
    $separatedData = Separate-PermissionItems -PermissionData $permissionData -MainSiteUrl $SiteUrl
    
    # Process main site items only
    if ($separatedData.MainSiteItems.Count -gt 0) {
        $filterCriteria = @{
            GroupsToRemove = $GroupsToRemove
            ItemTypes = $ItemTypesToProcess
            MinPeopleCount = 0
            RemoveEveryone = $RemoveEveryone
            RemoveEveryoneExceptExternal = $RemoveEveryoneExceptExternal
        }
        
        $mainSiteActions = New-MainSiteActions -MainSiteItems $separatedData.MainSiteItems -FilterCriteria $filterCriteria
        
        if ($mainSiteActions.Count -gt 0) {
            Write-Log "Processing $($mainSiteActions.Count) main site permission actions..." "INFO"
            Remove-MainSitePermissions -Actions $mainSiteActions -WhatIfMode $WhatIf
            Write-Log "Main site processing completed!" "SUCCESS"
        }
        else {
            Write-Log "No main site actions generated based on criteria" "WARNING"
        }
    }
    else {
        Write-Log "No main site items found to process" "INFO"
    }
    
    # Report subweb items
    Report-SubwebItems -Subwebs $separatedData.Subwebs
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Log "Script completed in $($duration.ToString('hh\:mm\:ss'))" "SUCCESS"
    
} catch {
    Write-Log "Script failed: $($_.Exception.Message)" "ERROR"
    throw
} finally {
    if ($fileInfo -and $fileInfo.FilePath -and (Test-Path $fileInfo.FilePath)) {
        Remove-Item $fileInfo.FilePath -Force -ErrorAction SilentlyContinue
    }
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    Write-Log "Script session ended" "INFO"
}

<#
.SYNOPSIS
    Main Site SharePoint Permission Removal Script

.DESCRIPTION
    Processes permission removal for the main SharePoint site only. 
    Identifies and reports subweb items for separate processing.

.PARAMETER SiteUrl
    Main SharePoint site URL

.PARAMETER GroupsToRemove
    Groups to remove permissions for

.PARAMETER RemoveEveryone
    Remove "Everyone" permissions

.PARAMETER RemoveEveryoneExceptExternal
    Remove "Everyone except external users" permissions

.EXAMPLE
    .\Remove-MainSitePermissions.ps1 -ClientId "your-id" -SiteUrl "https://tenant.sharepoint.com/teams/DAG" -RemoveEveryone -WhatIf

.EXAMPLE
    # Remove specific groups (simulation mode)
    .\Remove-MainSitePermissions.ps1 `
        -SiteUrl "https://tenant.sharepoint.com/teams/DAG" `
        -GroupsToRemove @("Visitors", "Members", "customgroup") `
        -WhatIf `
        -ClientId "<your-client-id>"
#>
