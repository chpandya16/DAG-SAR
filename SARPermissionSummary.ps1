# SiteCollectionPermissionSummary.ps1

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,            # Path to CSV file with SiteCollectionUrl column
    [Parameter(Mandatory=$true)]
    [string]$ClientId,           # Azure AD App ClientId for Interactive login
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = ".\SiteCollectionPermissionSummary.csv"
)

function Write-Log {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

# Import site collection URLs
$sites = Import-Csv -Path $CsvPath

# Prepare output
$summary = @()

foreach ($site in $sites) {
    $SiteUrl = $site.SiteCollectionUrl
    Write-Log "Processing site: $SiteUrl"

    try {
        # Connect to site
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ForceAuthentication -ErrorAction Stop

        # Get the list
        $list = Get-PnPList -Identity "DO_NOT_DELETE_REVIEW_INSTANCE" -ErrorAction Stop

        # Get latest file (by Created date)
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
$files = Get-PnPListItem -List $list -Query $camlQuery -ErrorAction Stop

foreach ($file in $files) {
    $fileName = $file.FieldValues.FileLeafRef
    $fileUrl = $file.FieldValues.FileRef
    Write-Log "Checking file: $fileName"
    Write-Log "FileRef: $fileUrl"
    if ($fileUrl) {
        Write-Log "Using file: $fileName"
        # Download file content
        $fileContent = Get-PnPFile -Url $fileUrl -AsString -ErrorAction Stop
        # ... (rest of your parsing logic here)
        break
    }
}

# Download file content
$fileContent = Get-PnPFile -Url $fileUrl -AsString -ErrorAction Stop

        # Parse JSON/JSONL
        $lines = ($fileContent -split "`r?`n") | Where-Object { $_.Trim() -ne "" }
        $items = @()
        if ($lines.Count -gt 1 -and $lines[0].Trim().StartsWith('{')) {
            foreach ($line in $lines) {
                try {
                    $obj = $line | ConvertFrom-Json
                    $items += $obj
                } catch {}
            }
        } else {
            try {
                $parsed = $fileContent | ConvertFrom-Json
                $items = if ($parsed -is [array]) { $parsed } else { @($parsed) }
            } catch {
                Write-Log "Failed to parse permission report file in $SiteUrl"
                continue
            }
        }

        # Summarize sharing info
        foreach ($item in $items) {
            $type = $item.ItemType
            $url = $item.ItemURL
            $groups = @()
            if ($item.GroupDetails) {
                foreach ($g in $item.GroupDetails) {
                    $groups += $g.Name
                }
            }
            $everyone = if ($item.HasEveryone) { "Everyone" } else { "" }
            $everyoneExceptExternal = if ($item.HasEveryoneExceptExternalUser) { "EveryoneExceptExternal" } else { "" }
            $sharedWith = ($groups + $everyone + $everyoneExceptExternal) -ne "" | Where-Object { $_ }
            $summary += [PSCustomObject]@{
                SiteCollection = $SiteUrl
                ItemType       = $type
                ItemURL        = $url
                SharedWith     = ($sharedWith -join "; ")
            }
        }
    } catch {
        Write-Log "Error processing $SiteUrl : $($_.Exception.Message)"
        continue
    } finally {
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    }
}

# Export summary
$summary | Export-Csv -Path $ReportPath -NoTypeInformation
Write-Log "Summary report written to $ReportPath"


<#
.SYNOPSIS
    Site Collection Permission Summary Script

.DESCRIPTION
    This script reads a CSV file containing SharePoint site collection URLs.
    For each site collection, it connects using PnP PowerShell, locates the 'DO_NOT_DELETE_REVIEW_INSTANCE' list,
    finds the latest permission report file (regardless of extension), parses the file (JSON or JSONL format),
    and creates a summary report showing what items (Web, File, List, Folder) are shared and with whom.

.PARAMETER CsvPath
    Path to the CSV file containing site collection URLs. The CSV must have a column named 'SiteCollectionUrl'.

.PARAMETER ClientId
    Azure AD App Client ID for Interactive authentication.

.PARAMETER ReportPath
    Path to output the summary report CSV. Defaults to '.\SiteCollectionPermissionSummary.csv'.

.EXAMPLE
    # Run the script with a CSV of site collections and your Azure AD App Client ID
    .\SiteCollectionPermissionSummary.ps1 `
        -CsvPath "C:\Scripts\SiteInput.csv" `
        -ClientId "<your-client-id>"

.EXAMPLE
    # Specify a custom output path for the summary report
    .\SiteCollectionPermissionSummary.ps1 `
        -CsvPath "C:\Scripts\SiteInput.csv" `
        -ClientId "<your-client-id>" `
        -ReportPath "C:\Scripts\MySummary.csv"

.NOTES
    - The script uses Interactive authentication. You will be prompted to sign in for each site collection.
    - The script auto-detects and parses JSON or JSONL formatted permission report files.
    - The summary report lists each item (Web, File, List, Folder) and who it is shared with.
    - The script skips sites where no valid permission report file is found.
    - Ensure your Azure AD App has the necessary permissions for SharePoint access.
    - The output CSV can be opened in Excel for review.

.REQUIRED PERMISSIONS
    The Azure AD App (used for ClientId) must have the following Microsoft Graph and SharePoint API permissions:
      - Microsoft Graph: Sites.Read.All (Application or Delegated)
      - SharePoint: Sites.Read.All (Application or Delegated)
      - Optionally, Sites.Manage.All if you want to modify permissions
    Admin consent is required for these permissions.

#>