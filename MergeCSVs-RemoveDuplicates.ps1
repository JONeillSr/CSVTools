<#
.SYNOPSIS
    Merges multiple CSV files and removes duplicate entries based on the Email field

.DESCRIPTION
    Merges multiple CSV files and removes duplicate entries based on the Email field

.EXAMPLE
    .\MergeCSVs-RemoveDuplicates.ps1 -InputCSVs "ExtractedToEmails_org.csv","ExtractedToEmails.csv" -OutputCSV "MergedUniqueEmails.csv" -CaseInsensitive

.EXAMPLE
    .\MergeCSVs-RemoveDuplicates.ps1 -InputFolder "C:\Path\ToYourCSVs" -OutputCSV "MergedUniqueEmails.csv" -CaseInsensitive

.NOTES
    Author: John A. O'Neill Sr.
    Date: 05/09/2025
    Version: 1.0
    Change Date:
    Change Purpose:

    Prerequisites:
    - PowerShell 5.1 or later

    Logging:


.LINK
    https://learn.microsoft.com/en-us/office/vba/api/overview/outlook

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. This script does not generate any output objects.
#>

[CmdletBinding(DefaultParameterSetName='ByFiles')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='ByFiles')]
    [string[]]$InputCSVs,
    
    [Parameter(Mandatory=$true, ParameterSetName='ByFolder')]
    [string]$InputFolder,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputCSV,
    
    [Parameter(Mandatory=$false)]
    [string]$KeyField = "Email",
    
    [Parameter(Mandatory=$false)]
    [switch]$CaseInsensitive
)

# Add timestamp to console output
function Write-TimeStamp {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

# Determine the list of files to process
$filesToProcess = @()
if ($PSCmdlet.ParameterSetName -eq 'ByFolder') {
    if (-not (Test-Path -Path $InputFolder -PathType Container)) {
        Write-Error "Input folder '$InputFolder' does not exist or is not a directory."
        exit 1
    }
    
    $filesToProcess = Get-ChildItem -Path $InputFolder -Filter "*.csv" | Select-Object -ExpandProperty FullName
    Write-TimeStamp "Found $($filesToProcess.Count) CSV files in '$InputFolder'"
} else {
    foreach ($file in $InputCSVs) {
        if (-not (Test-Path -Path $file -PathType Leaf)) {
            Write-Error "Input file '$file' does not exist."
            exit 1
        }
        $filesToProcess += $file
    }
    Write-TimeStamp "Processing $($filesToProcess.Count) specified CSV files"
}

if ($filesToProcess.Count -eq 0) {
    Write-Error "No CSV files to process."
    exit 1
}

# Import and combine all CSV files
$allData = @()
$totalRecords = 0

foreach ($file in $filesToProcess) {
    Write-TimeStamp "Reading $file"
    try {
        $data = Import-Csv -Path $file
        $fileRecords = $data.Count
        $totalRecords += $fileRecords
        $allData += $data
        Write-TimeStamp "  - Added $fileRecords records from $file"
    }
    catch {
        Write-Error "Error reading file '$file': $_"
    }
}

Write-TimeStamp "Total records before deduplication: $totalRecords"

# Create a hashtable to store unique records
$uniqueRecords = @{}

# Process each record and store only unique ones
foreach ($record in $allData) {
    # Get the key value, default to the Email field
    $key = $record.$KeyField
    
    # Skip records with no key value
    if ([string]::IsNullOrWhiteSpace($key)) {
        continue
    }
    
    # Make key case-insensitive if requested
    if ($CaseInsensitive) {
        $key = $key.ToLower()
    }
    
    # Add to unique records if not already present
    if (-not $uniqueRecords.ContainsKey($key)) {
        $uniqueRecords[$key] = $record
    }
}

# Convert hashtable values back to an array
$uniqueData = $uniqueRecords.Values

Write-TimeStamp "Unique records after deduplication: $($uniqueData.Count)"

# Create output directory if it doesn't exist
$outputDir = Split-Path -Path $OutputCSV -Parent
if ($outputDir -and -not (Test-Path -Path $outputDir -PathType Container)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    Write-TimeStamp "Created output directory: $outputDir"
}

# Export the unique records to the output CSV
$uniqueData | Export-Csv -Path $OutputCSV -NoTypeInformation

Write-TimeStamp "Merged CSV file with unique records saved to: $OutputCSV"
