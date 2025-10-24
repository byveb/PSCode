<#
.SYNOPSIS
    Simple optimized PowerBI OnPrem deployment script
.PARAMETER ConfigPath
    Path to JSON configuration file
.DESCRIPTION
    Features:
    - Auto-creates folders if they don't exist
    - Reports array in DataSources/Permissions/Parameters: applies config only to specified reports
    - No Reports array: applies config to all reports
    - ReportMappings with priority: ReportTargetPath > TargetFolderPath > global TargetFolderPath
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

# Helper function to check if config applies to current report
function Test-ConfigApplies {
    param($Config, $FileName)

    if (-not $Config.Reports) {
        return $true  # No Reports array = applies to all
    }

    return $Config.Reports -contains $FileName
}

# Load configuration
$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

# Process each report configuration
foreach ($report in $config.Reports) {
    if ($report.Type -ne 'OnPrem') { continue }

    Write-Host "`n=== Deploying to OnPrem: $($report.TargetServer) ===" -ForegroundColor Cyan

    # Get all .pbix files from FolderPath
    $pbixFiles = Get-ChildItem -Path $report.FolderPath -Filter "*.pbix" -Recurse

    foreach ($file in $pbixFiles) {
        $fileName = $file.Name

        # Find matching report mapping
        $mapping = $report.ReportMappings | Where-Object { $fileName -match $_.SourcePattern }

        # Determine target path with priority: ReportTargetPath > TargetFolderPath > global
        if ($mapping -and $mapping.ReportTargetPath) {
            $targetPath = $mapping.ReportTargetPath
        } elseif ($mapping -and $mapping.TargetFolderPath) {
            $targetPath = "$($mapping.TargetFolderPath)/$($file.BaseName)"
        } else {
            $targetPath = "$($report.TargetFolderPath)/$($file.BaseName)"
        }

        # Parse folder and report name
        $lastSlash = $targetPath.LastIndexOf('/')
        $targetFolder = $targetPath.Substring(0, $lastSlash)
        $reportName = $targetPath.Substring($lastSlash + 1)

        Write-Host "  Deploying: $fileName -> $targetPath" -ForegroundColor Yellow

        # Build API URL
        $apiUrl = "$($report.TargetServer)/api/v2.0"

        # Create folder if not exists
        $folderPath = $targetFolder -replace '^/', ''
        $folderParts = $folderPath -split '/'
        $currentPath = ''

        foreach ($part in $folderParts) {
            $currentPath += "/$part"
            try {
                $checkUrl = "$apiUrl/Folders(Path='$currentPath')"
                Invoke-RestMethod -Uri $checkUrl -Method Get -UseDefaultCredentials -ErrorAction SilentlyContinue | Out-Null
            } catch {
                # Create folder
                $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
                if (!$parentPath) { $parentPath = '/' }

                $createBody = @{
                    Path = $currentPath
                    Name = $part
                    ParentFolderPath = $parentPath
                } | ConvertTo-Json

                Invoke-RestMethod -Uri "$apiUrl/Folders" -Method Post -Body $createBody -ContentType 'application/json' -UseDefaultCredentials
                Write-Host "    Created folder: $currentPath" -ForegroundColor Green
            }
        }

        # Check if report exists
        $reportPath = "$targetFolder/$reportName"
        $exists = $false
        try {
            Invoke-RestMethod -Uri "$apiUrl/PowerBIReports(Path='$reportPath')" -Method Get -UseDefaultCredentials -ErrorAction SilentlyContinue | Out-Null
            $exists = $true
        } catch { }

        # Upload report
        if ($exists -and -not $report.Overwrite) {
            Write-Host "    Report exists, skipping (Overwrite=false)" -ForegroundColor Gray
            continue
        }

        $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
        $fileContent = [System.Convert]::ToBase64String($fileBytes)

        $uploadBody = @{
            Path = $reportPath
            Content = $fileContent
            ContentType = 'application/octet-stream'
        } | ConvertTo-Json

        $method = if ($exists) { 'PUT' } else { 'POST' }
        Invoke-RestMethod -Uri "$apiUrl/PowerBIReports" -Method $method -Body $uploadBody -ContentType 'application/json' -UseDefaultCredentials

        Write-Host "    Uploaded successfully" -ForegroundColor Green

        # Configure DataSources - filter by Reports array
        if ($report.DataSources -and $report.DataSources.Count -gt 0) {
            Start-Sleep -Seconds 2

            # Get report data sources
            $reportDataSources = Invoke-RestMethod -Uri "$apiUrl/PowerBIReports(Path='$reportPath')/DataSources" -Method Get -UseDefaultCredentials

            foreach ($ds in $reportDataSources.value) {
                # Find applicable datasource configs (with Reports array matching or no Reports array)
                $applicableConfigs = $report.DataSources | Where-Object {
                    $_.Name -eq $ds.Name -and (Test-ConfigApplies -Config $_ -FileName $fileName)
                }

                foreach ($configDs in $applicableConfigs) {
                    Write-Host "    Updating datasource: $($ds.Name)" -ForegroundColor Gray

                    $dsBody = @{
                        ConnectionString = $configDs.ConnectionString
                        CredentialType = $configDs.CredentialType
                    }

                    if ($configDs.Username) { $dsBody.UserName = $configDs.Username }
                    if ($configDs.Password) { $dsBody.Password = $configDs.Password }
                    if ($configDs.UseDefaultCredentials) { $dsBody.CredentialType = 'Windows' }

                    $dsUpdateBody = $dsBody | ConvertTo-Json
                    Invoke-RestMethod -Uri "$apiUrl/PowerBIReports(Path='$reportPath')/DataSources('$($ds.Id)')" -Method Patch -Body $dsUpdateBody -ContentType 'application/json' -UseDefaultCredentials
                }
            }
        }

        # Configure Permissions - filter by Reports array
        if ($report.Permissions) {
            # Get applicable permissions (with Reports array matching or no Reports array)
            $applicablePerms = $report.Permissions | Where-Object {
                Test-ConfigApplies -Config $_ -FileName $fileName
            }

            foreach ($perm in $applicablePerms) {
                Write-Host "    Setting permission: $($perm.Principal) = $($perm.Role)" -ForegroundColor Gray

                $permBody = @{
                    Principal = $perm.Principal
                    Roles = @($perm.Role)
                } | ConvertTo-Json

                try {
                    Invoke-RestMethod -Uri "$apiUrl/PowerBIReports(Path='$reportPath')/Policies" -Method Post -Body $permBody -ContentType 'application/json' -UseDefaultCredentials
                } catch {
                    Write-Host "    Permission may already exist" -ForegroundColor DarkGray
                }
            }
        }

        # Configure Parameters - filter by Reports array
        if ($report.Parameters) {
            # Get applicable parameters (with Reports array matching or no Reports array)
            $applicableParams = $report.Parameters | Where-Object {
                Test-ConfigApplies -Config $_ -FileName $fileName
            }

            foreach ($param in $applicableParams) {
                Write-Host "    Setting parameter: $($param.Name) = $($param.Value)" -ForegroundColor Gray

                $paramBody = @{
                    Name = $param.Name
                    Value = $param.Value
                } | ConvertTo-Json

                Invoke-RestMethod -Uri "$apiUrl/PowerBIReports(Path='$reportPath')/Parameters" -Method Patch -Body $paramBody -ContentType 'application/json' -UseDefaultCredentials
            }
        }
    }
}

Write-Host "`n=== Deployment Complete ===" -ForegroundColor Green
