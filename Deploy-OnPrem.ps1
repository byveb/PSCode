<#
.SYNOPSIS
    Optimized PowerShell script for deploying Power BI reports to OnPrem Report Server.

.DESCRIPTION
    Lightweight CI/CD script for Power BI OnPrem deployments with no logging overhead.
    Features: Auto-create folders, bulk deployment, custom mappings, datasource config, permissions.

.PARAMETER ConfigPath
    Path to JSON configuration file (default: config.json)

.EXAMPLE
    .\Deploy-OnPrem.ps1 -ConfigPath "config.json"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "config.json"
)

$ErrorActionPreference = "Stop"

#region Core Functions

function Invoke-RSWebRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [hashtable]$Headers = @{},
        [string]$ContentType = "application/json"
    )

    $params = @{
        Uri                  = $Uri
        Method               = $Method
        UseDefaultCredentials = $true
        Headers              = $Headers
        ContentType          = $ContentType
    }

    if ($Body) {
        $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
    }

    try {
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        Write-Error "API request failed: $($_.Exception.Message)"
        throw
    }
}

function Test-FolderExists {
    param([string]$BaseUrl, [string]$Path)

    $encodedPath = [System.Web.HttpUtility]::UrlEncode($Path)
    $uri = "$BaseUrl/CatalogItems(Path='$encodedPath')"

    try {
        $item = Invoke-RSWebRequest -Uri $uri -Method GET
        return ($item.Type -eq "Folder")
    }
    catch {
        return $false
    }
}

function New-RSFolder {
    param([string]$BaseUrl, [string]$Path)

    $parts = $Path.Trim('/') -split '/'
    $currentPath = ""

    foreach ($part in $parts) {
        $currentPath += "/$part"

        if (-not (Test-FolderExists -BaseUrl $BaseUrl -Path $currentPath)) {
            $parentPath = ($currentPath -replace '/[^/]+$', '')
            if ($parentPath -eq "") { $parentPath = "/" }

            $body = @{
                Name = $part
                Path = $parentPath
                Type = "Folder"
            }

            Invoke-RSWebRequest -Uri "$BaseUrl/Folders" -Method POST -Body $body | Out-Null
        }
    }
}

function Get-ReportTargetPath {
    param(
        [string]$ReportName,
        [array]$ReportMappings,
        [string]$DefaultTargetFolder,
        [string]$DefaultReportName
    )

    # Priority 1: ReportTargetPath in mappings
    foreach ($mapping in $ReportMappings) {
        if ($ReportName -match $mapping.Pattern) {
            if ($mapping.ReportTargetPath) {
                return @{
                    Path = $mapping.ReportTargetPath -replace '/[^/]+$', ''
                    Name = ($mapping.ReportTargetPath -replace '^.*/', '')
                }
            }

            # Priority 2: TargetFolderPath + ReportName from mapping
            if ($mapping.TargetFolderPath -or $mapping.ReportName) {
                return @{
                    Path = if ($mapping.TargetFolderPath) { $mapping.TargetFolderPath } else { $DefaultTargetFolder }
                    Name = if ($mapping.ReportName) { $mapping.ReportName } else { $DefaultReportName }
                }
            }
        }
    }

    # Priority 3: Default values
    return @{
        Path = $DefaultTargetFolder
        Name = $DefaultReportName
    }
}

function Publish-PowerBIReport {
    param(
        [string]$BaseUrl,
        [string]$FilePath,
        [string]$TargetPath,
        [string]$ReportName,
        [bool]$Overwrite
    )

    # Check if report exists
    $fullPath = "$TargetPath/$ReportName"
    $encodedPath = [System.Web.HttpUtility]::UrlEncode($fullPath)
    $existingReport = $null

    try {
        $existingReport = Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')" -Method GET
    }
    catch {
        # Report doesn't exist
    }

    if ($existingReport -and -not $Overwrite) {
        Write-Host "  [SKIP] Report already exists: $fullPath" -ForegroundColor Yellow
        return $null
    }

    # Read file content
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $boundary = [System.Guid]::NewGuid().ToString()

    # Prepare multipart form data
    $LF = "`r`n"
    $bodyLines = @(
        "--$boundary",
        "Content-Disposition: form-data; name=`"file`"; filename=`"$ReportName.pbix`"",
        "Content-Type: application/octet-stream$LF",
        [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($fileBytes),
        "--$boundary--$LF"
    )

    $body = $bodyLines -join $LF

    # Upload report
    $uri = "$BaseUrl/PowerBIReports(Path='$encodedPath')?`$expand=DataSources"
    $headers = @{ "Content-Type" = "multipart/form-data; boundary=$boundary" }

    if ($existingReport) {
        $method = "PUT"
        Write-Host "  [UPDATE] $fullPath" -ForegroundColor Cyan
    }
    else {
        # For new reports, use different endpoint
        $uri = "$BaseUrl/PowerBIReports?`$expand=DataSources"
        $method = "POST"

        # Add path and name to headers for new reports
        $headers["x-ms-catalog-path"] = $TargetPath
        $headers["x-ms-catalog-name"] = $ReportName

        Write-Host "  [CREATE] $fullPath" -ForegroundColor Green
    }

    try {
        $report = Invoke-RestMethod -Uri $uri -Method $method -Body $body -Headers $headers -UseDefaultCredentials
        return $report
    }
    catch {
        Write-Error "Failed to publish report: $($_.Exception.Message)"
        throw
    }
}

function Update-DataSource {
    param(
        [string]$BaseUrl,
        [string]$ReportPath,
        [object]$DataSourceConfig
    )

    $encodedPath = [System.Web.HttpUtility]::UrlEncode($ReportPath)
    $dataSources = Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/DataSources" -Method GET

    # Check if DataSourceConfig is an array (multiple datasources) or single object
    $isMultipleDataSources = $DataSourceConfig -is [System.Array]

    foreach ($ds in $dataSources.value) {
        $dsUpdate = @{}
        $matchedConfig = $null

        if ($isMultipleDataSources) {
            # Match by datasource name or index
            $matchedConfig = $DataSourceConfig | Where-Object {
                ($_.Name -and $_.Name -eq $ds.Name) -or
                ($_.Id -and $_.Id -eq $ds.Id)
            } | Select-Object -First 1

            # If no match found by name/id, skip this datasource
            if (-not $matchedConfig) {
                continue
            }
        }
        else {
            # Single datasource config - apply to all datasources
            $matchedConfig = $DataSourceConfig
        }

        # Build connection string based on type
        if ($matchedConfig.Server -or $matchedConfig.ConnectionString) {
            if ($matchedConfig.ConnectionString) {
                # Use provided connection string directly
                $dsUpdate.ConnectionString = $matchedConfig.ConnectionString
            }
            elseif ($matchedConfig.Server) {
                # Build SQL Server connection string
                if ($matchedConfig.Database) {
                    $dsUpdate.ConnectionString = "Data Source=$($matchedConfig.Server);Initial Catalog=$($matchedConfig.Database)"
                }
                else {
                    $dsUpdate.ConnectionString = "Data Source=$($matchedConfig.Server)"
                }
            }
        }

        if ($matchedConfig.CredentialType) {
            $dsUpdate.CredentialRetrieval = $matchedConfig.CredentialType
        }

        if ($matchedConfig.Username -and $matchedConfig.Password) {
            $dsUpdate.CredentialsByUser = @{
                UserName = $matchedConfig.Username
                Password = $matchedConfig.Password
            }
        }

        # Handle Windows credentials
        if ($matchedConfig.CredentialType -eq "Integrated" -or $matchedConfig.CredentialType -eq "Windows") {
            $dsUpdate.CredentialRetrieval = "Integrated"
            # Remove any stored credentials
            $dsUpdate.CredentialsByUser = $null
        }

        if ($dsUpdate.Count -gt 0) {
            $dsId = $ds.Id
            Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/DataSources($dsId)" -Method PATCH -Body $dsUpdate | Out-Null
        }
    }
}

function Set-ReportPermissions {
    param(
        [string]$BaseUrl,
        [string]$ReportPath,
        [object]$Permissions,
        [switch]$Replace
    )

    $encodedPath = [System.Web.HttpUtility]::UrlEncode($ReportPath)

    # Convert single permission object to array
    $permArray = if ($Permissions -is [System.Array]) {
        $Permissions
    } else {
        @($Permissions)
    }

    # Get existing permissions if not replacing
    $existingPolicies = @()
    if (-not $Replace) {
        try {
            $existingPermissions = Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/Policies" -Method GET
            $existingPolicies = $existingPermissions.Policies
        }
        catch {
            # No existing permissions or error - continue with new ones only
        }
    }

    # Build new policies
    $newPolicies = @()
    foreach ($perm in $permArray) {
        $newPolicies += @{
            GroupUserName = $perm.User
            Roles = @(@{ Name = $perm.Role })
        }
    }

    # Merge: Add existing policies that are not being overridden
    $finalPolicies = @()
    $newUserRoles = @{}

    # Track new permissions by user
    foreach ($policy in $newPolicies) {
        $newUserRoles[$policy.GroupUserName] = $policy
        $finalPolicies += $policy
    }

    # Add existing permissions for users not in the new list
    if (-not $Replace) {
        foreach ($existingPolicy in $existingPolicies) {
            if (-not $newUserRoles.ContainsKey($existingPolicy.GroupUserName)) {
                $finalPolicies += $existingPolicy
            }
        }
    }

    $body = @{ Policies = $finalPolicies }
    Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/Policies" -Method PUT -Body $body | Out-Null
}

function Update-ReportParameters {
    param(
        [string]$BaseUrl,
        [string]$ReportPath,
        [hashtable]$Parameters
    )

    $encodedPath = [System.Web.HttpUtility]::UrlEncode($ReportPath)

    # Get current parameters
    $currentParams = Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/DataModelParameters" -Method GET

    $updates = @()
    foreach ($param in $currentParams.value) {
        if ($Parameters.ContainsKey($param.Name)) {
            $updates += @{
                Name = $param.Name
                NewValue = $Parameters[$param.Name]
            }
        }
    }

    if ($updates.Count -gt 0) {
        Invoke-RSWebRequest -Uri "$BaseUrl/PowerBIReports(Path='$encodedPath')/DataModelParameters" -Method POST -Body @{ Parameters = $updates } | Out-Null
    }
}

#endregion

#region Main Execution

try {
    # Load configuration
    if (-not (Test-Path $ConfigPath)) {
        throw "Configuration file not found: $ConfigPath"
    }

    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

    # Add System.Web assembly for URL encoding
    Add-Type -AssemblyName System.Web

    # Build base URL
    $baseUrl = "$($config.ReportServerUrl)/reports/api/v2.0"

    Write-Host "`n=== Power BI OnPrem Deployment ===" -ForegroundColor Cyan
    Write-Host "Server: $($config.ReportServerUrl)" -ForegroundColor Gray
    Write-Host "Source: $($config.SourcePath)" -ForegroundColor Gray
    Write-Host ""

    # Get all .pbix files
    if (-not (Test-Path $config.SourcePath)) {
        throw "Source path not found: $($config.SourcePath)"
    }

    $reports = Get-ChildItem -Path $config.SourcePath -Filter "*.pbix" -File

    if ($reports.Count -eq 0) {
        Write-Host "No .pbix files found in source path." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "Found $($reports.Count) report(s) to deploy`n" -ForegroundColor Cyan

    # Deploy each report
    $deployed = 0
    $skipped = 0
    $failed = 0

    foreach ($report in $reports) {
        $reportNameWithoutExt = $report.BaseName

        Write-Host "Processing: $($report.Name)" -ForegroundColor White

        try {
            # Determine target path using priority resolution
            $target = Get-ReportTargetPath `
                -ReportName $reportNameWithoutExt `
                -ReportMappings $config.ReportMappings `
                -DefaultTargetFolder $config.TargetFolderPath `
                -DefaultReportName $reportNameWithoutExt

            # Create folder if needed
            New-RSFolder -BaseUrl $baseUrl -Path $target.Path

            # Publish report
            $publishedReport = Publish-PowerBIReport `
                -BaseUrl $baseUrl `
                -FilePath $report.FullName `
                -TargetPath $target.Path `
                -ReportName $target.Name `
                -Overwrite $config.OverwriteExisting

            if ($null -eq $publishedReport) {
                $skipped++
                continue
            }

            $fullPath = "$($target.Path)/$($target.Name)"

            # Find matching configuration
            $reportConfig = $config.ReportMappings | Where-Object { $reportNameWithoutExt -match $_.Pattern } | Select-Object -First 1

            # Update datasources (report-specific or global)
            # Support both "DataSource" and "GlobalDataSource" for backward compatibility
            $dsConfig = if ($reportConfig -and $reportConfig.DataSource) {
                $reportConfig.DataSource
            } elseif ($config.DataSource) {
                $config.DataSource
            } elseif ($config.GlobalDataSource) {
                $config.GlobalDataSource
            } else {
                $null
            }

            if ($dsConfig) {
                Write-Host "  [CONFIG] Updating datasources" -ForegroundColor Gray
                Update-DataSource -BaseUrl $baseUrl -ReportPath $fullPath -DataSourceConfig $dsConfig
            }

            # Set permissions (report-specific or global)
            # Support both "Permissions" and "GlobalPermissions" for backward compatibility
            $perms = if ($reportConfig -and $reportConfig.Permissions) {
                $reportConfig.Permissions
            } elseif ($config.Permissions) {
                $config.Permissions
            } elseif ($config.GlobalPermissions) {
                $config.GlobalPermissions
            } else {
                $null
            }

            if ($perms) {
                # Check if we should replace or merge permissions
                $replacePerms = $false
                if ($reportConfig -and $reportConfig.PSObject.Properties['ReplacePermissions']) {
                    $replacePerms = $reportConfig.ReplacePermissions
                }
                elseif ($config.PSObject.Properties['ReplacePermissions']) {
                    $replacePerms = $config.ReplacePermissions
                }

                $action = if ($replacePerms) { "Replacing" } else { "Merging" }
                Write-Host "  [CONFIG] $action permissions" -ForegroundColor Gray
                Set-ReportPermissions -BaseUrl $baseUrl -ReportPath $fullPath -Permissions $perms -Replace:$replacePerms
            }

            # Update parameters (report-specific or global)
            # Support both "Parameters" and "GlobalParameters" for backward compatibility
            $params = if ($reportConfig -and $reportConfig.Parameters) {
                $reportConfig.Parameters
            } elseif ($config.Parameters) {
                $config.Parameters
            } elseif ($config.GlobalParameters) {
                $config.GlobalParameters
            } else {
                $null
            }

            if ($params) {
                Write-Host "  [CONFIG] Updating parameters" -ForegroundColor Gray
                Update-ReportParameters -BaseUrl $baseUrl -ReportPath $fullPath -Parameters $params
            }

            $deployed++
            Write-Host ""
        }
        catch {
            $failed++
            Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
        }
    }

    # Summary
    Write-Host "=== Deployment Summary ===" -ForegroundColor Cyan
    Write-Host "Deployed: $deployed" -ForegroundColor Green
    Write-Host "Skipped:  $skipped" -ForegroundColor Yellow
    Write-Host "Failed:   $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Gray" })
    Write-Host ""

    if ($failed -gt 0) {
        exit 1
    }
}
catch {
    Write-Host "`n[FATAL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

#endregion
