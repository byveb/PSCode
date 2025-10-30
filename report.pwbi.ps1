# Flexible Power BI Report Deployment Script
# Supports global and report-specific configurations with inheritance
# No external modules required

<#
.SYNOPSIS
    Flexible deployment solution for Power BI reports with advanced configuration support
    
.DESCRIPTION
    This script performs complete deployment with:
    1. Auto-discovery of reports from FolderPath
    2. Global configurations (DataSources, Parameters, Permissions)
    3. Report-specific overrides with inheritance
    4. Pattern-based ReportMappings for custom routing
    5. Permission management
    6. Auto-discovery mode: Uses environment variables when config file not provided
    
.PARAMETER ConfigFile
    Path to JSON configuration file. If not provided, will use environment variables.
    
.PARAMETER ServerUrl
    Power BI Report Server URL (required when not using config file)
    
.PARAMETER TargetFolder
    Target folder path (required when not using config file)
    
.PARAMETER ReportPath
    Local path to .pbix files (required when not using config file)
    
.EXAMPLE
    .\Deploy-PowerBI-Flexible.ps1 -ConfigFile "config\config-flexible.json"
    Deploy with config file
    
.EXAMPLE
    .\Deploy-PowerBI-Flexible.ps1 -ServerUrl "http://server:8080/PWBIReports" -TargetFolder "/Reports" -ReportPath "E:\PWBI"
    Deploy using environment variables for Parameters and DataSources
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = "",
    
    [Parameter(Mandatory = $false)]
    [string]$ServerUrl = "",
    
    [Parameter(Mandatory = $false)]
    [string]$TargetFolder = "",
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ""
)

#region Helper Functions

function Write-Log {
    param([string]$Message, [string]$Level = "Info")
    $color = switch ($Level) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        "Cyan" { "Cyan" }
        default { "White" }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function New-FolderIfNotExists {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$FolderPath
    )
    
    $baseUri = "$ServerUrl$ReportPath"
    
    # Split path into parts and create each level
    $pathParts = $FolderPath -split '/' | Where-Object { $_ -ne '' }
    $currentPath = ""
    
    foreach ($part in $pathParts) {
        $currentPath = "$currentPath/$part"
        
        # Check if folder exists
        $checkUri = "$baseUri/api/v2.0/Folders(Path='$currentPath')"
        
        try {
            $null = Invoke-RestMethod -Uri $checkUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication -ErrorAction Stop
            Write-Log "  Folder exists: $currentPath" -Level Info
        }
        catch {
            # Folder doesn't exist, create it
            Write-Log "  Creating folder: $currentPath" -Level Warning
            
            $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
            if ($parentPath -eq '') { $parentPath = '/' }
            
            $createUri = "$baseUri/api/v2.0/Folders"
            $body = @{
                Path = $parentPath
                Name = $part
            } | ConvertTo-Json
            
            try {
                Invoke-RestMethod -Uri $createUri -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ContentType 'application/json' -UseDefaultCredentials -AllowUnencryptedAuthentication | Out-Null
                Write-Log "  ✓ Folder created: $currentPath" -Level Success
            }
            catch {
                Write-Log "  ERROR creating folder: $($_.Exception.Message)" -Level Error
                throw
            }
        }
    }
}

function Get-ReportsFromFolder {
    param(
        [string]$FolderPath,
        [bool]$Recurse = $false
    )
    
    if (-not (Test-Path $FolderPath)) {
        Write-Log "ERROR: Folder not found: $FolderPath" -Level Error
        return @()
    }
    
    if ($Recurse) {
        $reports = Get-ChildItem -Path $FolderPath -Filter "*.pbix" -File -Recurse
    } else {
        $reports = Get-ChildItem -Path $FolderPath -Filter "*.pbix" -File
    }
    
    Write-Log "Found $($reports.Count) report(s) in: $FolderPath $(if($Recurse){'(including subfolders)'})" -Level Info
    
    return $reports
}

function Get-ReportMapping {
    param(
        [string]$ReportFileName,
        [array]$ReportMappings
    )
    
    foreach ($mapping in $ReportMappings) {
        if ($ReportFileName -match $mapping.SourcePattern) {
            Write-Log "  ✓ Matched pattern: $($mapping.SourcePattern)" -Level Success
            return $mapping
        }
    }
    
    return $null
}

function Merge-Configuration {
    param(
        [array]$GlobalConfig,
        [array]$SpecificConfig,
        [string]$ReportName,
        [string]$ReportRelativePath,  # New: relative path like "subfolder\Sales.pbix"
        [string]$ConfigType
    )
    
    $merged = @()
    
    # Add global configurations (those without "Reports" property)
    foreach ($item in $GlobalConfig) {
        if (-not $item.Reports) {
            $merged += $item
        }
        else {
            # Check if any of the report names match
            $matched = $false
            foreach ($configReport in $item.Reports) {
                # Extract filename from config value
                $configFileName = if ($configReport -match '[/\\]') {
                    Split-Path -Path $configReport -Leaf
                } else {
                    $configReport
                }
                
                # Match by filename
                if ($configFileName -eq $ReportName) {
                    $matched = $true
                    break
                }
                
                # Also try matching with relative path (e.g., "subfolder/Sales.pbix" or "subfolder\Sales.pbix")
                if ($ReportRelativePath) {
                    $normalizedConfigPath = $configReport -replace '/', '\\'
                    $normalizedReportPath = $ReportRelativePath -replace '/', '\\'
                    if ($normalizedConfigPath -eq $normalizedReportPath) {
                        $matched = $true
                        break
                    }
                }
            }
            
            if ($matched) {
                $merged += $item
            }
        }
    }
    
    # Add mapping-specific configurations
    if ($SpecificConfig) {
        foreach ($item in $SpecificConfig) {
            $merged += $item
        }
    }
    
    return $merged
}

function Deploy-Report {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$LocalFilePath,
        [string]$TargetFolder,
        [string]$ReportName
    )
    
    $baseUri = "$ServerUrl$ReportPath"
    
    # Read file
    if (-not (Test-Path $LocalFilePath)) {
        Write-Log "ERROR: File not found: $LocalFilePath" -Level Error
        return $false
    }
    
    $fileBytes = [System.IO.File]::ReadAllBytes($LocalFilePath)
    $reportBaseName = [System.IO.Path]::GetFileNameWithoutExtension($LocalFilePath)
    
    # Build item path for the report
    $itemPath = if ($TargetFolder -eq "/") { "/$reportBaseName" } else { "$TargetFolder/$reportBaseName" }
    
    $payload = @{
        "@odata.type" = "#Model.PowerBIReport"
        "Content"     = [System.Convert]::ToBase64String($fileBytes)
        "ContentType" = ""
        "Name"        = $reportBaseName
        "Path"        = $itemPath
    }
    
    $payloadJson = ConvertTo-Json $payload
    $body = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
    
    # Upload to CatalogItems endpoint
    $uploadUri = "$baseUri/api/v2.0/CatalogItems"
    
    try {
        Invoke-RestMethod -Uri $uploadUri -Method Post -Body $body -ContentType "application/json" -UseDefaultCredentials -AllowUnencryptedAuthentication
        Write-Log "Report uploaded: $ReportName" -Level Success
        return $true
    }
    catch {
        # Check if it actually succeeded despite error
        $reportFullPath = "$TargetFolder/$ReportName"
        $encodedReportPath = [Uri]::EscapeDataString($reportFullPath)
        $checkUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedReportPath')"
        try {
            $null = Invoke-RestMethod -Uri $checkUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication
            Write-Log "Report uploaded (verified): $ReportName" -Level Success
            return $true
        }
        catch {
            Write-Log "ERROR uploading report: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
}

function Update-DataSourceCredentials {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$ReportFullPath,
        [array]$ConfiguredDataSources
    )
    
    # Validate Order values in configured data sources
    $orderBasedConfigs = $ConfiguredDataSources | Where-Object { $_.Order }
    if ($orderBasedConfigs.Count -gt 0) {
        # Check for duplicate Order values
        $orderGroups = $orderBasedConfigs | Group-Object -Property Order
        $duplicates = $orderGroups | Where-Object { $_.Count -gt 1 }
        if ($duplicates) {
            $dupeOrders = ($duplicates | ForEach-Object { $_.Name }) -join ', '
            Write-Log "  ERROR: Duplicate Order values found in DataSources config: $dupeOrders" -Level Error
            Write-Log "  Each data source must have a unique Order value. Skipping credential update." -Level Error
            return
        }
        
        # Check for invalid Order values (must be positive integers)
        $invalidOrders = $orderBasedConfigs | Where-Object { $_.Order -lt 1 }
        if ($invalidOrders) {
            Write-Log "  ERROR: Order values must be positive integers (1, 2, 3, etc.)" -Level Error
            Write-Log "  Found invalid Order: $($invalidOrders.Order -join ', '). Skipping credential update." -Level Error
            return
        }
    }
    
    $baseUri = "$ServerUrl$ReportPath"
    $encodedPath = [Uri]::EscapeDataString($ReportFullPath)
    $getUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')?`$expand=DataSources"
    
    try {
        # Get current data sources
        $response = Invoke-RestMethod -Uri $getUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication -ContentType 'application/json'
        $dataSources = $response.DataSources
        
        if (-not $dataSources -or $dataSources.Count -eq 0) {
            Write-Log "  No data sources found" -Level Warning
            return
        }
        
        Write-Log "  Found $($dataSources.Count) data source(s)" -Level Info
        
        # Validate Order values against actual data source count
        $maxOrder = ($orderBasedConfigs | Measure-Object -Property Order -Maximum).Maximum
        if ($maxOrder -gt $dataSources.Count) {
            Write-Log "  ERROR: Order $maxOrder specified in config, but report only has $($dataSources.Count) data source(s)" -Level Error
            Write-Log "  Valid Order values are 1 to $($dataSources.Count). Skipping credential update." -Level Error
            return
        }
        
        # Update credentials
        $updated = $false
        foreach ($configDS in $ConfiguredDataSources) {
            # Support both ConnectionString field and ServerName/DatabaseName fields
            $connectionString = if ($configDS.ConnectionString) { 
                $configDS.ConnectionString 
            } else { 
                "$($configDS.ServerName);$($configDS.DatabaseName)" 
            }
            
            # Try to match by Name first, then by connection string, then by Order
            $ds = $null
            
            if ($configDS.Name) {
                # Match by Name property
                $ds = $dataSources | Where-Object { $_.Name -eq $configDS.Name } | Select-Object -First 1
                if ($ds) {
                    Write-Log "    Matched by Name: $($configDS.Name)" -Level Info
                }
            }
            
            if (-not $ds -and $configDS.Order) {
                # Match by Order (1-based index)
                $dsIndex = $configDS.Order - 1
                if ($dsIndex -ge 0 -and $dsIndex -lt $dataSources.Count) {
                    $ds = $dataSources[$dsIndex]
                    Write-Log "    Matched by Order: $($configDS.Order)" -Level Info
                }
            }
            
            if (-not $ds) {
                # Match by connection string
                $ds = $dataSources | Where-Object { $_.ConnectionString -eq $connectionString } | Select-Object -First 1
                if ($ds) {
                    Write-Log "    Matched by ConnectionString: $connectionString" -Level Info
                }
            }
            
            if (-not $ds) {
                Write-Log "    ⚠ No matching data source for: $connectionString" -Level Warning
                continue
            }
            
            # Initialize DataModelDataSource if needed
            if (-not $ds.DataModelDataSource) {
                $ds.DataModelDataSource = @{}
            }
            
            # Update based on CredentialType/AuthType
            $authType = if ($configDS.CredentialType) { $configDS.CredentialType } else { $configDS.AuthType }
            
            switch ($authType) {
                "Windows" {
                    $ds.DataModelDataSource.AuthType = 'Windows'
                    if ($configDS.Username) {
                        $ds.DataModelDataSource.Username = $configDS.Username
                        $ds.DataModelDataSource.Secret = $configDS.Password
                    }
                    else {
                        $ds.DataModelDataSource.PSObject.Properties.Remove('Username')
                        $ds.DataModelDataSource.PSObject.Properties.Remove('Secret')
                    }
                }
                { $_ -in @("UsernamePassword", "SQL") } {
                    $ds.DataModelDataSource.AuthType = 'UsernamePassword'
                    $ds.DataModelDataSource.Username = $configDS.Username
                    $ds.DataModelDataSource.Secret = $configDS.Password
                }
                "Key" {
                    $ds.DataModelDataSource.AuthType = 'Key'
                    $ds.DataModelDataSource.Secret = $configDS.Key
                }
            }
            
            $updated = $true
            $dsName = if ($configDS.Name) { $configDS.Name } else { $ds.ConnectionString }
            Write-Log "    ✓ Configured $dsName (AuthType: $($ds.DataModelDataSource.AuthType))" -Level Success
        }
        
        if (-not $updated) {
            Write-Log "  No data sources updated" -Level Warning
            return
        }
        
        # Send PATCH request (Microsoft's pattern for PowerBI reports)
        # Reference: Set-RsRestItemDataSource uses PATCH for PowerBIReport
        $patchUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')/DataSources"
        $dataSourcesArray = @($dataSources)
        $payloadJson = ConvertTo-Json -InputObject $dataSourcesArray -Depth 10
        
        Write-Log "  → Sending PATCH to: $patchUri" -Level Info
        Write-Log "  → Payload (first 500 chars): $($payloadJson.Substring(0, [Math]::Min(500, $payloadJson.Length)))" -Level Info
        
        $patchParams = @{
            Uri                            = $patchUri
            Method                         = 'Patch'
            Body                           = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
            ContentType                    = 'application/json; charset=utf-8'
            UseDefaultCredentials          = $true
            AllowUnencryptedAuthentication = $true
        }
        
        Invoke-RestMethod @patchParams | Out-Null
        Write-Log "  ✓ Credentials updated" -Level Success
    }
    catch {
        $errorDetail = ""
        if ($_.ErrorDetails.Message) {
            try {
                $errorObj = $_.ErrorDetails.Message | ConvertFrom-Json
                $errorDetail = " | Details: $($errorObj.error.message)"
            } catch {}
        }
        Write-Log "  ERROR updating credentials: $($_.Exception.Message)$errorDetail" -Level Error
    }
}

function Update-ReportParameters {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$ReportFullPath,
        [array]$ConfiguredParameters
    )
    
    if (-not $ConfiguredParameters -or $ConfiguredParameters.Count -eq 0) {
        return
    }
    
    $baseUri = "$ServerUrl$ReportPath"
    $encodedPath = [Uri]::EscapeDataString($ReportFullPath)
    $getUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')?`$expand=DataModelParameters"
    
    try {
        # Get current parameters
        $response = Invoke-RestMethod -Uri $getUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication
        
        if (-not $response.DataModelParameters -or $response.DataModelParameters.Count -eq 0) {
            Write-Log "  No parameters found in report" -Level Info
            return
        }
        
        $parameters = $response.DataModelParameters
        Write-Log "  Found $($parameters.Count) parameter(s)" -Level Info
        
        # Check parameter values and prepare updates
        $needsUpdate = $false
        foreach ($param in $parameters) {
            $configParam = $ConfiguredParameters | Where-Object { $_.Name -eq $param.Name }
            if ($configParam) {
                if ($param.Value -eq $configParam.Value) {
                    Write-Log "    ✓ $($param.Name) = $($param.Value) (already correct)" -Level Success
                }
                else {
                    Write-Log "    ⚠ $($param.Name) = $($param.Value) (will update to: $($configParam.Value))" -Level Warning
                    $param.Value = $configParam.Value
                    $needsUpdate = $true
                }
            }
        }
        
        # Warn about parameters defined in config but not found in report
        foreach ($configParam in $ConfiguredParameters) {
            $reportParam = $parameters | Where-Object { $_.Name -eq $configParam.Name }
            if (-not $reportParam) {
                Write-Log "    ⚠ Parameter '$($configParam.Name)' defined in config but NOT FOUND in report" -Level Warning
            }
        }
        
        # Update parameters if needed
        if ($needsUpdate) {
            Write-Log "  Updating parameters..." -Level Info
            
            $parametersArray = @($parameters)
            $payloadJson = ConvertTo-Json -InputObject $parametersArray -Depth 3
            
            $postUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')/DataModelParameters"
            $postParams = @{
                Uri                            = $postUri
                Method                         = 'Post'
                Body                           = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
                ContentType                    = 'application/json; charset=utf-8'
                UseDefaultCredentials          = $true
                AllowUnencryptedAuthentication = $true
            }
            
            Invoke-RestMethod @postParams | Out-Null
            Write-Log "  ✓ Parameters updated successfully" -Level Success
        }
    }
    catch {
        Write-Log "  ERROR with parameters: $($_.Exception.Message)" -Level Error
    }
}

function Update-ReportParametersFromEnvironment {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$ReportFullPath,
        [string]$ReportName
    )
    
    $baseUri = "$ServerUrl$ReportPath"
    $encodedPath = [Uri]::EscapeDataString($ReportFullPath)
    $getUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')?`$expand=DataModelParameters"
    
    try {
        # Get current parameters from the report
        $response = Invoke-RestMethod -Uri $getUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication
        
        if (-not $response.DataModelParameters -or $response.DataModelParameters.Count -eq 0) {
            Write-Log "  No parameters found in report" -Level Info
            return
        }
        
        $parameters = $response.DataModelParameters
        Write-Log "  Found $($parameters.Count) parameter(s) in report" -Level Info
        
        # Check each parameter against environment variables with hierarchical lookup
        # Priority: 1) ReportName_ParameterName, 2) ParameterName
        $needsUpdate = $false
        foreach ($param in $parameters) {
            # Try report-specific variable first (e.g., Sales_ServerName)
            $reportSpecificVar = "${ReportName}_$($param.Name)"
            $envValue = [Environment]::GetEnvironmentVariable($reportSpecificVar)
            $source = "report-specific"
            
            # Fallback to shared variable (e.g., ServerName)
            if (-not $envValue) {
                $envValue = [Environment]::GetEnvironmentVariable($param.Name)
                $source = "shared"
            }
            
            if ($envValue) {
                if ($param.Value -eq $envValue) {
                    Write-Log "    ✓ $($param.Name) = $($param.Value) (matches $source env var)" -Level Success
                }
                else {
                    Write-Log "    ↻ $($param.Name): '$($param.Value)' → '$envValue' (from $source env var)" -Level Warning
                    $param.Value = $envValue
                    $needsUpdate = $true
                }
            }
            else {
                Write-Log "    - $($param.Name) = $($param.Value) (no env var, keeping current)" -Level Info
            }
        }
        
        # Update parameters if any changes
        if ($needsUpdate) {
            Write-Log "  Updating parameters from environment variables..." -Level Info
            
            $parametersArray = @($parameters)
            $payloadJson = ConvertTo-Json -InputObject $parametersArray -Depth 3
            
            $postUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')/DataModelParameters"
            $postParams = @{
                Uri                            = $postUri
                Method                         = 'Post'
                Body                           = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
                ContentType                    = 'application/json; charset=utf-8'
                UseDefaultCredentials          = $true
                AllowUnencryptedAuthentication = $true
            }
            
            Invoke-RestMethod @postParams | Out-Null
            Write-Log "  ✓ Parameters updated from environment" -Level Success
        }
        else {
            Write-Log "  No parameter updates needed" -Level Info
        }
    }
    catch {
        Write-Log "  ERROR with parameters: $($_.Exception.Message)" -Level Error
    }
}

function Update-DataSourceCredentialsFromEnvironment {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$ReportFullPath,
        [string]$ReportName
    )
    
    $baseUri = "$ServerUrl$ReportPath"
    $encodedPath = [Uri]::EscapeDataString($ReportFullPath)
    $getUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')?`$expand=DataSources"
    
    try {
        # Get current data sources
        $response = Invoke-RestMethod -Uri $getUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication -ContentType 'application/json'
        $dataSources = $response.DataSources
        
        if (-not $dataSources -or $dataSources.Count -eq 0) {
            Write-Log "  No data sources found" -Level Warning
            return
        }
        
        Write-Log "  Found $($dataSources.Count) data source(s)" -Level Info
        
        # Check for environment variables for data sources with hierarchical lookup
        # Priority: 1) ReportName_DS<Order>_*, 2) DS<Order>_*
        $updated = $false
        for ($i = 0; $i -lt $dataSources.Count; $i++) {
            $ds = $dataSources[$i]
            $dsOrder = $i + 1
            
            Write-Log "    Checking DataSource $dsOrder..." -Level Info
            
            # Try report-specific variables first (e.g., Sales_DS1_CredentialType)
            $reportPrefix = "${ReportName}_DS$dsOrder"
            $sharedPrefix = "DS$dsOrder"
            
            $credentialType = [Environment]::GetEnvironmentVariable("${reportPrefix}_CredentialType")
            $varSource = "report-specific"
            
            if (-not $credentialType) {
                $credentialType = [Environment]::GetEnvironmentVariable("${sharedPrefix}_CredentialType")
                $varSource = "shared"
            }
            
            if (-not $credentialType) {
                Write-Log "      - No ${reportPrefix}_CredentialType or ${sharedPrefix}_CredentialType env var found" -Level Info
                continue
            }
            
            # Get username and password with same hierarchy
            $username = [Environment]::GetEnvironmentVariable("${reportPrefix}_Username")
            if (-not $username) {
                $username = [Environment]::GetEnvironmentVariable("${sharedPrefix}_Username")
            }
            
            $password = [Environment]::GetEnvironmentVariable("${reportPrefix}_Password")
            if (-not $password) {
                $password = [Environment]::GetEnvironmentVariable("${sharedPrefix}_Password")
            }
            
            # Initialize DataModelDataSource if needed
            if (-not $ds.DataModelDataSource) {
                $ds.DataModelDataSource = @{}
            }
            
            switch ($credentialType) {
                "Windows" {
                    $ds.DataModelDataSource.AuthType = 'Windows'
                    if ($username -and $password) {
                        $ds.DataModelDataSource.Username = $username
                        $ds.DataModelDataSource.Secret = $password
                        Write-Log "      ✓ Windows auth with credentials (from $varSource env)" -Level Success
                    }
                    else {
                        Write-Log "      ✓ Windows auth (no credentials in env)" -Level Success
                    }
                    $updated = $true
                }
                { $_ -in @("UsernamePassword", "SQL") } {
                    if ($username -and $password) {
                        $ds.DataModelDataSource.AuthType = 'UsernamePassword'
                        $ds.DataModelDataSource.Username = $username
                        $ds.DataModelDataSource.Secret = $password
                        Write-Log "      ✓ SQL auth configured (from $varSource env)" -Level Success
                        $updated = $true
                    }
                    else {
                        Write-Log "      ⚠ SQL auth requires username and password env vars" -Level Warning
                    }
                }
                "Key" {
                    # Check for Key with hierarchy
                    $key = [Environment]::GetEnvironmentVariable("${reportPrefix}_Key")
                    if (-not $key) {
                        $key = [Environment]::GetEnvironmentVariable("${sharedPrefix}_Key")
                    }
                    
                    if ($key) {
                        $ds.DataModelDataSource.AuthType = 'Key'
                        $ds.DataModelDataSource.Secret = $key
                        Write-Log "      ✓ Key auth configured (from $varSource env)" -Level Success
                        $updated = $true
                    }
                    else {
                        Write-Log "      ⚠ Key auth requires Key env var" -Level Warning
                    }
                }
            }
        }
        
        if ($updated) {
            # Send PATCH request
            $patchUri = "$baseUri/api/v2.0/PowerBIReports(Path='$encodedPath')/DataSources"
            $dataSourcesArray = @($dataSources)
            $payloadJson = ConvertTo-Json -InputObject $dataSourcesArray -Depth 10
            
            $patchParams = @{
                Uri                            = $patchUri
                Method                         = 'Patch'
                Body                           = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
                ContentType                    = 'application/json; charset=utf-8'
                UseDefaultCredentials          = $true
                AllowUnencryptedAuthentication = $true
            }
            
            Invoke-RestMethod @patchParams | Out-Null
            Write-Log "  ✓ Credentials updated from environment" -Level Success
        }
        else {
            Write-Log "  No data source updates from environment" -Level Info
        }
    }
    catch {
        Write-Log "  ERROR updating credentials: $($_.Exception.Message)" -Level Error
    }
}

function Set-ReportPermissions {
    param(
        [string]$ServerUrl,
        [string]$ReportPath,
        [string]$ReportFullPath,
        [array]$ConfiguredPermissions
    )
    
    if (-not $ConfiguredPermissions -or $ConfiguredPermissions.Count -eq 0) {
        return
    }
    
    $baseUri = "$ServerUrl$ReportPath"
    $encodedPath = [Uri]::EscapeDataString($ReportFullPath)
    
    Write-Log "  Configuring $($ConfiguredPermissions.Count) permission(s)" -Level Info
    
    try {
        # Step 1: Get Item ID using Path (Microsoft's pattern)
        Write-Log "    → Getting item ID..." -Level Info
        $getItemUri = "$baseUri/api/v2.0/CatalogItems(Path='$encodedPath')"
        $itemResponse = Invoke-RestMethod -Uri $getItemUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication
        $itemId = $itemResponse.Id
        
        if (-not $itemId) {
            Write-Log "  ⚠ Could not retrieve item ID" -Level Warning
            return
        }
        
        Write-Log "    → Item ID: $itemId" -Level Info
        
        # Step 2: GET current policies using Item ID (Microsoft's pattern)
        $getPoliciesUri = "$baseUri/api/v2.0/CatalogItems($itemId)/Policies"
        $policiesResponse = Invoke-RestMethod -Uri $getPoliciesUri -Method Get -UseDefaultCredentials -AllowUnencryptedAuthentication
        
        $existingPolicies = $policiesResponse.Policies
        if (-not $existingPolicies) {
            $existingPolicies = @()
        }
        
        Write-Log "    → Found $($existingPolicies.Count) existing policy/policies" -Level Info
        
        # Step 3: Merge new permissions with existing ones (Microsoft's pattern)
        foreach ($perm in $ConfiguredPermissions) {
            # Support both single role (string) and multiple roles (array)
            $roles = if ($perm.Role -is [array]) { $perm.Role } else { @($perm.Role) }
            $roleDisplay = $roles -join ", "
            
            # Check if permission already exists
            $existing = $existingPolicies | Where-Object { $_.GroupUserName -eq $perm.Principal }
            
            if ($existing) {
                # Build role objects with Name and Description
                $roleObjects = @()
                foreach ($role in $roles) {
                    $roleObjects += @{
                        Name = $role
                        Description = ''
                    }
                }
                
                $existing.Roles = $roleObjects
                Write-Log "    ↻ $($perm.Principal) -> $roleDisplay (updated)" -Level Info
            }
            else {
                # Build role objects with Name and Description (Microsoft's format)
                $roleObjects = @()
                foreach ($role in $roles) {
                    $roleObjects += @{
                        Name = $role
                        Description = ''
                    }
                }
                
                # Add new policy
                $newPolicy = @{
                    GroupUserName = $perm.Principal
                    Roles = $roleObjects
                }
                $existingPolicies += $newPolicy
                Write-Log "    + $($perm.Principal) -> $roleDisplay (added)" -Level Success
            }
        }
        
        # Step 4: PUT the complete policy object back (Microsoft's pattern)
        $policyPayload = @{
            Policies = $existingPolicies
            InheritParentPolicy = $false
        }
        
        $payloadJson = ConvertTo-Json $policyPayload -Depth 15
        $body = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
        
        # Use PUT method with Item ID (Microsoft's pattern)
        $putPoliciesUri = "$baseUri/api/v2.0/CatalogItems($itemId)/Policies"
        Invoke-RestMethod -Uri $putPoliciesUri -Method Put -Body $body -ContentType "application/json; charset=utf-8" -UseDefaultCredentials -AllowUnencryptedAuthentication | Out-Null
        
        Write-Log "  ✓ Permissions updated successfully" -Level Success
    }
    catch {
        Write-Log "  ⚠ Failed to set permissions: $($_.Exception.Message)" -Level Warning
        Write-Log "    Details: $($_.Exception.Response.StatusCode) - $($_.Exception.Response.StatusDescription)" -Level Warning
    }
}

#endregion

#region Main Script

Write-Log "========================================" -Level Cyan
Write-Log "Power BI Flexible Deployment" -Level Cyan
Write-Log "========================================" -Level Cyan
Write-Log "" -Level Info

# Determine deployment mode
$useEnvironmentVariables = $false
$config = $null

if ($ConfigFile -and (Test-Path $ConfigFile)) {
    # Config file mode
    Write-Log "Loading configuration: $ConfigFile" -Level Info
    $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    
    # Parse TargetServer to extract base URL and report path
    if ($config.TargetServer -match '^(https?://[^/]+)(/.+)?$') {
        $serverUrl = $matches[1]
        $reportPath = if ($matches[2]) { $matches[2] } else { "" }
    }
    else {
        $serverUrl = $config.TargetServer
        $reportPath = ""
    }
    
    $targetFolderPath = $config.TargetFolderPath
    $overwrite = $config.Overwrite
    $recurse = if ($null -ne $config.Recurse) { $config.Recurse } else { $false }
    $folderPath = $config.FolderPath
}
elseif ($ServerUrl -and $TargetFolder -and $ReportPath) {
    # Environment variable mode
    Write-Log "Using environment variable mode (no config file)" -Level Warning
    $useEnvironmentVariables = $true
    
    # Parse ServerUrl
    if ($ServerUrl -match '^(https?://[^/]+)(/.+)?$') {
        $serverUrl = $matches[1]
        $reportPath = if ($matches[2]) { $matches[2] } else { "" }
    }
    else {
        $serverUrl = $ServerUrl
        $reportPath = ""
    }
    
    $targetFolderPath = $TargetFolder
    $folderPath = $ReportPath
    $overwrite = $true
    $recurse = $false
}
else {
    Write-Log "ERROR: Must provide either -ConfigFile or (-ServerUrl, -TargetFolder, -ReportPath)" -Level Error
    Write-Log "  Config File Mode: -ConfigFile 'config\config.json'" -Level Error
    Write-Log "  Environment Mode: -ServerUrl 'http://server/reports' -TargetFolder '/Reports' -ReportPath 'E:\PWBI'" -Level Error
    exit 1
}

Write-Log "Server: $serverUrl$reportPath" -Level Info
Write-Log "Target Folder: $targetFolderPath" -Level Info
if ($useEnvironmentVariables) {
    Write-Log "Mode: Environment Variables" -Level Warning
}
Write-Log "" -Level Info

# Discover reports
$reports = Get-ReportsFromFolder -FolderPath $folderPath -Recurse $recurse
if ($reports.Count -eq 0) {
    Write-Log "No reports found to deploy" -Level Warning
    exit 0
}
Write-Log "" -Level Info

# Process each report
$successCount = 0
$failCount = 0

foreach ($reportFile in $reports) {
    $reportName = [System.IO.Path]::GetFileNameWithoutExtension($reportFile.Name)
    $reportFileName = $reportFile.Name
    
    # Calculate relative path from FolderPath (e.g., "subfolder\Sales.pbix" or just "Sales.pbix")
    $reportRelativePath = $reportFile.FullName.Substring($config.FolderPath.Length).TrimStart('\', '/')
    
    Write-Log "========================================" -Level Cyan
    Write-Log "Report: $reportFileName" -Level Cyan
    if ($reportRelativePath -ne $reportFileName) {
        Write-Log "  Location: $reportRelativePath" -Level Cyan
    }
    Write-Log "========================================" -Level Cyan
    Write-Log "" -Level Info
    
    try {
        # Check for report mapping
        $mapping = Get-ReportMapping -ReportFileName $reportFileName -ReportMappings $config.ReportMappings
        
        # Determine target path
        if ($mapping -and $mapping.ReportTargetPath) {
            # Full path specified
            $reportFullPath = $mapping.ReportTargetPath
            $folderPath = $reportFullPath.Substring(0, $reportFullPath.LastIndexOf('/'))
        }
        elseif ($mapping -and $mapping.TargetFolderPath) {
            # Folder override
            $folderPath = $mapping.TargetFolderPath
            $reportFullPath = "$folderPath/$reportName"
        }
        else {
            # Use default target folder
            $folderPath = $targetFolderPath
            $reportFullPath = "$folderPath/$reportName"
        }
        
        Write-Log "Target path: $reportFullPath" -Level Info
        Write-Log "" -Level Info
        
        # Step 1: Create folder
        Write-Log "Step 1: Ensuring folder exists..." -Level Cyan
        New-FolderIfNotExists -ServerUrl $serverUrl -ReportPath $reportPath -FolderPath $folderPath
        Write-Log "" -Level Info
        
        # Step 2: Deploy report
        Write-Log "Step 2: Deploying report..." -Level Cyan
        $deployed = Deploy-Report -ServerUrl $serverUrl -ReportPath $reportPath -LocalFilePath $reportFile.FullName -TargetFolder $folderPath -ReportName $reportName
        
        if (-not $deployed) {
            Write-Log "  Deployment failed, skipping further configuration" -Level Warning
            $failCount++
            Write-Log "" -Level Info
            continue
        }
        Write-Log "" -Level Info
        
        # Step 3: Update parameters
        Write-Log "Step 3: Checking report parameters..." -Level Cyan
        if ($useEnvironmentVariables) {
            # Environment variable mode: auto-discover parameters and match with env vars
            Update-ReportParametersFromEnvironment -ServerUrl $serverUrl -ReportPath $reportPath -ReportFullPath $reportFullPath -ReportName $reportName
        }
        else {
            # Config file mode: use merged configuration
            $mergedParameters = Merge-Configuration -GlobalConfig $config.Parameters -SpecificConfig $mapping.Parameters -ReportName $reportFileName -ReportRelativePath $reportRelativePath -ConfigType "Parameters"
            
            if ($mergedParameters.Count -gt 0) {
                Write-Log "  Applying $($mergedParameters.Count) parameter configuration(s)" -Level Info
                Update-ReportParameters -ServerUrl $serverUrl -ReportPath $reportPath -ReportFullPath $reportFullPath -ConfiguredParameters $mergedParameters
            }
            else {
                Write-Log "  No parameters configured" -Level Info
            }
        }
        Write-Log "" -Level Info
        
        # Step 4: Update data sources
        Write-Log "Step 4: Updating data source credentials..." -Level Cyan
        if ($useEnvironmentVariables) {
            # Environment variable mode: auto-discover data sources and match with env vars
            Update-DataSourceCredentialsFromEnvironment -ServerUrl $serverUrl -ReportPath $reportPath -ReportFullPath $reportFullPath -ReportName $reportName
        }
        else {
            # Config file mode: use merged configuration
            $mergedDataSources = Merge-Configuration -GlobalConfig $config.DataSources -SpecificConfig $mapping.DataSources -ReportName $reportFileName -ReportRelativePath $reportRelativePath -ConfigType "DataSources"
            
            if ($mergedDataSources.Count -gt 0) {
                Write-Log "  Applying $($mergedDataSources.Count) data source configuration(s)" -Level Info
                Update-DataSourceCredentials -ServerUrl $serverUrl -ReportPath $reportPath -ReportFullPath $reportFullPath -ConfiguredDataSources $mergedDataSources
            }
            else {
                Write-Log "  No data sources configured" -Level Info
            }
        }
        Write-Log "" -Level Info
        
        # Step 5: Set permissions (only in config file mode)
        Write-Log "Step 5: Setting permissions..." -Level Cyan
        if ($useEnvironmentVariables) {
            Write-Log "  Permissions not supported in environment variable mode" -Level Info
        }
        else {
            $mergedPermissions = Merge-Configuration -GlobalConfig $config.Permissions -SpecificConfig $mapping.Permissions -ReportName $reportFileName -ReportRelativePath $reportRelativePath -ConfigType "Permissions"
            
            if ($mergedPermissions.Count -gt 0) {
                Set-ReportPermissions -ServerUrl $serverUrl -ReportPath $reportPath -ReportFullPath $reportFullPath -ConfiguredPermissions $mergedPermissions
            }
            else {
                Write-Log "  No permissions configured" -Level Info
            }
        }
        
        $successCount++
        Write-Log "" -Level Info
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" -Level Error
        $failCount++
        Write-Log "" -Level Info
    }
}

Write-Log "========================================" -Level Cyan
Write-Log "Deployment Summary" -Level Cyan
Write-Log "========================================" -Level Cyan
Write-Log "" -Level Info
Write-Log "Total Reports: $($reports.Count)" -Level Info
Write-Log "Successful: $successCount" -Level Success
Write-Log "Failed: $failCount" -Level $(if ($failCount -gt 0) { "Error" }else { "Info" })
Write-Log "" -Level Info

if ($successCount -gt 0) {
    Write-Log "✓ Deployment complete!" -Level Success
}

#endregion
