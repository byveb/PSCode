<#
.SYNOPSIS
    Simple optimized PowerBI Cloud deployment script
.PARAMETER ConfigPath
    Path to JSON configuration file
.DESCRIPTION
    Features:
    - Auto-creates workspace if it doesn't exist
    - Reports array in DataSources/Permissions/Parameters/RefreshSchedule: applies config only to specified reports
    - No Reports array: applies config to all reports
    - ReportMappings with priority: ReportTargetPath > ReportName > file base name
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

# Import PowerBI module
Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop

# Load configuration
$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

# Process each report configuration
foreach ($report in $config.Reports) {
    if ($report.Type -ne 'Cloud') { continue }

    Write-Host "`n=== Deploying to Cloud: $($report.WorkspaceName) ===" -ForegroundColor Cyan

    # Connect to Power BI Service
    if ($report.ServicePrincipal) {
        $securePassword = ConvertTo-SecureString $report.ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($report.ClientId, $securePassword)
        Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId $report.TenantId
    } else {
        Connect-PowerBIServiceAccount
    }

    # Get or create workspace
    $workspace = Get-PowerBIWorkspace -Name $report.WorkspaceName -ErrorAction SilentlyContinue

    if (-not $workspace) {
        Write-Host "  Creating workspace: $($report.WorkspaceName)" -ForegroundColor Yellow
        $workspace = New-PowerBIWorkspace -Name $report.WorkspaceName
    }

    $workspaceId = $workspace.Id

    # Get all .pbix files from FolderPath
    $pbixFiles = Get-ChildItem -Path $report.FolderPath -Filter "*.pbix" -Recurse

    foreach ($file in $pbixFiles) {
        $fileName = $file.Name

        # Find matching report mapping
        $mapping = $report.ReportMappings | Where-Object { $fileName -match $_.SourcePattern }

        # Determine report name with priority: ReportTargetPath > ReportName > FileName
        if ($mapping -and $mapping.ReportTargetPath) {
            $reportName = $mapping.ReportTargetPath
        } elseif ($mapping -and $mapping.ReportName) {
            $reportName = $mapping.ReportName
        } else {
            $reportName = $file.BaseName
        }

        Write-Host "  Deploying: $fileName -> $reportName" -ForegroundColor Yellow

        # Check if report exists
        $existingReport = Get-PowerBIReport -WorkspaceId $workspaceId | Where-Object { $_.Name -eq $reportName }

        if ($existingReport -and -not $report.Overwrite) {
            Write-Host "    Report exists, skipping (Overwrite=false)" -ForegroundColor Gray
            continue
        }

        # Upload report
        $conflictAction = if ($report.Overwrite) { 'CreateOrOverwrite' } else { 'Abort' }

        $uploadParams = @{
            Path = $file.FullName
            WorkspaceId = $workspaceId
            ConflictAction = $conflictAction
        }

        if ($reportName -ne $file.BaseName) {
            $uploadParams.Name = $reportName
        }

        $uploadedReport = New-PowerBIReport @uploadParams
        Write-Host "    Uploaded successfully" -ForegroundColor Green

        # Get dataset ID
        $datasetId = $uploadedReport.DatasetId

        # Wait for dataset to be ready
        Start-Sleep -Seconds 3

        # Configure DataSources - filter by Reports array
        if ($report.DataSources -and $report.DataSources.Count -gt 0) {
            # Get applicable datasources (with Reports array matching or no Reports array)
            $applicableDataSources = $report.DataSources | Where-Object {
                Test-ConfigApplies -Config $_ -FileName $fileName
            }

            foreach ($ds in $applicableDataSources) {
                Write-Host "    Configuring datasource: $($ds.Name)" -ForegroundColor Gray

                # Get dataset datasources
                $datasetDatasources = Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/datasources" -Method Get | ConvertFrom-Json

                foreach ($datasetDs in $datasetDatasources.value) {
                    if ($datasetDs.datasourceType -eq $ds.Type) {
                        # Update connection details
                        if ($ds.Server -and $ds.Database) {
                            $updateBody = @{
                                connectionDetails = @{
                                    server = $ds.Server
                                    database = $ds.Database
                                }
                            } | ConvertTo-Json -Depth 5

                            Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/Default.UpdateDatasources" -Method Post -Body $updateBody
                        }

                        # Bind to gateway if specified
                        if ($ds.GatewayId -and $ds.DatasourceId) {
                            $gatewayBody = @{
                                gatewayObjectId = $ds.GatewayId
                            } | ConvertTo-Json

                            Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/Default.BindToGateway" -Method Post -Body $gatewayBody
                        }

                        # Update credentials
                        if ($ds.Username -and $ds.Password) {
                            $credBody = @{
                                credentialDetails = @{
                                    credentialType = "Basic"
                                    credentials = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($ds.Username):$($ds.Password)"))
                                    encryptedConnection = "Encrypted"
                                    encryptionAlgorithm = "None"
                                    privacyLevel = "Organizational"
                                }
                            } | ConvertTo-Json -Depth 5

                            Invoke-PowerBIRestMethod -Url "gateways/$($ds.GatewayId)/datasources/$($ds.DatasourceId)" -Method Patch -Body $credBody
                        }
                    }
                }
            }
        }

        # Configure Parameters - filter by Reports array
        if ($report.Parameters) {
            # Get applicable parameters (with Reports array matching or no Reports array)
            $applicableParams = $report.Parameters | Where-Object {
                Test-ConfigApplies -Config $_ -FileName $fileName
            }

            if ($applicableParams.Count -gt 0) {
                $updateParams = @{
                    updateDetails = @()
                }

                foreach ($param in $applicableParams) {
                    $updateParams.updateDetails += @{
                        name = $param.Name
                        newValue = $param.Value
                    }
                }

                Write-Host "    Updating parameters ($($applicableParams.Count))" -ForegroundColor Gray
                $paramBody = $updateParams | ConvertTo-Json -Depth 5
                Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/Default.UpdateParameters" -Method Post -Body $paramBody
            }
        }

        # Configure Refresh Schedule - filter by Reports array
        if ($report.RefreshSchedule) {
            # Get applicable refresh schedule (with Reports array matching or no Reports array)
            $applicableSchedule = $null

            if ($report.RefreshSchedule -is [Array]) {
                # Multiple refresh schedules - find the one that applies
                $applicableSchedule = $report.RefreshSchedule | Where-Object {
                    Test-ConfigApplies -Config $_ -FileName $fileName
                } | Select-Object -First 1
            } else {
                # Single refresh schedule object
                if (Test-ConfigApplies -Config $report.RefreshSchedule -FileName $fileName) {
                    $applicableSchedule = $report.RefreshSchedule
                }
            }

            if ($applicableSchedule) {
                Write-Host "    Configuring refresh schedule" -ForegroundColor Gray

                $scheduleBody = @{
                    value = @{
                        enabled = $applicableSchedule.Enabled
                        days = $applicableSchedule.Days
                        times = $applicableSchedule.Times
                        localTimeZoneId = $applicableSchedule.TimeZone
                        notifyOption = if ($applicableSchedule.NotifyOnFailure) { "MailOnFailure" } else { "NoNotification" }
                    }
                } | ConvertTo-Json -Depth 5

                Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/refreshSchedule" -Method Patch -Body $scheduleBody

                # Trigger Refresh
                if ($applicableSchedule.TriggerInitialRefresh) {
                    Write-Host "    Triggering dataset refresh" -ForegroundColor Gray
                    Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/refreshes" -Method Post

                    if ($applicableSchedule.WaitForRefresh) {
                        $timeout = if ($applicableSchedule.RefreshTimeout) { $applicableSchedule.RefreshTimeout } else { 30 }
                        $elapsed = 0

                        while ($elapsed -lt $timeout) {
                            Start-Sleep -Seconds 10
                            $elapsed += 10

                            $refreshes = Invoke-PowerBIRestMethod -Url "groups/$workspaceId/datasets/$datasetId/refreshes?`$top=1" -Method Get | ConvertFrom-Json
                            $lastRefresh = $refreshes.value[0]

                            if ($lastRefresh.status -eq 'Completed') {
                                Write-Host "    Refresh completed successfully" -ForegroundColor Green
                                break
                            } elseif ($lastRefresh.status -eq 'Failed') {
                                Write-Host "    Refresh failed: $($lastRefresh.serviceExceptionJson)" -ForegroundColor Red
                                break
                            }

                            Write-Host "    Refresh in progress... ($elapsed/$timeout minutes)" -ForegroundColor Gray
                        }
                    }
                }
            }
        }

        # Set Permissions (Workspace level) - filter by Reports array
        if ($report.Permissions) {
            # Get applicable permissions (with Reports array matching or no Reports array)
            $applicablePerms = $report.Permissions | Where-Object {
                Test-ConfigApplies -Config $_ -FileName $fileName
            }

            foreach ($perm in $applicablePerms) {
                Write-Host "    Setting permission: $($perm.Principal) = $($perm.AccessRight)" -ForegroundColor Gray

                $permBody = @{
                    identifier = $perm.Principal
                    principalType = if ($perm.Principal -like "*@*") { "User" } else { "Group" }
                    groupUserAccessRight = $perm.AccessRight
                } | ConvertTo-Json

                try {
                    Invoke-PowerBIRestMethod -Url "groups/$workspaceId/users" -Method Post -Body $permBody
                } catch {
                    Write-Host "    User may already have access" -ForegroundColor DarkGray
                }
            }
        }
    }
}

Write-Host "`n=== Deployment Complete ===" -ForegroundColor Green
