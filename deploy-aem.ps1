<#
.SYNOPSIS
  Trigger and monitor an Adobe AEM Cloud Manager pipeline execution.

.DESCRIPTION
  Uses Adobe Cloud Manager API to start a pipeline (with optional branch override) and polls until completion or timeout.
  Supports mapping an EnvironmentName to a PipelineId via a JSON dictionary.
  Outputs detailed status and a machine-readable summary file.

.PARAMETER ProgramId
  Cloud Manager Program ID (integer).

.PARAMETER PipelineId
  Cloud Manager Pipeline ID. Optional if EnvironmentName is supplied.

.PARAMETER EnvironmentName
  Friendly name used to look up PipelineId from -EnvironmentMapPath JSON.

.PARAMETER EnvironmentMapPath
  Path to JSON file mapping environment names to pipeline IDs.

.PARAMETER Branch
  Git branch to deploy (override). Leave blank to use pipeline default.

.PARAMETER ClientId
  Adobe Developer Console client ID.

.PARAMETER ClientSecret
  Adobe Developer Console client secret.

.PARAMETER OrgId
  IMS Organization ID (x-gw-ims-org-id).

.PARAMETER Scopes
  Space-separated scopes; default "cloudmanager".

.PARAMETER RegionHost
  IMS auth host base (default ims-na1.adobelogin.com). Change if your org uses a different region.

.PARAMETER PollSeconds
  Interval between status checks (default 45).

.PARAMETER TimeoutMinutes
  Maximum time to wait for pipeline completion (default 90).

.PARAMETER SummaryPath
  Where to write a JSON summary (default aem-deploy-summary.json). Set to $null to skip.

.PARAMETER NoColor
  Disable colored output if set.

.PARAMETER VerboseApi
  Show raw API responses for debugging.

.EXAMPLE
  ./deploy-aem.ps1 -ProgramId 123 -PipelineId 456 -Branch feature/my-change `
    -ClientId $env:AEM_CLIENT_ID -ClientSecret $env:AEM_CLIENT_SECRET -OrgId $env:AEM_ORG_ID

.EXAMPLE
  ./deploy-aem.ps1 -ProgramId 123 -EnvironmentName dev -EnvironmentMapPath ./config-aem-environments.json `
    -ClientId (Get-Item env:AEM_CLIENT_ID).Value -ClientSecret (Get-Item env:AEM_CLIENT_SECRET).Value -OrgId (Get-Item env:AEM_ORG_ID).Value

.NOTES
  Exit codes:
    0 success
    2 timeout
    3 pipeline failure
    4 parameter error
    5 authentication error
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][int]$ProgramId,
  [int]$PipelineId,
  [string]$EnvironmentName,
  [string]$EnvironmentMapPath = "./config-aem-environments.json",
  [string]$Branch,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$OrgId,
  [string]$Scopes = "cloudmanager",
  [string]$RegionHost = "ims-na1.adobelogin.com",
  [int]$PollSeconds = 45,
  [int]$TimeoutMinutes = 90,
  [string]$SummaryPath = "aem-deploy-summary.json",
  [switch]$NoColor,
  [switch]$VerboseApi
)

# ------------------------ Helpers: Color ------------------------
function Write-Info($msg){ if(-not $NoColor){ Write-Host $msg -ForegroundColor Cyan } else { Write-Host $msg } }
function Write-Ok($msg){ if(-not $NoColor){ Write-Host $msg -ForegroundColor Green } else { Write-Host $msg } }
function Write-Warn($msg){ if(-not $NoColor){ Write-Host $msg -ForegroundColor Yellow } else { Write-Host $msg } }
function Write-Err($msg){ if(-not $NoColor){ Write-Host $msg -ForegroundColor Red } else { Write-Host $msg } }

# ------------------------ Validate Inputs ------------------------
if(-not $PipelineId -and -not $EnvironmentName){
  Write-Err "Either -PipelineId or -EnvironmentName must be provided."
  exit 4
}

if($EnvironmentName -and -not $PipelineId){
  if(-not (Test-Path $EnvironmentMapPath)){
    Write-Err "Environment map file not found at $EnvironmentMapPath"
    exit 4
  }
  try {
    $envMap = Get-Content $EnvironmentMapPath -Raw | ConvertFrom-Json
  } catch {
    Write-Err "Failed to parse environment map JSON: $($_.Exception.Message)"
    exit 4
  }
  if(-not $envMap.PSObject.Properties.Name.Contains($EnvironmentName)){
    Write-Err "Environment '$EnvironmentName' not found in map."
    exit 4
  }
  $PipelineId = [int]$envMap.$EnvironmentName
  Write-Info "Resolved EnvironmentName '$EnvironmentName' to PipelineId $PipelineId"
}

# ------------------------ Globals ------------------------
$baseCloudManager = "https://cloudmanager.adobe.io/api"
$token = $null
$tokenObtainedAt = $null
$tokenTTLSeconds = 60 * 50  # Refresh after ~50 mins proactively.

# ------------------------ Auth ------------------------
function Get-AccessToken {
  param([string]$ClientId,[string]$ClientSecret,[string]$Scopes,[string]$RegionHost)

  $body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = $Scopes
  }

  try {
    $resp = Invoke-RestMethod -Method Post -Uri "https://$RegionHost/ims/token/v3" -Body $body -ErrorAction Stop
  } catch {
    Write-Err "Token request failed: $($_.Exception.Message)"
    exit 5
  }

  if(-not $resp.access_token){
    Write-Err "No access_token in response."
    exit 5
  }
  return $resp.access_token
}

function Ensure-Token {
  if(-not $token){
    Write-Info "Obtaining access token..."
    $script:token = Get-AccessToken -ClientId $ClientId -ClientSecret $ClientSecret -Scopes $Scopes -RegionHost $RegionHost
    $script:tokenObtainedAt = Get-Date
  } else {
    $elapsed = (Get-Date) - $tokenObtainedAt
    if($elapsed.TotalSeconds -ge $tokenTTLSeconds){
      Write-Info "Refreshing access token (elapsed $([math]::Round($elapsed.TotalMinutes,2)) min)..."
      $script:token = Get-AccessToken -ClientId $ClientId -ClientSecret $ClientSecret -Scopes $Scopes -RegionHost $RegionHost
      $script:tokenObtainedAt = Get-Date
    }
  }
}

# ------------------------ API Call Wrapper ------------------------
function Invoke-AemApi {
  param(
    [string]$Method,
    [string]$Path,
    [object]$Body
  )
  Ensure-Token
  $headers = @{
    Authorization        = "Bearer $token"
    "x-api-key"          = $ClientId
    "x-gw-ims-org-id"    = $OrgId
    Accept               = "application/json"
  }
  $uri = "$baseCloudManager/$Path"

  $invokeParams = @{
    Method      = $Method
    Uri         = $uri
    Headers     = $headers
    ErrorAction = 'Stop'
  }
  if($Body){
    $invokeParams.Body = ($Body | ConvertTo-Json -Depth 10)
    $invokeParams.ContentType = "application/json"
  }

  try {
    $resp = Invoke-RestMethod @invokeParams
    if($VerboseApi){ Write-Warn "Raw API response ($Method $Path): $(($resp | ConvertTo-Json -Depth 10))" }
    return $resp
  } catch {
    Write-Err "API call failed ($Method $Path): $($_.Exception.Message)"
    throw
  }
}

# ------------------------ Trigger Execution ------------------------
function Start-Pipeline {
  param([int]$ProgramId,[int]$PipelineId,[string]$Branch)

  $body = $null
  if($Branch){
    $body = @{ branch = $Branch }
    Write-Info "Using branch override: $Branch"
  } else {
    Write-Info "No branch override provided; pipeline default will be used."
  }
  $resp = Invoke-AemApi -Method Post -Path "program/$ProgramId/pipelines/$PipelineId/execution" -Body $body

  # Execution ID can appear as .id or nested .execution.id
  $execId = $resp.id
  if(-not $execId -and $resp.execution){ $execId = $resp.execution.id }

  if(-not $execId){
    Write-Err "Unable to parse execution ID from response."
    throw "NoExecutionId"
  }
  Write-Ok "Pipeline execution started. ExecutionId: $execId"
  return $execId
}

# ------------------------ Fetch Status ------------------------
function Get-ExecutionStatus {
  param([int]$ProgramId,[int]$PipelineId,[string]$ExecutionId)

  $resp = Invoke-AemApi -Method Get -Path "program/$ProgramId/pipelines/$PipelineId/executions/$ExecutionId"
  # Status field may differ; accommodate .status or .execution.status
  $status = $resp.status
  if(-not $status -and $resp.execution){ $status = $resp.execution.status }
  $steps = $resp.steps
  if(-not $steps -and $resp.execution){ $steps = $resp.execution.steps }

  return [PSCustomObject]@{
    Raw     = $resp
    Status  = $status
    Steps   = $steps
  }
}

# ------------------------ Derive Failure Reason ------------------------
function Get-FailureDetails {
  param($Steps)
  if(-not $Steps){ return $null }
  $failed = @()
  foreach($s in $Steps){
    if($s.status -in @("ERROR","FAILED","CANCELLED")){
      $failed += [PSCustomObject]@{
        Action     = $s.action
        Status     = $s.status
        StartedAt  = $s.startedAt
        FinishedAt = $s.finishedAt
        Message    = ($s.message, $s.errorMessage, $s.details) -join " | "
      }
    }
  }
  if($failed.Count -eq 0){ return $null }
  return $failed
}

# ------------------------ Wait Loop ------------------------
function Wait-ForCompletion {
  param(
    [int]$ProgramId,
    [int]$PipelineId,
    [string]$ExecutionId,
    [int]$PollSeconds,
    [int]$TimeoutMinutes
  )

  $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
  while($true){
    $st = Get-ExecutionStatus -ProgramId $ProgramId -PipelineId $PipelineId -ExecutionId $ExecutionId
    $status = $st.Status
    $now = Get-Date
    $remaining = [math]::Round(($deadline - $now).TotalMinutes,2)

    Write-Info ("[{0}] Status: {1} (time left: {2} min)" -f (Get-Date -Format 'HH:mm:ss'), $status, $remaining)

    if($st.Steps){
      foreach($step in $st.Steps){
        $line = "  - Step {0,-25} | {1,-10}" -f $step.action, $step.status
        Write-Host $line
      }
    }

    switch ($status) {
      "FINISHED" {
        Write-Ok "Pipeline finished successfully."
        return @{ Success = $true; Raw = $st.Raw; Steps = $st.Steps }
      }
      {$_ -in "ERROR","FAILED","CANCELLED"} {
        Write-Err "Pipeline ended with status: $status"
        $fail = Get-FailureDetails -Steps $st.Steps
        if($fail){
          Write-Err "Failure details:"
          $fail | Format-Table -AutoSize
        }
        return @{ Success = $false; Raw = $st.Raw; Steps = $st.Steps; FailureDetails = $fail }
      }
    }

    if($now -ge $deadline){
      Write-Err "Timeout reached after $TimeoutMinutes minutes."
      return @{ Success = $false; TimedOut = $true; Raw = $st.Raw; Steps = $st.Steps }
    }

    Start-Sleep -Seconds $PollSeconds
  }
}

# ------------------------ Main Flow ------------------------
try {
  $executionId = Start-Pipeline -ProgramId $ProgramId -PipelineId $PipelineId -Branch $Branch
} catch {
  Write-Err "Failed starting pipeline: $($_.Exception.Message)"
  exit 3
}

$result = Wait-ForCompletion -ProgramId $ProgramId -PipelineId $PipelineId -ExecutionId $executionId -PollSeconds $PollSeconds -TimeoutMinutes $TimeoutMinutes

# ------------------------ Summary Output ------------------------
if($SummaryPath){
  $summary = [PSCustomObject]@{
    ProgramId      = $ProgramId
    PipelineId     = $PipelineId
    Environment    = $EnvironmentName
    Branch         = $Branch
    ExecutionId    = $executionId
    Success        = $result.Success
    TimedOut       = $result.TimedOut
    FinishedAt     = (Get-Date)
    FailureDetails = $result.FailureDetails
    Steps          = $result.Steps
  }
  try {
    $summary | ConvertTo-Json -Depth 10 | Out-File -FilePath $SummaryPath -Encoding UTF8
    Write-Info "Summary written to $SummaryPath"
  } catch {
    Write-Warn "Failed writing summary file: $($_.Exception.Message)"
  }
}

# ------------------------ Exit Codes ------------------------
if($result.Success){
  exit 0
} elseif($result.TimedOut){
  exit 2
} else {
  exit 3
}