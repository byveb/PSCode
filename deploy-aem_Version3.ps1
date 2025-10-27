<#
Adds dependency pre-check: ensure prerequisite environments (pipelines) have a recent successful execution with matching branch before triggering target pipeline.
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
  [int]$LogPollSeconds = 20,
  [int]$TimeoutMinutes = 90,

  [string]$SummaryPath = "aem-deploy-summary.json",
  [switch]$StreamLogs,
  [switch]$SaveLogs,
  [switch]$NoColor,
  [switch]$VerboseApi,

  # Dependency / prerequisite enforcement
  [string]$DependencyMapPath = "./dependency-aem.json",
  [switch]$RequirePreviousSuccess,
  [int]$MaxPrereqAgeHours = 24
)

function W-Info($m){ if($NoColor){Write-Host $m}else{Write-Host $m -ForegroundColor Cyan} }
function W-Ok($m){ if($NoColor){Write-Host $m}else{Write-Host $m -ForegroundColor Green} }
function W-Warn($m){ if($NoColor){Write-Host $m}else{Write-Host $m -ForegroundColor Yellow} }
function W-Err($m){ if($NoColor){Write-Host $m}else{Write-Host $m -ForegroundColor Red} }

if(-not $PipelineId -and -not $EnvironmentName){
  W-Err "Provide -PipelineId or -EnvironmentName."
  exit 4
}

# Load environment -> pipeline map
if($EnvironmentName -and -not $PipelineId){
  if(-not (Test-Path $EnvironmentMapPath)){ W-Err "Environment map not found: $EnvironmentMapPath"; exit 4 }
  try { $envMap = Get-Content $EnvironmentMapPath -Raw | ConvertFrom-Json } catch { W-Err "Parse error in $EnvironmentMapPath: $($_.Exception.Message)"; exit 4 }
  if(-not $envMap.PSObject.Properties.Name.Contains($EnvironmentName)){ W-Err "Environment '$EnvironmentName' missing in map."; exit 4 }
  $PipelineId = [int]$envMap.$EnvironmentName
  W-Info "Resolved environment '$EnvironmentName' -> pipeline $PipelineId"
}

# Load dependency map if needed
$dependencies = $null
if($RequirePreviousSuccess){
  if(-not (Test-Path $DependencyMapPath)){
    W-Err "Dependency map file not found: $DependencyMapPath"
    exit 4
  }
  try {
    $dependencies = Get-Content $DependencyMapPath -Raw | ConvertFrom-Json
  } catch {
    W-Err "Failed parsing dependency map: $($_.Exception.Message)"
    exit 4
  }
  W-Info "Loaded dependency map from $DependencyMapPath"
}

$baseCM = "https://cloudmanager.adobe.io/api"
$token = $null
$tokenObtainedAt = $null
$tokenRefreshAfterSeconds = 60 * 50

# --- Auth + API wrappers (same as previous script; omitted for brevity) ---
function Get-AccessToken { param($ClientId,$ClientSecret,$Scopes,$RegionHost)
  $body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = $Scopes
  }
  try {
    $r = Invoke-RestMethod -Method Post -Uri "https://$RegionHost/ims/token/v3" -Body $body -ErrorAction Stop
  } catch {
    W-Err "Token request failed: $($_.Exception.Message)"; exit 5
  }
  if(-not $r.access_token){ W-Err "No access_token."; exit 5 }
  $r.access_token
}

function Ensure-Token {
  if(-not $token){
    W-Info "Obtaining token..."
    $script:token = Get-AccessToken $ClientId $ClientSecret $Scopes $RegionHost
    $script:tokenObtainedAt = Get-Date
  } else {
    $elapsed = (Get-Date) - $tokenObtainedAt
    if($elapsed.TotalSeconds -ge $tokenRefreshAfterSeconds){
      W-Info "Refreshing token..."
      $script:token = Get-AccessToken $ClientId $ClientSecret $Scopes $RegionHost
      $script:tokenObtainedAt = Get-Date
    }
  }
}

function Invoke-CM {
  param([string]$Method,[string]$Path,[object]$Body,[string]$AcceptOverride)
  Ensure-Token
  $headers = @{
    Authorization     = "Bearer $token"
    "x-api-key"       = $ClientId
    "x-gw-ims-org-id" = $OrgId
    Accept            = ($AcceptOverride ? $AcceptOverride : "application/json")
  }
  $uri = "$baseCM/$Path"
  $p = @{ Method=$Method; Uri=$uri; Headers=$headers; ErrorAction='Stop' }
  if($Body){
    $p.ContentType = "application/json"
    $p.Body = ($Body | ConvertTo-Json -Depth 10)
  }
  try {
    $resp = Invoke-RestMethod @p
    if($VerboseApi){ W-Warn "API $Method $Path => $(($resp | ConvertTo-Json -Depth 10))" }
    $resp
  } catch {
    W-Err "API failure ($Method $Path): $($_.Exception.Message)"
    throw
  }
}

# --- Helper: list executions to find last successful for a pipeline ---
function Get-LastSuccessfulExecution {
  param([int]$ProgramId,[int]$PipelineId)
  $resp = Invoke-CM -Method Get -Path "program/$ProgramId/pipelines/$PipelineId/executions"
  $execs = $resp._embedded.executions
  if(-not $execs){ return $null }
  # Filter by terminal success status FINISHED
  $successes = $execs | Where-Object { $_.status -eq "FINISHED" }
  # Sort by finishedAt or startedAt desc
  $ordered = $successes | Sort-Object -Property finishedAt -Descending
  $ordered | Select-Object -First 1
}

# --- Precondition Check ---
function Check-Prerequisites {
  param([string]$TargetEnv,[object]$Dependencies,[object]$EnvMap,[string]$Branch,[int]$ProgramId,[int]$MaxAgeHours)

  if(-not $Dependencies.PSObject.Properties.Name.Contains($TargetEnv)){
    W-Info "No prerequisites defined for environment '$TargetEnv'."
    return
  }
  $prereqs = $Dependencies.$TargetEnv
  W-Info "Prerequisites for '$TargetEnv': $($prereqs -join ', ')"

  foreach($pre in $prereqs){
    if(-not $EnvMap.PSObject.Properties.Name.Contains($pre)){
      W-Err "Prerequisite environment '$pre' missing in env map."
      exit 4
    }
    $prePipelineId = [int]$EnvMap.$pre
    W-Info "Checking last successful execution of '$pre' (pipeline $prePipelineId)..."
    $last = Get-LastSuccessfulExecution -ProgramId $ProgramId -PipelineId $prePipelineId
    if(-not $last){
      W-Err "No successful execution found for prerequisite '$pre'. Aborting."
      exit 3
    }
    # Validate age
    $finishedAt = if($last.finishedAt){ [DateTime]$last.finishedAt } elseif($last.startedAt){ [DateTime]$last.startedAt } else { Get-Date }
    $ageHours = ((Get-Date) - $finishedAt).TotalHours
    if($ageHours -gt $MaxAgeHours){
      W-Err ("Prerequisite '$pre' success is too old ({0:N1}h > {1}h). Aborting." -f $ageHours,$MaxAgeHours)
      exit 3
    }
    # Validate branch (if the execution object exposes it). Some versions nest branch under pipeline or execution.
    $execBranch = $last.branch
    if($Branch){
      if($execBranch -and ($execBranch -ne $Branch)){
        W-Err "Prerequisite '$pre' last success branch '$execBranch' does not match target branch '$Branch'. Aborting."
        exit 3
      } elseif(-not $execBranch){
        W-Warn "Prerequisite '$pre' execution did not report branch; cannot verify branch match."
      }
    }
    W-Ok "Prerequisite '$pre' OK (ExecId=$($last.id), FinishedAt=$finishedAt, Branch=$execBranch)"
  }
  W-Ok "All prerequisites satisfied."
}

# --- Start pipeline ---
function Start-Pipeline {
  param([int]$ProgramId,[int]$PipelineId,[string]$Branch)
  $body = $null
  if($Branch){ $body = @{ branch = $Branch }; W-Info "Branch override: $Branch" }
  $resp = Invoke-CM -Method Post -Path "program/$ProgramId/pipelines/$PipelineId/execution" -Body $body
  $execId = $resp.id
  if(-not $execId -and $resp.execution){ $execId = $resp.execution.id }
  if(-not $execId){ throw "Could not extract execution ID" }
  W-Ok "Execution started (ID=$execId)."
  $execId
}

# --- Execution polling & log streaming (same as previous enhanced script, omitted for brevity) ---
# For brevity here, reuse your existing Monitor-Execution, Stream-StepLogs, etc.
# (You can copy those unchanged from the previous version.)

# To keep this example concise, imagine the existing functions are present below:
function Monitor-Execution { param([int]$ProgramId,[int]$PipelineId,[string]$ExecutionId,[int]$PollSeconds,[int]$LogPollSeconds,[int]$TimeoutMinutes,[switch]$StreamLogs)
  # (Implementation identical to previously provided enhanced script.)
  # Placeholder minimal logic:
  $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
  while($true){
    $resp = Invoke-CM -Method Get -Path "program/$ProgramId/pipelines/$PipelineId/executions/$ExecutionId"
    $status = $resp.status
    if(-not $status -and $resp.execution){ $status = $resp.execution.status }
    Write-Host ("Status: {0}" -f $status)
    if($status -in "FINISHED"){ return @{Success=$true; Raw=$resp} }
    if($status -in @("FAILED","ERROR","CANCELLED")){ return @{Success=$false; Raw=$resp} }
    if((Get-Date) -gt $deadline){ return @{Success=$false; TimedOut=$true; Raw=$resp} }
    Start-Sleep -Seconds $PollSeconds
  }
}

# --- Precondition enforcement ---
if($RequirePreviousSuccess){
  if(-not $EnvironmentName){
    W-Err "-RequirePreviousSuccess needs -EnvironmentName (to look up dependency chain)."
    exit 4
  }
  Check-Prerequisites -TargetEnv $EnvironmentName -Dependencies $dependencies -EnvMap $envMap -Branch $Branch -ProgramId $ProgramId -MaxAgeHours $MaxPrereqAgeHours
}

# --- Start + monitor ---
try {
  $executionId = Start-Pipeline -ProgramId $ProgramId -PipelineId $PipelineId -Branch $Branch
} catch { W-Err "Failed to start pipeline: $($_.Exception.Message)"; exit 3 }

$result = Monitor-Execution -ProgramId $ProgramId -PipelineId $PipelineId -ExecutionId $executionId -PollSeconds $PollSeconds -LogPollSeconds $LogPollSeconds -TimeoutMinutes $TimeoutMinutes -StreamLogs:$StreamLogs

# --- Summary (simplified) ---
if($SummaryPath){
  $summary = [PSCustomObject]@{
    ProgramId    = $ProgramId
    PipelineId   = $PipelineId
    Environment  = $EnvironmentName
    Branch       = $Branch
    ExecutionId  = $executionId
    Success      = $result.Success
    TimedOut     = $result.TimedOut
    GeneratedAt  = (Get-Date)
    RawStatus    = $result.Raw.status
  }
  $summary | ConvertTo-Json -Depth 10 | Out-File -FilePath $SummaryPath -Encoding UTF8
  W-Info "Summary: $SummaryPath"
}

if($result.Success){ exit 0 }
elseif($result.TimedOut){ exit 2 }
else { exit 3 }