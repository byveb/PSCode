# Requires: $ProgramId, $PipelineId, $ExecutionId, headers (token already obtained)
# Data structures to track log progress
$logCache = @{}  # Key: "phaseId-stepId" -> number of lines already printed

function Get-StepLogs {
  param(
    [int]$ProgramId,
    [int]$PipelineId,
    [string]$ExecutionId,
    [int]$PhaseId,
    [int]$StepId
  )
  Ensure-Token
  $headers = @{
    Authorization     = "Bearer $token"
    "x-api-key"       = $ClientId
    "x-gw-ims-org-id" = $OrgId
    Accept            = "text/plain"
  }
  $uri = "https://cloudmanager.adobe.io/api/program/$ProgramId/pipelines/$PipelineId/executions/$ExecutionId/phases/$PhaseId/steps/$StepId/logs"
  try {
    $logText = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
    return $logText
  } catch {
    Write-Warn "Failed to fetch log for phase=$PhaseId step=$StepId: $($_.Exception.Message)"
    return $null
  }
}

function Stream-Running-StepLogs {
  param(
    $Steps
  )
  foreach($step in $Steps){
    if($step.status -eq "RUNNING"){
      # Some responses provide phaseId/stepId; if missing, adapt retrieval logic.
      $phaseId = $step.phaseId
      $stepId  = $step.id
      if(-not $phaseId -or -not $stepId){
        continue
      }
      $key = "$phaseId-$stepId"
      $logText = Get-StepLogs -ProgramId $ProgramId -PipelineId $PipelineId -ExecutionId $ExecutionId -PhaseId $phaseId -StepId $stepId
      if($null -eq $logText){ continue }
      $lines = $logText -split "`r?`n"
      $already = if($logCache.ContainsKey($key)) { $logCache[$key] } else { 0 }
      if($lines.Count -gt $already){
        $newLines = $lines[$already..($lines.Count-1)]
        foreach($l in $newLines){
          Write-Host ("[LOG:{0}] {1}" -f $step.action, $l)
        }
        $logCache[$key] = $lines.Count
      }
    }
  }
}

# Example integration into main polling loop:
# Inside while loop after fetching $st = Get-ExecutionStatus ...
# Stream-Running-StepLogs -Steps $st.Steps