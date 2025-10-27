<#
.SYNOPSIS
  Non-destructive one-way sync from local repository to AEM Cloud Git (no dry run option).

.DESCRIPTION
  - Pushes local branches and tags to AEM Cloud without deleting or pruning any remote refs.
  - Supports exact exclusions (-ExcludeBranches) and pattern exclusions (-ExcludePatterns) using simple wildcards (*, ?).
  - Detects divergent branches; if -AllowForce is specified, force-pushes them, otherwise skips them.
  - Batches normal branch pushes (falls back to per-branch if batch fails).
  - Retries pushes on transient failures with exponential backoff.
  - Uses a temporary credential helper instead of embedding credentials in remote URL.
  - Optionally writes a JSON summary of the sync outcome (-SummaryPath).
  - Optional -SkipUntrackedCheck to include branches that do not yet have an upstream tracking ref (origin/<branch>).

.PARAMETER SourceRemote
  Source remote name (default: origin)

.PARAMETER TargetRemote
  Target remote name (default: aemcloud)

.PARAMETER ExcludeBranches
  Exact branch names to exclude.

.PARAMETER ExcludePatterns
  Wildcard exclusion patterns (e.g. feature/*, temp?, wip*)

.PARAMETER AllowForce
  If set, force-pushes divergent branches.

.PARAMETER AemCloudGitUrl
  HTTPS remote URL for AEM Cloud Git repository.

.PARAMETER AemCloudUsername
  Username / PAT user for authentication.

.PARAMETER AemCloudPassword
  Password / PAT token.

.PARAMETER SummaryPath
  Path to write JSON summary (optional).

.PARAMETER MaxRetries
  Maximum retry attempts per push (default: 3).

.PARAMETER RetryDelaySeconds
  Base delay used for exponential backoff (default: 3). Actual delay = (2^(attempt-1)) * base.

.PARAMETER SkipUntrackedCheck
  Include branches that do not have a corresponding origin/<branch> tracking ref.

.EXAMPLE
  pwsh -File scripts/Sync-AEMCloud-Enhanced-NoDryRun.ps1 `
    -AemCloudGitUrl $env:AEM_CLOUD_GIT_URL `
    -AemCloudUsername $env:AEM_CLOUD_GIT_USERNAME `
    -AemCloudPassword $env:AEM_CLOUD_GIT_PASSWORD `
    -ExcludePatterns @('feature/*','sandbox*') -AllowForce -SummaryPath sync-summary.json

.NOTES
  Ensure the clone is not shallow (disable depth) for accurate merge-base divergence detection.
#>

[CmdletBinding()]
param(
  [string]$SourceRemote = 'origin',
  [string]$TargetRemote = 'aemcloud',
  [string[]]$ExcludeBranches = @(),
  [string[]]$ExcludePatterns = @(),
  [switch]$AllowForce,
  [string]$AemCloudGitUrl,
  [string]$AemCloudUsername,
  [string]$AemCloudPassword,
  [string]$SummaryPath,
  [int]$MaxRetries = 3,
  [int]$RetryDelaySeconds = 3,
  [switch]$SkipUntrackedCheck
)

function Write-Log {
  param([string]$Level='INFO',[string]$Message)
  $ts = (Get-Date).ToString('u')
  Write-Host "[$ts][$Level] $Message"
}

# --- Preconditions ---
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Log ERROR "git CLI not found in PATH."
  exit 20
}
if (-not $AemCloudGitUrl) {
  Write-Log ERROR "AemCloudGitUrl is required."
  exit 21
}
if (-not $AemCloudUsername -or -not $AemCloudPassword) {
  Write-Log ERROR "AemCloudUsername and AemCloudPassword are required."
  exit 22
}

Write-Log INFO "Starting non-destructive sync (no dry run). Force=$($AllowForce.IsPresent)"

# --- Fetch source state ---
git fetch --prune $SourceRemote | Out-Null

# --- Credential helper (avoids storing credentials in remote URL/config permanently) ---
git config --local credential.helper "!f() { echo username=$AemCloudUsername; echo password=$AemCloudPassword; }; f"

# --- Configure target remote ---
if (git remote | Where-Object { $_ -eq $TargetRemote }) {
  git remote set-url $TargetRemote $AemCloudGitUrl
  Write-Log INFO "Updated remote '$TargetRemote'."
} else {
  git remote add $TargetRemote $AemCloudGitUrl
  Write-Log INFO "Added remote '$TargetRemote'."
}

# --- Fetch target (needed for divergence detection) ---
try {
  git fetch $TargetRemote | Out-Null
} catch {
  Write-Log WARN "Fetch from target failed (might be first push): $($_.Exception.Message)"
}

# --- Enumerate local branches ---
$localBranches = git for-each-ref --format='%(refname:short)' refs/heads | Where-Object { $_ }
Write-Log INFO "Discovered $($localBranches.Count) local branches."

function Matches-Pattern {
  param([string]$Name,[string[]]$Patterns)
  foreach ($p in $Patterns) {
    $regex = '^' + ($p -replace '\*','.*' -replace '\?','.') + '$'
    if ($Name -match $regex) { return $true }
  }
  return $false
}

function Is-Excluded {
  param([string]$Branch)
  if ($ExcludeBranches -contains $Branch) { return $true }
  if (Matches-Pattern -Name $Branch -Patterns $ExcludePatterns) { return $true }
  return $false
}

# --- Summary object ---
$summary = [ordered]@{
  pushed       = @()
  forced       = @()
  skipped      = @()
  divergent    = @()
  errors       = @()
  timestampUtc = (Get-Date).ToUniversalTime().ToString('o')
  forceEnabled = $AllowForce.IsPresent
}

$branchesToNormalPush = @()
$branchesToForcePush  = @()

# --- Classification ---
foreach ($br in $localBranches) {
  if (Is-Excluded $br) {
    Write-Log INFO "Excluded branch '$br'."
    $summary.skipped += $br
    continue
  }

  if (-not $SkipUntrackedCheck) {
    git show-ref --verify --quiet "refs/remotes/$SourceRemote/$br"
    if (-not $LASTEXITCODE -eq 0) {
      Write-Log WARN "Skipping '$br' (no upstream $SourceRemote/$br). Use -SkipUntrackedCheck to include."
      $summary.skipped += $br
      continue
    }
  }

  $divergent = $false
  git show-ref --verify --quiet "refs/remotes/$TargetRemote/$br"
  if ($LASTEXITCODE -eq 0) {
    $localHash  = (git rev-parse "$SourceRemote/$br").Trim()
    $remoteHash = (git rev-parse "$TargetRemote/$br").Trim()
    $mergeBase  = (git merge-base "$SourceRemote/$br" "$TargetRemote/$br").Trim()
    if ($localHash -ne $remoteHash -and $mergeBase -ne $remoteHash) {
      $divergent = $true
      $summary.divergent += $br
      Write-Log INFO "Branch '$br' is divergent."
    }
  }

  if ($divergent) {
    if ($AllowForce) {
      $branchesToForcePush += $br
    } else {
      Write-Log ERROR "Divergent branch '$br' skipped (force not enabled)."
      $summary.skipped += $br
    }
  } else {
    $branchesToNormalPush += $br
  }
}

# --- Retry helper ---
function Retry-Push {
  param(
    [scriptblock]$Action,
    [string]$RefName
  )
  for ($attempt=1; $attempt -le $MaxRetries; $attempt++) {
    & $Action
    if ($LASTEXITCODE -eq 0) { return $true }
    $delay = [int]([math]::Pow(2, ($attempt-1))) * $RetryDelaySeconds
    Write-Log WARN "Push failed for '$RefName' attempt $attempt/$MaxRetries; retrying in $delay s."
    Start-Sleep -Seconds $delay
  }
  return $false
}

# --- Push normal branches (batch) ---
if ($branchesToNormalPush.Count -gt 0) {
  Write-Log INFO "Normal branches to push: $([string]::Join(', ', $branchesToNormalPush))"
  $refspecs = $branchesToNormalPush | ForEach-Object { "$_:$_" }
  $batchAction = { git push $TargetRemote $refspecs }
  if (Retry-Push -Action $batchAction -RefName 'batch-normal') {
    $summary.pushed += $branchesToNormalPush
  } else {
    Write-Log ERROR "Batch push failed; retrying per branch."
    foreach ($b in $branchesToNormalPush) {
      $singleAction = { git push $TargetRemote "$b:$b" }
      if (Retry-Push -Action $singleAction -RefName $b) {
        $summary.pushed += $b
      } else {
        Write-Log ERROR "Failed to push branch '$b'."
        $summary.errors += "$b (push failed)"
      }
    }
  }
} else {
  Write-Log INFO "No normal branches to push."
}

# --- Force pushes ---
foreach ($b in $branchesToForcePush) {
  Write-Log INFO "Force pushing divergent branch '$b'."
  $forceAction = { git push $TargetRemote "+$b:$b" }
  if (Retry-Push -Action $forceAction -RefName $b) {
    $summary.forced += $b
  } else {
    Write-Log ERROR "Force push failed for '$b'."
    $summary.errors += "$b (force push failed)"
  }
}

# --- Tags ---
Write-Log INFO "Pushing tags (non-destructive)."
$tagAction = { git push $TargetRemote --tags }
if (-not (Retry-Push -Action $tagAction -RefName 'tags')) {
  Write-Log WARN "Tag push failed."
  $summary.errors += "tags (push failed)"
}

# --- Summary JSON (optional) ---
if ($SummaryPath) {
  try {
    ($summary | ConvertTo-Json -Depth 5) | Out-File -Encoding UTF8 $SummaryPath
    Write-Log INFO "Summary written to $SummaryPath."
  } catch {
    Write-Log WARN "Failed to write summary JSON: $($_.Exception.Message)"
  }
}

# --- Cleanup credential helper ---
git config --local --unset credential.helper | Out-Null

# --- Final status ---
if ($summary.errors.Count -gt 0) {
  Write-Log ERROR "Sync completed with errors."
  exit 1
}

Write-Log INFO "Sync completed successfully. Pushed=$($summary.pushed.Count) Forced=$($summary.forced.Count) Skipped=$($summary.skipped.Count)"
exit 0