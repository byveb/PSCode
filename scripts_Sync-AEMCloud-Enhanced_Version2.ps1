<#
.SYNOPSIS
  Enhanced non-destructive one-way sync from local repo to AEM Cloud Git.

.DESCRIPTION
  - Pushes branches & tags without deleting remote refs.
  - Supports exact exclusions (-ExcludeBranches) AND pattern exclusions (-ExcludePatterns) using simple wildcard (*, ?).
  - Divergence detection with optional force.
  - Multi-branch consolidated push (optional), falling back to per-branch if force needed for some.
  - Retry logic for pushes.
  - Optional JSON summary output.

.PARAMETER SourceRemote
  Source remote name (default origin).

.PARAMETER TargetRemote
  Target remote name (default aemcloud).

.PARAMETER ExcludeBranches
  Exact names to exclude.

.PARAMETER ExcludePatterns
  Wildcard patterns (e.g. feature/*, temp?, wip*).

.PARAMETER AllowForce
  Force push divergent branches.

.PARAMETER AemCloudGitUrl
  Remote HTTPS URL.

.PARAMETER AemCloudUsername
  Username / PAT user.

.PARAMETER AemCloudPassword
  Password / PAT.

.PARAMETER DryRun
  Preview only.

.PARAMETER SummaryPath
  Path to write JSON summary (optional).

.PARAMETER MaxRetries
  Max retries per push (default 3).

.PARAMETER RetryDelaySeconds
  Base delay for exponential backoff (default 3).

.PARAMETER SkipUntrackedCheck
  If set, pushes local branches even if no origin/<branch> ref exists.

.EXAMPLE
  pwsh -File scripts/Sync-AEMCloud-Enhanced.ps1 -AemCloudGitUrl $env:AEM_CLOUD_GIT_URL `
    -AemCloudUsername $env:AEM_CLOUD_GIT_USERNAME -AemCloudPassword $env:AEM_CLOUD_GIT_PASSWORD `
    -ExcludePatterns @('feature/*','sandbox*') -AllowForce -SummaryPath sync-summary.json
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
  [switch]$DryRun,
  [string]$SummaryPath,
  [int]$MaxRetries = 3,
  [int]$RetryDelaySeconds = 3,
  [switch]$SkipUntrackedCheck
)

function Write-Log { param([string]$Level='INFO',[string]$Message)
  $ts = (Get-Date).ToString('u'); Write-Host "[$ts][$Level] $Message"
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) { Write-Log ERROR "git not found."; exit 20 }
if (-not $AemCloudGitUrl) { Write-Log ERROR "AemCloudGitUrl required."; exit 21 }
if (-not $AemCloudUsername -or -not $AemCloudPassword) { Write-Log ERROR "Credentials required."; exit 22 }

Write-Log INFO "Enhanced sync starting. DryRun=$($DryRun.IsPresent) Force=$($AllowForce.IsPresent)"

# Ensure full history (shallow clones can break merge-base)
git fetch --prune $SourceRemote | Out-Null

# Temporary credential helper (avoid storing in remote URL)
git config --local credential.helper "!f() { echo username=$AemCloudUsername; echo password=$AemCloudPassword; }; f"
# Add/update remote without credentials embedded
if (git remote | Where-Object { $_ -eq $TargetRemote }) {
  git remote set-url $TargetRemote $AemCloudGitUrl
  Write-Log INFO "Updated remote '$TargetRemote'."
} else {
  git remote add $TargetRemote $AemCloudGitUrl
  Write-Log INFO "Added remote '$TargetRemote'."
}

# Fetch target for divergence detection
try { git fetch $TargetRemote | Out-Null } catch { Write-Log WARN "Fetch target failed: $($_.Exception.Message)" }

# Collect local branches
$localBranches = git for-each-ref --format='%(refname:short)' refs/heads | Where-Object { $_ }
Write-Log INFO "Local branches count: $($localBranches.Count)"

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

$summary = [ordered]@{
  pushed        = @()
  forced        = @()
  skipped       = @()
  divergent     = @()
  errors        = @()
  dryRun        = $DryRun.IsPresent
  timestampUtc  = (Get-Date).ToUniversalTime().ToString('o')
}

# Determine divergence & classify
$branchesToNormalPush = @()
$branchesToForcePush  = @()

foreach ($br in $localBranches) {
  if (Is-Excluded $br) {
    Write-Log INFO "Excluded '$br'."
    $summary.skipped += $br
    continue
  }

  if (-not $SkipUntrackedCheck) {
    git show-ref --verify --quiet "refs/remotes/$SourceRemote/$br"
    if (-not $LASTEXITCODE -eq 0) {
      Write-Log WARN "No upstream tracking ref for '$br'; skipping (use -SkipUntrackedCheck to include)."
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
      Write-Log INFO "Divergent branch '$br'."
    }
  }

  if ($divergent) {
    if ($AllowForce) {
      $branchesToForcePush += $br
    } else {
      Write-Log ERROR "Divergent without force: '$br' skipped."
      $summary.skipped += $br
    }
  } else {
    $branchesToNormalPush += $br
  }
}

function Retry-Push {
  param([scriptblock]$Action,[string]$RefName)
  for ($i=1; $i -le $MaxRetries; $i++) {
    & $Action
    if ($LASTEXITCODE -eq 0) { return $true }
    $delay = [int]([math]::Pow(2, ($i-1))) * $RetryDelaySeconds
    Write-Log WARN "Push failed for '$RefName' attempt $i/$MaxRetries; retrying in $delay s."
    Start-Sleep -Seconds $delay
  }
  return $false
}

# Perform pushes
if ($DryRun) {
  Write-Log INFO "[DRY-RUN] Would normal-push: $([string]::Join(', ', $branchesToNormalPush))"
  Write-Log INFO "[DRY-RUN] Would force-push: $([string]::Join(', ', $branchesToForcePush))"
} else {
  # Normal pushes can be batched; create refspec list
  if ($branchesToNormalPush.Count -gt 0) {
    Write-Log INFO "Pushing normal branches: $([string]::Join(', ', $branchesToNormalPush))"
    $refspecs = $branchesToNormalPush | ForEach-Object { "$_:$_" }
    $normalAction = { git push $TargetRemote $refspecs }
    if (Retry-Push -Action $normalAction -RefName "batch-normal") {
      $summary.pushed += $branchesToNormalPush
    } else {
      Write-Log ERROR "Batch push failed; falling back to per-branch."
      foreach ($b in $branchesToNormalPush) {
        $act = { git push $TargetRemote "$b:$b" }
        if (Retry-Push -Action $act -RefName $b) {
          $summary.pushed += $b
        } else {
          $summary.errors += "$b (push failed)"
          Write-Log ERROR "Failed to push '$b'."
        }
      }
    }
  }

  foreach ($b in $branchesToForcePush) {
    Write-Log INFO "Force pushing '$b'."
    $forceAction = { git push $TargetRemote "+$b:$b" }
    if (Retry-Push -Action $forceAction -RefName $b) {
      $summary.forced += $b
    } else {
      $summary.errors += "$b (force push failed)"
      Write-Log ERROR "Force push failed '$b'."
    }
  }
}

# Tags
if ($DryRun) {
  Write-Log INFO "[DRY-RUN] Would push tags."
} else {
  Write-Log INFO "Pushing tags."
  $tagAction = { git push $TargetRemote --tags }
  if (-not (Retry-Push -Action $tagAction -RefName 'tags')) {
    Write-Log WARN "Tag push failed."
    $summary.errors += "tags (push failed)"
  }
}

# Write summary JSON
if ($SummaryPath) {
  try {
    ($summary | ConvertTo-Json -Depth 5) | Out-File -Encoding UTF8 $SummaryPath
    Write-Log INFO "Summary written to $SummaryPath."
  } catch {
    Write-Log WARN "Failed to write summary JSON: $($_.Exception.Message)"
  }
}

# Cleanup credential helper (optional)
git config --local --unset credential.helper | Out-Null

if ($summary.errors.Count -gt 0) {
  Write-Log ERROR "Sync completed with errors."
  exit 1
}

Write-Log INFO "Enhanced sync completed successfully."
exit 0