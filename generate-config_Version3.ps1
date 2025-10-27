param(
  [Parameter(Mandatory=$true)][int]$ProgramId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$OrgId,
  [string]$Scopes = "cloudmanager",
  [string]$RegionHost = "ims-na1.adobelogin.com",
  [string]$OutputPath = "config-aem-environments.json"
)

Write-Host "Getting access token..."
$body = @{
  grant_type    = "client_credentials"
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = $Scopes
}
$tokenResp = Invoke-RestMethod -Method Post -Uri "https://$RegionHost/ims/token/v3" -Body $body
$token = $tokenResp.access_token

Write-Host "Listing pipelines..."
$pipesResp = Invoke-RestMethod -Method Get -Uri "https://cloudmanager.adobe.io/api/program/$ProgramId/pipelines" -Headers @{
  Authorization     = "Bearer $token"
  "x-api-key"       = $ClientId
  "x-gw-ims-org-id" = $OrgId
  Accept            = "application/json"
}

$pipes = $pipesResp._embedded.pipelines
if(-not $pipes){
  Write-Error "No pipelines returned."
  exit 1
}

Write-Host "Found pipelines:"
$pipes | ForEach-Object {
  Write-Host ("  ID={0}  Name='{1}'  Type={2}  Branch={3}" -f $_.id, $_.name, $_.type, $_.branch)
}

$mapping = @{}
foreach($p in $pipes){
  $suggest = $null
  $n = $p.name.ToLower()
  if($n -match "dev"){ $suggest = "dev" }
  elseif($n -match "stag"){ $suggest = "stage" }
  elseif($n -match "prod"){ $suggest = "prod" }
  else { $suggest = ($n -replace '[^a-z0-9]+','-') }

  $answer = Read-Host "Map pipeline ID $($p.id) ('$($p.name)') to environment key (suggest: $suggest)"
  if([string]::IsNullOrWhiteSpace($answer)){ $answer = $suggest }
  if($mapping.ContainsKey($answer)){
    Write-Host "Environment key '$answer' already used; appending numeric suffix."
    $i = 2
    while($mapping.ContainsKey("$answer-$i")){ $i++ }
    $answer = "$answer-$i"
  }
  $mapping[$answer] = $p.id
}

Write-Host "Resulting mapping:"
$mapping.GetEnumerator() | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Key, $_.Value) }

($mapping | ConvertTo-Json) | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Saved $OutputPath"