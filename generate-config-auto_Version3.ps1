param(
  [Parameter(Mandatory=$true)][int]$ProgramId,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$ClientSecret,
  [Parameter(Mandatory=$true)][string]$OrgId,
  [string]$Scopes = "cloudmanager",
  [string]$RegionHost = "ims-na1.adobelogin.com",
  [string]$OutputPath = "config-aem-environments.json"
)

$body = @{
  grant_type    = "client_credentials"
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = $Scopes
}
$token = (Invoke-RestMethod -Method Post -Uri "https://$RegionHost/ims/token/v3" -Body $body).access_token

$pipesResp = Invoke-RestMethod -Method Get -Uri "https://cloudmanager.adobe.io/api/program/$ProgramId/pipelines" -Headers @{
  Authorization     = "Bearer $token"
  "x-api-key"       = $ClientId
  "x-gw-ims-org-id" = $OrgId
  Accept            = "application/json"
}

$pipes = $pipesResp._embedded.pipelines
$mapping = @{}
foreach($p in $pipes){
  $n = $p.name.ToLower()
  $key = if($n -match "dev"){ "dev" }
         elseif($n -match "stag"){ "stage" }
         elseif($n -match "prod"){ "prod" }
         else { $n -replace '[^a-z0-9]+','-' }

  # Avoid collisions
  $orig = $key
  $i = 2
  while($mapping.ContainsKey($key)){
    $key = "$orig-$i"
    $i++
  }
  $mapping[$key] = $p.id
}

($mapping | ConvertTo-Json) | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Generated mapping:"
$mapping