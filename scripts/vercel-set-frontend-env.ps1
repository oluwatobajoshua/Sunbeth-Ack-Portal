param(
  [Parameter(Mandatory=$true)][string]$Project,                 # e.g. sunbeth-ack-portal
  [Parameter(Mandatory=$false)][string]$Scope,                  # optional Vercel org scope (team slug)
  [Parameter(Mandatory=$true)][string]$ApiBase,                 # e.g. https://sunbeth-ack-portal-backend-spny.vercel.app
  [Parameter(Mandatory=$false)][string]$AdminEmails             # optional: comma-separated admin emails
)

$ErrorActionPreference = 'Stop'

function Add-Env([string]$Name, [string]$Value) {
  Write-Host "Setting $Name (production)" -ForegroundColor Cyan
  $args = "env add `"$Name`" production"
  if ($Scope) { $args = $args + " --scope `"$Scope`"" }

  $pinfo = New-Object System.Diagnostics.ProcessStartInfo
  $pinfo.FileName = 'cmd.exe'
  $pinfo.Arguments = "/c vercel $args"
  $pinfo.RedirectStandardInput = $true
  $pinfo.RedirectStandardOutput = $true
  $pinfo.RedirectStandardError = $true
  $pinfo.UseShellExecute = $false
  $pinfo.WorkingDirectory = $script:RepoRoot

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $pinfo
  if (-not $p.Start()) { throw "Failed to start vercel env add for $Name" }
  $p.StandardInput.Write($Value)
  $p.StandardInput.Close()
  $out = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  $p.WaitForExit()
  if ($p.ExitCode -ne 0) {
    if ($out) { Write-Host $out -ForegroundColor DarkGray }
    if ($err) { Write-Host $err -ForegroundColor Red }
    if (($out -match 'already exists') -or ($err -match 'already exists')) {
      Write-Host "Variable $Name already exists. Attempting to remove and re-add..." -ForegroundColor Yellow
      Remove-Env -Name $Name
      return Add-Env -Name $Name -Value $Value
    }
    if (($out -match 'already been added to all Environments') -or ($err -match 'already been added to all Environments')) {
      Write-Host "Variable $Name exists in All Environments. Removing globally and re-adding..." -ForegroundColor Yellow
      Remove-Env -Name $Name -All
      return Add-Env -Name $Name -Value $Value
    }
    throw "vercel env add failed for $Name (exit $($p.ExitCode))"
  }
  if ($out) { Write-Host $out -ForegroundColor Green }
}

function Remove-Env([string]$Name, [switch]$All) {
  Write-Host "Removing $Name (production)" -ForegroundColor DarkYellow
  $args = if ($All) { "env rm `"$Name`" -y" } else { "env rm `"$Name`" production -y" }
  if ($Scope) { $args = $args + " --scope `"$Scope`"" }

  $pinfo = New-Object System.Diagnostics.ProcessStartInfo
  $pinfo.FileName = 'cmd.exe'
  $pinfo.Arguments = "/c vercel $args"
  $pinfo.RedirectStandardOutput = $true
  $pinfo.RedirectStandardError = $true
  $pinfo.UseShellExecute = $false
  $pinfo.WorkingDirectory = $script:RepoRoot

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $pinfo
  if (-not $p.Start()) { throw "Failed to start vercel env rm for $Name" }
  $out = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  $p.WaitForExit()
  if ($p.ExitCode -ne 0) {
    if ($out) { Write-Host $out -ForegroundColor DarkGray }
    if ($err) { Write-Host $err -ForegroundColor Red }
    throw "vercel env rm failed for $Name (exit $($p.ExitCode))"
  }
  if ($out) { Write-Host $out -ForegroundColor DarkGray }
}

# 0) Ensure CLI available and link project
vercel --version | Out-Null
try { Set-Location (Split-Path $PSScriptRoot -Parent) } catch {}
$script:RepoRoot = (Get-Location).Path
$linkCmd = "vercel link --yes --project `"$Project`"" + ($(if($Scope){" --scope `"$Scope`""} else {''}))
Write-Host "Linking to project '$Project'..." -ForegroundColor Cyan
& cmd.exe /c $linkCmd

# 1) Set required envs
Add-Env -Name 'REACT_APP_API_BASE' -Value $ApiBase
if ($AdminEmails) {
  Add-Env -Name 'REACT_APP_ADMIN_EMAILS' -Value $AdminEmails
}

# 2) Show current env
$envCmd = "vercel env ls" + ($(if($Scope){" --scope `"$Scope`""} else {''}))
& cmd.exe /c $envCmd | Select-String -Pattern 'production' | ForEach-Object { $_.Line }

# 3) Deploy frontend
Write-Host "Triggering Production deploy..." -ForegroundColor Cyan
$deployCmd = "vercel deploy --prod --yes" + ($(if($Scope){" --scope `"$Scope`""} else {''}))
& cmd.exe /c $deployCmd
Write-Host "Done." -ForegroundColor Green
