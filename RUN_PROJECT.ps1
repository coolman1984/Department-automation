$ErrorActionPreference = "SilentlyContinue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# Keep startup independent from broken machine-level Python settings. The
# project code adds vendor.zip itself, so a user's PYTHONPATH is not needed.
Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
$env:PYTHONPATH = "$root\vendor.zip;$root\vendor"

$script:candidates = @()
$script:seen = @{}
$script:tested = @{}
$script:diagnostics = @()
$script:lastCandidate = $null
$script:chosen = $null

function Add-PythonCandidate([string]$path, [string]$kind, [string]$label) {
  if (-not $path -or -not (Test-Path $path)) { return }
  $full = (Resolve-Path $path).Path
  $key = "$kind|$($full.ToLowerInvariant())"
  if (-not $script:seen.ContainsKey($key)) {
    $script:seen[$key] = $true
    $script:candidates += [pscustomobject]@{ Path = $full; Kind = $kind; Label = $label }
  }
}

function Add-PythonPaths($paths) {
  foreach ($path in $paths) { Add-PythonCandidate ([string]$path) "python" ([string]$path) }
}

# Launcher, PATH, registry, and the usual per-user/system locations.
$launcher = Get-Command py.exe
if ($launcher) {
  Add-PythonCandidate $launcher.Source "launcher" "$($launcher.Source) -3"
  # py -0p prints the real interpreter paths.  Prefer those over a launcher
  # that may point at a removed or incompatible installation.
  foreach ($line in (& $launcher.Source -0p 2>$null)) {
    if ($line -match '(?i)([A-Za-z]:\\.*python\.exe)\s*$') {
      Add-PythonCandidate $matches[1] "python" $matches[1]
    }
  }
}
Add-PythonPaths (& where.exe python.exe 2>$null)
$command = Get-Command python.exe
if ($command) { Add-PythonCandidate $command.Source "python" $command.Source }
Add-PythonPaths (Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
Add-PythonPaths (Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
Add-PythonPaths (Get-ChildItem "$env:ProgramFiles\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
Add-PythonPaths (Get-ChildItem "C:\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
Add-PythonPaths @(
  "$env:USERPROFILE\anaconda3\python.exe",
  "$env:USERPROFILE\miniconda3\python.exe",
  "$env:ProgramData\Anaconda3\python.exe",
  "$env:ProgramData\Miniconda3\python.exe"
)
foreach ($hive in "HKCU:\Software\Python\PythonCore", "HKLM:\Software\Python\PythonCore", "HKLM:\Software\WOW6432Node\Python\PythonCore") {
  Get-ChildItem $hive -ErrorAction SilentlyContinue | ForEach-Object {
    $path = (Get-ItemProperty "$($_.PSPath)\InstallPath" -ErrorAction SilentlyContinue).'(default)'
    if ($path) { Add-PythonCandidate (Join-Path $path "python.exe") "python" (Join-Path $path "python.exe") }
  }
}

function Test-PythonCandidate($candidate) {
  $script:lastCandidate = $candidate
  $testKey = "$($candidate.Kind)|$($candidate.Path.ToLowerInvariant())"
  if ($script:tested.ContainsKey($testKey)) { return $false }
  $script:tested[$testKey] = $true
  if ($candidate.Kind -eq "launcher") {
    & $candidate.Path -3 -E -c "import sys;assert sys.version_info >= (3,10)" 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }
    $report = (& $candidate.Path -3 -E "$root\CHECK_ENVIRONMENT.py" --brief 2>&1 | Out-String).Trim()
  } else {
    & $candidate.Path -E -c "import sys;assert sys.version_info >= (3,10)" 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }
    $report = (& $candidate.Path -E "$root\CHECK_ENVIRONMENT.py" --brief 2>&1 | Out-String).Trim()
  }
  if ($LASTEXITCODE -eq 0) {
    $script:chosen = $candidate
    return $true
  }
  if ($report) { $script:diagnostics += "$($candidate.Label): $report" }
  return $false
}

foreach ($candidate in $script:candidates) {
  if (Test-PythonCandidate $candidate) { break }
}

# Only if the normal discovery paths failed, search the user's and system
# program folders for a Python executable. This avoids a slow full-disk scan.
if (-not $script:chosen) {
  $searchRoots = @($env:USERPROFILE, $env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
  foreach ($base in $searchRoots) {
    $paths = Get-ChildItem -Path $base -Filter python.exe -File -Recurse -Depth 6 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    Add-PythonPaths $paths
  }
  foreach ($candidate in $script:candidates) {
    if ($script:chosen) { break }
    if (Test-PythonCandidate $candidate) { break }
  }
}

if ($script:chosen) {
  if ($script:chosen.Kind -eq "launcher") { & $script:chosen.Path -3 -E "$root\engine.py" }
  else { & $script:chosen.Path -E "$root\engine.py" }
  exit $LASTEXITCODE
}

Write-Host "No usable Python environment was found." -ForegroundColor Yellow
Write-Host "The project checked the Python launcher, PATH, registry, common folders, and a limited program-folder search." -ForegroundColor Yellow
if ($script:diagnostics.Count -gt 0) {
  Write-Host "Environment details:" -ForegroundColor Yellow
  $script:diagnostics | Select-Object -Last 3 | ForEach-Object { Write-Host "- $_" }
}
if ($script:lastCandidate) {
  Write-Host "Detailed check for the last Python found:" -ForegroundColor Yellow
  if ($script:lastCandidate.Kind -eq "launcher") { & $script:lastCandidate.Path -3 -E "$root\CHECK_ENVIRONMENT.py" }
  else { & $script:lastCandidate.Path -E "$root\CHECK_ENVIRONMENT.py" }
}
Write-Host "Install only the package(s) named above, then run START.bat again." -ForegroundColor Yellow
exit 1
