<#
.SYNOPSIS
  Export a Word document to PDF with Microsoft Word COM automation.

.DESCRIPTION
  Used by the dsh-render-document plugin as its "word" engine. The plugin runs
  inside the DSH host process, which is not confined by the Windows sandbox, so
  Word starts with the user's normal token and can create its working files.
  (Word started from a sandboxed command cannot: it fails with "Word could not
  create the work file. Check the TEMP environment variable.")

  Prints exactly one JSON object on stdout:
    {"ok":true,"word":"<exe>","version":"16.0"}                       for -Check
    {"ok":true,"pages":3,"pdf":"C:\\path\\out.pdf","version":"16.0"}  for an export
    {"ok":false,"error":"..."}                                        on failure

  Exit codes: 0 success, 2 Word not installed, 3 export failed.

.PARAMETER Check
  Only report whether Microsoft Word is installed.

.PARAMETER Source
  Document to open read-only (absolute path).

.PARAMETER Target
  PDF to create (absolute path, must not exist).

.PARAMETER TimeoutSeconds
  Reserved for callers that enforce the deadline themselves.
#>
[CmdletBinding()]
param(
  [switch]$Check,
  [string]$Source = '',
  [string]$Target = '',
  [int]$TimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'

function Write-Result($value) {
  Write-Output ($value | ConvertTo-Json -Compress)
}

function Find-Word {
  $candidates = @()
  foreach ($key in @(
      'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe',
      'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\winword.exe')) {
    try {
      $value = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
      if ($value) { $candidates += $value }
    } catch { }
  }
  foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if (-not $root) { continue }
    foreach ($relative in @(
        'Microsoft Office\Root\Office16\WINWORD.EXE',
        'Microsoft Office\root\Office16\WINWORD.EXE',
        'Microsoft Office\Office16\WINWORD.EXE',
        'Microsoft Office\Root\Office15\WINWORD.EXE',
        'Microsoft Office\Office15\WINWORD.EXE')) {
      $candidates += (Join-Path $root $relative)
    }
  }
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
  }
  return ''
}

function Get-WinWordIds {
  @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

$wordExe = Find-Word
if (-not $wordExe) {
  Write-Result ([pscustomobject]@{ ok = $false; error = 'Microsoft Word is not installed' })
  exit 2
}

if ($Check) {
  Write-Result ([pscustomobject]@{ ok = $true; word = $wordExe })
  exit 0
}

if (-not $Source -or -not $Target) {
  Write-Result ([pscustomobject]@{ ok = $false; error = 'Source and Target are required' })
  exit 3
}
if (-not (Test-Path -LiteralPath $Source)) {
  Write-Result ([pscustomobject]@{ ok = $false; error = "Source not found: $Source" })
  exit 3
}

$before = Get-WinWordIds
$word = $null
$doc = $null
$owns = $false
try {
  $word = New-Object -ComObject Word.Application
  $after = Get-WinWordIds
  $owns = @($after | Where-Object { $before -notcontains $_ }).Count -gt 0
  if ($owns) { try { $word.Visible = $false } catch { } }
  try { $word.DisplayAlerts = 0 } catch { }
  try { $word.Options.ConfirmConversions = $false } catch { }
  try { $word.Options.SaveNormalPrompt = $false } catch { }

  $doc = $word.Documents.Open($Source, $false, $true, $false)
  $pages = 0
  try { $pages = [int]$doc.ComputeStatistics(2) } catch { $pages = 0 }

  $targetDir = Split-Path -Parent $Target
  if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
  }
  if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Force }
  $doc.ExportAsFixedFormat($Target, 17)

  $version = ''
  try { $version = [string]$word.Version } catch { $version = '' }
  if (-not (Test-Path -LiteralPath $Target)) {
    Write-Result ([pscustomobject]@{ ok = $false; error = 'Word reported success but no PDF appeared' })
    exit 3
  }
  Write-Result ([pscustomobject]@{ ok = $true; pages = $pages; pdf = $Target; version = $version })
  exit 0
} catch {
  Write-Result ([pscustomobject]@{ ok = $false; error = $_.Exception.Message })
  exit 3
} finally {
  if ($doc) { try { $doc.Close(0) } catch { } }
  if ($word) {
    # Never close a Word instance the user already had open.
    if ($owns) { try { $word.Quit() } catch { } }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
