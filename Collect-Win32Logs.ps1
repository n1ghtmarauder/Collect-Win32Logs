<#
.SYNOPSIS
  Intune Win32 IME Log Collector - produces a ZIP directly compatible with
  the Win32 Deployment Analyzer.

.DESCRIPTION
  Collects Win32 app deployment artifacts and packages them into a flat ZIP
  with ODC-compatible filenames so the analyzer can parse them without any
  modifications.

  Collected artifacts:
    Sidecar logs (ALL rolled logs merged oldest-first per base name, no time filter):
      {COMPUTER}_AppWorkload.log
      {COMPUTER}_AppActionProcessor.log
      {COMPUTER}_IntuneManagementExtension.log
      {COMPUTER}_AgentExecutor.log

    Delivery Optimization log (Get-DeliveryOptimizationLog cmdlet output):
      {COMPUTER}_Get-DeliveryOptimizationLog.txt

    Registry exports (full key trees):
      {COMPUTER}_REG_SW_Microsoft_IntuneManagementExtension.txt
      {COMPUTER}_REG_SW_Microsoft_EnterpriseDesktopAppManagement.txt

    Application Event Log (XML, time-filtered):
      {COMPUTER}_Application.xml

    Store Event Log (XML, time-filtered):
      {COMPUTER}_Store.xml

    AppxDeployment-Server Event Log (XML, time-filtered):
      {COMPUTER}_AppxDeployment-Server.xml

    BITS Event Log (XML, time-filtered):
      {COMPUTER}_BITS.xml

  Output ZIP: <COMPUTERNAME>_Win32Logs_<yyyyMMdd-HHmmss>.zip

.PARAMETER DaysBack
  Days of history for the Application event log export (default 7).
  IME Sidecar logs are collected in full - rolled logs are essential for
  proper iteration tracking and are often older than N days.

.PARAMETER OutputRoot
  Folder where the ZIP will be created (default: current user Desktop).

.PARAMETER MaxAppEvents
  Cap for Application event log XML export. 0 = no cap (default 500).

.PARAMETER NoDOLog
  Skip Delivery Optimization log collection.

.PARAMETER NoZip
  Keep the staging folder and skip ZIP creation.

.PARAMETER NoOpen
  Do not open Explorer to the ZIP after creation.

.NOTES
  Run elevated (as Administrator) for best results.
  The resulting ZIP can be dropped directly into the Win32 Deployment
  Analyzer alongside or instead of an ODC ZIP.
#>

[CmdletBinding()]
param(
  [int]$DaysBack       = 7,
  [string]$OutputRoot  = [Environment]::GetFolderPath('Desktop'),
  [int]$MaxAppEvents   = 500,
  [switch]$NoDOLog,
  [switch]$NoZip,
  [switch]$NoOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
#  Setup
# ---------------------------------------------------------------
$comp      = $env:COMPUTERNAME
$now       = Get-Date
$stamp     = $now.ToString('yyyyMMdd-HHmmss')
$tempDir   = Join-Path $env:TEMP ('{0}_Win32Logs_{1}' -f $comp, $stamp)
$zipName   = '{0}_Win32Logs_{1}.zip' -f $comp, $stamp
$zipPath   = Join-Path $OutputRoot $zipName

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$collectorLog = Join-Path $tempDir ('{0}_Collector.log' -f $comp)

function Log {
  param([string]$Msg, [string]$Level = 'INFO')
  $line = '{0} [{1}] {2}' -f (Get-Date).ToString('HH:mm:ss'), $Level, $Msg
  $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
  Write-Host $line -ForegroundColor $color
  Add-Content -Path $collectorLog -Value $line -Encoding UTF8
}

function Safe {
  param([scriptblock]$Block, [string]$Label)
  try { & $Block }
  catch { Log ('{0} - {1}' -f $Label, $_.Exception.Message) 'ERROR' }
}

Log '=== Win32 IME Log Collector ==='
Log ('Computer  : {0}' -f $comp)
Log ('Timestamp : {0}' -f $now.ToString('s'))
Log ('DaysBack  : {0} - event logs only; IME logs collected in full' -f $DaysBack)
Log ('Output    : {0}' -f $OutputRoot)
Log ''

# ---------------------------------------------------------------
#  1. IME Sidecar Logs - merged, no time filter
# ---------------------------------------------------------------
$imeLogDir = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'

if (Test-Path -LiteralPath $imeLogDir -PathType Container) {
  Log ('Collecting IME Sidecar logs from: {0}' -f $imeLogDir)

  # @() forces array even if Get-ChildItem returns 0 or 1 result
  [array]$logFiles = @(Get-ChildItem -LiteralPath $imeLogDir -File -Filter '*.log' -ErrorAction SilentlyContinue)
  $groups = @{}

  foreach ($f in $logFiles) {
    if ($f.BaseName -match '^(?<base>.+)-\d{8}-\d{6}$') {
      $base = $Matches['base']
    } else {
      $base = $f.BaseName
    }
    if (-not $groups.ContainsKey($base)) {
      $groups[$base] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    }
    $groups[$base].Add($f)
  }

  foreach ($base in @($groups.Keys | Sort-Object)) {
    $list = $groups[$base]
    if ($null -eq $list -or $list.Count -eq 0) { continue }

    # Sort: rolled logs by embedded timestamp oldest first, current log last
    # @() ensures result is always an array
    [array]$sorted = @($list | Sort-Object @{
      Expression = {
        if ($_.BaseName -match '-(?<d>\d{8})-(?<t>\d{6})$') {
          $ds = $Matches['d']; $ts = $Matches['t']
          try { [datetime]::ParseExact(('{0}{1}' -f $ds, $ts), 'yyyyMMddHHmmss', $null) }
          catch { $_.LastWriteTime }
        } else {
          [datetime]::MaxValue
        }
      }
    })

    $outFile = Join-Path $tempDir ('{0}_{1}.log' -f $comp, $base)

    $writer = $null
    try {
      $writer = [System.IO.StreamWriter]::new($outFile, $false, [System.Text.Encoding]::UTF8)

      foreach ($f in $sorted) {
        $marker = '>>>>> FILE: {0} | Size: {1} | Modified: {2} <<<<<' -f $f.Name, $f.Length, $f.LastWriteTime.ToString('s')
        $writer.WriteLine($marker)
        $reader = $null
        $fs = $null
        try {
          # FileShare.ReadWrite allows reading files locked by IME service
          $fs = [System.IO.FileStream]::new(
            $f.FullName,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
          )
          $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8, $true)
          while ($null -ne ($line = $reader.ReadLine())) {
            $writer.WriteLine($line)
          }
        } catch {
          $writer.WriteLine(('>>>>> ERROR reading {0}: {1} <<<<<' -f $f.Name, $_.Exception.Message))
          Log ('  WARNING: Could not read {0}: {1}' -f $f.Name, $_.Exception.Message) 'WARN'
        } finally {
          # StreamReader.Close() also closes the underlying FileStream
          if ($null -ne $reader) { $reader.Close() }
          elseif ($null -ne $fs) { $fs.Close() }
        }
        $writer.WriteLine('')
      }
    } finally {
      if ($null -ne $writer) { $writer.Close() }
    }

    $totalKB = [math]::Round(($sorted | Measure-Object -Property Length -Sum).Sum / 1KB, 1)
    Log ('  {0}: {1} file(s) merged -> {2}_{0}.log - {3} KB' -f $base, $sorted.Count, $comp, $totalKB)
  }
} else {
  Log ('IME Logs directory not found: {0}' -f $imeLogDir) 'WARN'
}

# ---------------------------------------------------------------
#  2. Delivery Optimization Log
# ---------------------------------------------------------------
if (-not $NoDOLog) {
  Log 'Collecting Delivery Optimization log...'
  Safe -Label 'DO Log' -Block {
    $doFile = Join-Path $tempDir ('{0}_Get-DeliveryOptimizationLog.txt' -f $comp)
    [array]$doEntries = @(Get-DeliveryOptimizationLog -ErrorAction Stop)
    if ($doEntries.Count -gt 0) {
      $doEntries | Format-List TimeCreated, LevelName, Function, Message |
        Out-String -Width 300 |
        Set-Content -Path $doFile -Encoding UTF8
      Log ('  DO log: {0} entries collected' -f $doEntries.Count) 'OK'
    } else {
      Log '  DO log: No entries returned' 'WARN'
    }
  }
}

# ---------------------------------------------------------------
#  3. Registry Exports
# ---------------------------------------------------------------
Log 'Exporting registry keys...'

$regKeys = @(
  @{
    Path = 'HKLM\SOFTWARE\Microsoft\IntuneManagementExtension'
    Name = 'REG_SW_Microsoft_IntuneManagementExtension'
  }
  @{
    Path = 'HKLM\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement'
    Name = 'REG_SW_Microsoft_EnterpriseDesktopAppManagement'
  }
)

foreach ($rk in $regKeys) {
  Safe -Label ('Registry: {0}' -f $rk.Path) -Block {
    $outFile = Join-Path $tempDir ('{0}_{1}.txt' -f $comp, $rk.Name)
    & reg.exe export $rk.Path $outFile /y 2>$null | Out-Null
    if (Test-Path -LiteralPath $outFile) {
      $sz = [math]::Round((Get-Item -LiteralPath $outFile).Length / 1KB, 1)
      Log ('  {0} - {1} KB' -f $rk.Name, $sz) 'OK'
    } else {
      Log ('  Key not found or empty: {0}' -f $rk.Path) 'WARN'
    }
  }
}

# ---------------------------------------------------------------
#  4. Event Logs (XML, time-filtered)
# ---------------------------------------------------------------
$ms = [Math]::Abs($DaysBack) * 86400000
$xpath = '*[System[TimeCreated[timediff(@SystemTime) <= {0}]]]' -f $ms

$eventLogs = @(
  @{ Channel = 'Application';                                          FileName = 'Application';          Cap = $MaxAppEvents }
  @{ Channel = 'Microsoft-Windows-Store/Operational';                  FileName = 'Store';                Cap = 0 }
  @{ Channel = 'Microsoft-Windows-AppXDeploymentServer/Operational';   FileName = 'AppxDeployment-Server'; Cap = 0 }
  @{ Channel = 'Microsoft-Windows-Bits-Client/Operational';            FileName = 'BITS';                 Cap = 0 }
)

foreach ($el in $eventLogs) {
  Log ('Exporting {0} - last {1} days...' -f $el.Channel, $DaysBack)
  Safe -Label $el.Channel -Block {
    $outFile = Join-Path $tempDir ('{0}_{1}.xml' -f $comp, $el.FileName)

    $wevtArgs = @('qe', $el.Channel, ('/q:{0}' -f $xpath), '/f:xml')
    if ($el.Cap -gt 0) { $wevtArgs += ('/c:{0}' -f $el.Cap) }

    [array]$raw = @(& wevtutil @wevtArgs 2>$null)

    if ($raw.Count -gt 0) {
      $wrapped = '<Events>' + "`r`n" + ($raw -join "`r`n") + "`r`n" + '</Events>'
      [System.IO.File]::WriteAllText($outFile, $wrapped, [System.Text.Encoding]::UTF8)
      $evtCount = ([regex]::Matches($wrapped, '<Event[\s>]')).Count
      $sz = [math]::Round((Get-Item -LiteralPath $outFile).Length / 1KB, 1)
      Log ('  {0}: {1} events - {2} KB' -f $el.FileName, $evtCount, $sz) 'OK'
    } else {
      Log ('  {0}: No events in time range' -f $el.FileName) 'WARN'
    }
  }
}

# ---------------------------------------------------------------
#  5. Manifest
# ---------------------------------------------------------------
Log ''
Log '=== Collection Summary ==='
[array]$collectedFiles = @(Get-ChildItem -LiteralPath $tempDir -File | Sort-Object Name)
foreach ($cf in $collectedFiles) {
  if ($cf.Length -ge 1MB) {
    $sz = '{0:N2} MB' -f ($cf.Length / 1MB)
  } else {
    $sz = '{0:N1} KB' -f ($cf.Length / 1KB)
  }
  Log ('  {0}  - {1}' -f $cf.Name, $sz)
}
Log ('Total files: {0}' -f $collectedFiles.Count)

# ---------------------------------------------------------------
#  6. Create ZIP
# ---------------------------------------------------------------
if ($NoZip) {
  Log ''
  Log ('NoZip specified - staging folder: {0}' -f $tempDir)
} else {
  Log ''
  Log 'Creating ZIP...'

  if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
      $tempDir, $zipPath,
      [System.IO.Compression.CompressionLevel]::Optimal,
      $false
    )
  } catch {
    Log ('ZipFile failed, falling back to Compress-Archive: {0}' -f $_.Exception.Message) 'WARN'
    Compress-Archive -Path (Join-Path $tempDir '*') -DestinationPath $zipPath -Force
  }

  $zipSizeMB = [math]::Round((Get-Item -LiteralPath $zipPath).Length / 1MB, 2)
  Log ('ZIP created: {0} - {1} MB' -f $zipName, $zipSizeMB) 'OK'

  # Cleanup staging AFTER final log writes
  Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

  if (-not $NoOpen -and (Test-Path -LiteralPath $zipPath)) {
    explorer.exe /select,"$zipPath"
  }

  Write-Host ''
  Write-Host '----------------------------------------------------' -ForegroundColor Cyan
  Write-Host ('  {0}' -f $zipPath) -ForegroundColor Green
  Write-Host '  Upload this ZIP file using the File Share link you were provided' -ForegroundColor Cyan
  Write-Host '----------------------------------------------------' -ForegroundColor Cyan
}
