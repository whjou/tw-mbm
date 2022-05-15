#-----------------------------------------------------------
# TiddlyWiki Monitor, Backup and Move
# See https://powershell.one/tricks/filesystem/filesystemwatcher
#-----------------------------------------------------------
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process



#-----------------------------------------------------------
# Customize as required
#-----------------------------------------------------------
$global:MonitorPath  = $Env:USERPROFILE + "\Downloads"
$global:BackupPath   = $Env:USERPROFILE + "\Downloads\tmp"
$global:MovePath     = $Env:USERPROFILE + "\Downloads\tmp"
$global:MonitorFile  = "TWtest.html"
# -1 = infinite, 0 = none, +n = n copies
$global:BackupNumber = 5



#-----------------------------------------------------------
$global:MonitorFileName      = [System.IO.Path]::GetFileNameWithoutExtension($MonitorFile)
$global:MonitorFileExtension = [System.IO.Path]::GetExtension($MonitorFile)
$global:MonitorWaitTimeout   = 5
$MonitorSubdirectories       = $false
$MonitorFileFilter           = '*'
$MonitorFilter               = [IO.NotifyFilters]::FileName, [IO.NotifyFilters]::LastWrite



#-----------------------------------------------------------
# Retain $BackupNumber backups and remove the rest
#-----------------------------------------------------------
function Prune-Backup {
  param (
    [Parameter()]
    [string]$MonitorFile,
    [string]$BackupPath,
    [int]   $BackupNumber
  )
  $MonitorFileName      = [System.IO.Path]::GetFileNameWithoutExtension($MonitorFile)
  $MonitorFileExtension = [System.IO.Path]::GetExtension($MonitorFile)

  If ($BackupNumber -ne -1) {
    # Process existing files
    Get-ChildItem -Path $BackupPath -File `
      | Where-Object {
          $_.Name.StartsWith($MonitorFileName+"-") -and $_.Name.EndsWith($MonitorFileExtension)
        } `
      | Sort -Descending LastWriteTime `
      | Select-Object -Skip $BackupNumber `
      | ForEach-Object {
          $MatchingFile = $_.FullName
          $ts           = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
          $logEntry     = "[{0}] Deleting {1}..." -f $ts, $MatchingFile
          Write-Host $logEntry
          Remove-Item $MatchingFile
        }
  }
}



#-----------------------------------------------------------
function Backup-File {
  param (
    [Parameter()]
    [string]$MovePath,
    [string]$MonitorFile,
    [string]$BackupPath
  )

  $MonitorFileName      = [System.IO.Path]::GetFileNameWithoutExtension($MonitorFile)
  $MonitorFileExtension = [System.IO.Path]::GetExtension($MonitorFile)

  # Backup original file
  $lwt = (Get-Item "$MovePath\$MonitorFile").LastWriteTime.ToString("yyyyMMddHHmmss")
  $src = "{0}\{1}{2}"     -f $MovePath  , $MonitorFileName,       $MonitorFileExtension
  $tgt = "{0}\{1}-{2}{3}" -f $BackupPath, $MonitorFileName, $lwt, $MonitorFileExtension

  If (Test-Path -Path $tgt -PathType Leaf) {
    # $tgt exists, so check if contents is the same
    $tgtMD5 = Get-FileHash -Path $tgt -Algorithm MD5
    $srcMD5 = Get-FileHash -Path $src -Algorithm MD5
    If ($tgtMD5 -eq $srcMD5) {
      # contents same, so not moving, just delete src
      $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
      $logEntry = "[{0}] ... {1} already backed up as {2}, deleting..." -f $ts, $src, $tgt
      Write-Host $logEntry
      Remove-Item $src
      $tgt = $null
    } Else {
      # contents different, so append MD5
      $tgt = "{0}\{1}-{2}-{3}{4}" -f $BackupPath, $MonitorFileName, $lwt, $tgtMD5, $MonitorFileExtension
    }
  }

  If ($tgt -ne $null) {
    $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logEntry = "[{0}] ... Moving {1} to {2}..." -f $ts, $src, $tgt
    Write-Host $logEntry

    Move-Item -Path $src -Destination $tgt -Force
  }
}



#-----------------------------------------------------------
function global:Move-MatchingFile {
  param (
    [Parameter()]
    [string]$MatchingFile,
    [string]$MovePath,
    [string]$MonitorFile,
    [int]   $BackupNumber
  )

  Backup-File `
    -MovePath    $MovePath `
    -MonitorFile $MonitorFile `
    -BackupPath  $BackupPath

  # Move renamed file to source location
  $src = $MatchingFile
  $tgt = "{0}\{1}{2}" -f $MovePath, $MonitorFileName, $MonitorFileExtension

  $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $logEntry = "[{0}] ... Moving {1} to {2}..." -f $ts, $src, $tgt
  Write-Host $logEntry

  Move-Item -Path $src -Destination $tgt

  Prune-Backup `
    -MonitorFile  $MonitorFile `
    -BackupPath   $BackupPath `
    -BackupNumber $BackupNumber
}



#-----------------------------------------------------------
# Main
#-----------------------------------------------------------
try
{
  $MonitorWatcher = New-Object -TypeName System.IO.FileSystemWatcher -Property @{
    Path                  = $MonitorPath
    Filter                = $MonitorFileFilter
    IncludeSubdirectories = $MonitorSubdirectories
    NotifyFilter          = $MonitorFilter
  }



  # Action callback : BEGIN
  $MonitorAction = {
    $details     = $event.SourceEventArgs
    $Name        = $details.Name
    $FullPath    = $details.FullPath
    $OldFullPath = $details.OldFullPath
    $OldName     = $details.OldName
    $ChangeType  = $details.ChangeType
    $Timestamp   = $event.TimeGenerated

    Write-Host ""

    # Execute code based on change type here:
    switch ($ChangeType)
      {
        'Renamed' {
          if ($Name.StartsWith($MonitorFileName) -and $Name.EndsWith($MonitorFileExtension)) {
            # Renamed file matched
            $MatchingFile = $FullPath

            $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            $logEntry = "[{0}] Matching file detected : {1}..." -f $ts, $MatchingFile
            Write-Host $logEntry

            Move-MatchingFile `
              -MatchingFile   $MatchingFile `
              -MovePath       $MovePath `
              -MonitorFile    $MonitorFile `
              -BackupNumber   $BackupNumber
          }
        }
      }
  }
  # Action callback : END



  # Process existing files, maybe from incomplete previous runs
  Get-ChildItem -Path $MonitorPath -File `
    | Where-Object {
        $_.Name.StartsWith($MonitorFileName) -and $_.Name.EndsWith($MonitorFileExtension)
      } `
    | Sort LastWriteTime `
    | ForEach-Object {
        $MatchingFile = $_.FullName

        $ts           = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $logEntry     = "[{0}] Matching file found : {1}..." -f $ts, $MatchingFile
        Write-Host $logEntry

        Move-MatchingFile `
          -MatchingFile   $MatchingFile `
          -MovePath       $MovePath `
          -MonitorFile    $MonitorFile `
          -BackupNumber   $BackupNumber
      }

  # Start monitoring
  $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $logEntry = "[{0}] Monitoring {1}\ for {2}*{3} : BEGIN" -f $ts, $MonitorPath, $MonitorFileName, $MonitorFileExtension
  Write-Host  $logEntry

  $handlers = . {
    Register-ObjectEvent -InputObject $MonitorWatcher -EventName Renamed  -Action $MonitorAction
  }
  $MonitorWatcher.EnableRaisingEvents = $true

  # Idle loop till CTRL-C
  do
    {
      Wait-Event -Timeout $MonitorWaitTimeout
      Write-Host "." -NoNewline

  } while ($true)
}

finally
{
  # When user presses CTRL-C:
  Write-Host ""

  # Stop monitoring
  $MonitorWatcher.EnableRaisingEvents = $false
  $handlers | ForEach-Object {
    Unregister-Event -SourceIdentifier $_.Name
  }
  $handlers | Remove-Job
  $MonitorWatcher.Dispose()

  $ts       = $(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $logEntry = "[{0}] Monitoring {1}\ for {2}*{3} : END" -f $ts, $MonitorPath, $MonitorFileName, $MonitorFileExtension
  Write-Host $logEntry
}
#-----------------------------------------------------------
