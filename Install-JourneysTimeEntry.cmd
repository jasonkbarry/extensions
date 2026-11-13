@echo off
title Journeys Time Entry installer
rem Double-click to install or update the Journeys Time Entry Chrome/Edge extension.
rem The PowerShell below is run from this same file (everything after the #__PS__ line).
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f='%~f0'; $s=[IO.File]::ReadAllText($f); $i=$s.LastIndexOf('#__'+'PS__'); Invoke-Expression $s.Substring($i)"
exit /b

#__PS__
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Site      = 'https://jasonkbarry.github.io/extensions'
$ExtName   = 'Journeys Time Entry'
$Default   = Join-Path $env:LOCALAPPDATA 'Nextiva Extensions\journeys-time-entry'

function Say($t, $c = 'Gray') { Write-Host $t -ForegroundColor $c }
Say ''
Say '  Journeys Time Entry installer' 'Cyan'
Say '  -----------------------------' 'Cyan'

try {
  # 1. What is the latest version?
  $info = Invoke-RestMethod -UseBasicParsing -Uri ("$Site/journeys-time-entry.json?t=" + [DateTime]::Now.Ticks)
  Say ("  Latest version: " + $info.version)

  # 2. Where is it installed already? Look at the unpacked extensions in every Chrome / Edge profile.
  $found = New-Object System.Collections.Generic.List[string]
  $roots = @("$env:LOCALAPPDATA\Google\Chrome\User Data", "$env:LOCALAPPDATA\Microsoft\Edge\User Data")
  foreach ($root in $roots) {
    if (-not (Test-Path $root)) { continue }
    foreach ($prof in (Get-ChildItem $root -Directory -ErrorAction SilentlyContinue)) {
      foreach ($pf in @('Preferences', 'Secure Preferences')) {
        $file = Join-Path $prof.FullName $pf
        if (-not (Test-Path $file)) { continue }
        try { $text = [IO.File]::ReadAllText($file) } catch { continue }
        foreach ($m in [regex]::Matches($text, '"path"\s*:\s*"([A-Za-z]:\\\\[^"]+)"')) {
          $p = $m.Groups[1].Value -replace '\\\\', '\'
          $mf = Join-Path $p 'manifest.json'
          if (Test-Path $mf) {
            try { if ((Get-Content $mf -Raw | ConvertFrom-Json).name -eq $ExtName) { $found.Add($p) } } catch {}
          }
        }
      }
    }
  }
  if (Test-Path (Join-Path $Default 'manifest.json')) { $found.Add($Default) }
  $targets = @($found | Sort-Object -Unique)
  $isNew = $targets.Count -eq 0
  if ($isNew) { $targets = @($Default) }

  # 3. Download and unpack the new version.
  $tmp = Join-Path $env:TEMP ('jte-' + [Guid]::NewGuid())
  New-Item -ItemType Directory -Path $tmp | Out-Null
  $zip = Join-Path $tmp 'jte.zip'
  Say ("  Downloading " + $info.zip + ' ...')
  Invoke-WebRequest -UseBasicParsing -Uri ("$Site/" + $info.zip) -OutFile $zip
  Expand-Archive -Path $zip -DestinationPath $tmp -Force
  $src = Join-Path $tmp 'journeys-time-entry'
  if (-not (Test-Path (Join-Path $src 'manifest.json'))) { throw 'The download did not contain the extension.' }

  # 4. Copy into each install folder (folder name is kept, so Chrome keeps your boards and settings).
  foreach ($t in $targets) {
    $old = '(new)'
    if (Test-Path (Join-Path $t 'manifest.json')) { try { $old = (Get-Content (Join-Path $t 'manifest.json') -Raw | ConvertFrom-Json).version } catch {} }
    New-Item -ItemType Directory -Path $t -Force | Out-Null
    Get-ChildItem $t -Filter 'VERSION *.txt' -ErrorAction SilentlyContinue | Remove-Item -Force
    # copy everything except manifest.json first, then manifest.json last, so the extension only reloads once the new code is in place
    robocopy "$src" "$t" /E /XF manifest.json /R:2 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "Could not copy files to $t (robocopy code $LASTEXITCODE). Is the folder open or read-only?" }
    Copy-Item (Join-Path $src 'manifest.json') -Destination $t -Force
    Say ("  Updated: $t  ($old -> " + $info.version + ')') 'Green'
  }
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

  Say ''
  if ($isNew) {
    Set-Clipboard -Value $targets[0]
    Say '  First-time setup (one time only):' 'Yellow'
    Say '    1. Chrome will open the Extensions page.'
    Say '    2. Turn on "Developer mode" (top right).'
    Say '    3. Click "Load unpacked", click the address bar of the folder window,'
    Say '       press Ctrl+V (the folder path is on your clipboard) and Enter, then "Select Folder".'
    Say ("       Folder: " + $targets[0])
    Say '    4. Pin "Journeys Time Entry" from the puzzle-piece menu.'
    try { Start-Process 'chrome.exe' 'chrome://extensions' } catch { Say '    (Open chrome://extensions yourself - Chrome was not found.)' }
  } else {
    Say '  Done. The extension reloads itself within about a minute' 'Green'
    Say '  (or right away the next time you open its panel).'
    Say '  Coming from 2.5 or earlier? Click the reload arrow on the extension in chrome://extensions once.'
    Say '  If the Burn Down script changed, the release notes on the site say so.'
  }
} catch {
  Say ''
  Say ('  Install failed: ' + $_.Exception.Message) 'Red'
  Say ('  You can also download the zip from ' + $Site + '/journeys-time-entry.html') 'Red'
}
Say ''
[void](Read-Host '  Press Enter to close')
