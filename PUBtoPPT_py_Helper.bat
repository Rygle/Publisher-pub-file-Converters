:: Batch convert Microsoft Publisher files to PPTX using the excellent https://github.com/dllmr/PUBtoPPT.
:: Requires https://github.com/dllmr/PUBtoPPT/blob/main/pub2pptx.py in the same directory
:: Save as a .bat file including the above, and either run from a CMD shell or double-click in Windows File Explorer.
:: Tested in Windows 11 with PowerShell 7.

@echo off
title PUB to PPTX Batch Converter

powershell -NoExit -ExecutionPolicy Bypass -Command ^
    "$recursive = $true;" ^
    "Write-Host 'Scan all subdirectories? [Y/n] (Auto-selecting Yes in 30 seconds...): ' -NoNewline -ForegroundColor Yellow;" ^
    "$secondsLeft = 30;" ^
    "while ($secondsLeft -gt 0) {" ^
    "    if ([Console]::KeyAvailable) {" ^
    "        $key = [Console]::ReadKey($true);" ^
    "        if ($key.Key -eq 'Enter') { Write-Host 'Y (User selection)'; break };" ^
    "        $char = $key.KeyChar.ToString().ToLower();" ^
    "        if ($char -eq 'n') { $recursive = $false; Write-Host 'N (User selection)'; break };" ^
    "        if ($char -eq 'y') { $recursive = $true; Write-Host 'Y (User selection)'; break };" ^
    "    }" ^
    "    Start-Sleep -Milliseconds 200;" ^
    "    $secondsLeft -= 0.2;" ^
    "};" ^
    "if ($secondsLeft -le 0) { Write-Host 'Y (Timed out - default selected)' -ForegroundColor Gray };" ^
    "Write-Host '';" ^
    "$HOME_PATH = $env:USERPROFILE;" ^
    "$SCRIPT_DIR = Get-Location;" ^
    "if ($recursive) {" ^
    "    Write-Host 'Scanning current directory AND subdirectories for Microsoft Publisher (.pub) files...' -ForegroundColor Yellow;" ^
    "    $getFiles = { Get-ChildItem -Recurse -Filter *.pub -ErrorAction SilentlyContinue };" ^
    "} else {" ^
    "    Write-Host 'Scanning ONLY the current directory for Microsoft Publisher (.pub) files...' -ForegroundColor Yellow;" ^
    "    $getFiles = { Get-ChildItem -Filter *.pub -ErrorAction SilentlyContinue };" ^
    "}" ^
    "$pubFiles = @(); $scanCounter = 0;" ^
    "& $getFiles | ForEach-Object {" ^
    "    $scanCounter++; $pubFiles += $_;" ^
    "    Write-Progress -Activity 'Scanning for Publisher Files' -Status ('Found ' + $scanCounter + ' file(s) so far...') -CurrentOperation $_.FullName;" ^
    "};" ^
    "Write-Progress -Activity 'Scanning for Publisher Files' -Completed;" ^
    "$totalFiles = $pubFiles.Count; $currentIndex = 0;" ^
    "if ($totalFiles -eq 0) { Write-Host 'No Publisher (.pub) files were found.' -ForegroundColor Red; return };" ^
    "Write-Host ('Found ' + $totalFiles + ' Publisher file(s) ready to convert.`n') -ForegroundColor Green;" ^
    "foreach ($file in $pubFiles) {" ^
    "    $currentIndex++;" ^
    "    $filePath = $file.FullName;" ^
    "    $pptxPath = [System.IO.Path]::ChangeExtension($filePath, '.pptx');" ^
    "    $percent = [math]::Round(($currentIndex / $totalFiles) * 100);" ^
    "    Write-Progress -Activity 'Converting Publisher Files to PowerPoint' -Status ('Processing file ' + $currentIndex + ' of ' + $totalFiles + ' (' + $percent + '%%)') -CurrentOperation $file.Name -PercentComplete $percent;" ^
    "    if (Test-Path $pptxPath) {" ^
    "        Write-Host ('Skipping (Already converted): ' + $filePath) -ForegroundColor DarkGray;" ^
    "        continue;" ^
    "    }" ^
    "    Write-Host ('Processing [' + $currentIndex + '/' + $totalFiles + ']: ' + $filePath) -ForegroundColor Cyan;" ^
    "    $job = Start-Job -ScriptBlock {" ^
    "        param($path, $userHome, $scriptDir)" ^
    "        Set-Location $scriptDir;" ^
    "        $uvPath = Join-Path $userHome '.local\bin\uv.exe';" ^
    "        $pyScript = Join-Path $scriptDir 'pub2pptx.py';" ^
    "        & $uvPath run --system-certs $pyScript $path 2>&1;" ^
    "    } -ArgumentList $filePath, $HOME_PATH, $SCRIPT_DIR;" ^
    "    $completed = Wait-Job $job -Timeout 60;" ^
    "    if ($null -eq $completed) {" ^
    "        Write-Warning ('TIMED OUT (60s exceeded): ' + $filePath + ' - skipping file.');" ^
    "        Stop-Job $job;" ^
    "    } else {" ^
    "        $rawOutput = Receive-Job $job;" ^
    "        foreach ($item in $rawOutput) {" ^
    "            $line = $item.ToString().Trim();" ^
    "            if ([string]::IsNullOrWhiteSpace($line)) { continue }" ^
    "            if ($line -match '^(warning:)' -or $line -match 'skipping image|no pages found in document') {" ^
    "                Write-Host ('  [Conversion Warning] ' + $line) -ForegroundColor Yellow;" ^
    "            } elseif ($line -match '^(Installed|Resolved|Audited|Built|Prepared|Downloaded|Using Python|Creating virtualenv)') {" ^
    "                Write-Host ('  [Environment] ' + $line) -ForegroundColor DarkGray;" ^
    "            } elseif ($line -match '^(Traceback|ValueError|FileNotFoundError|PermissionError|error:)' -or $item -is [System.Management.Automation.ErrorRecord]) {" ^
    "                Write-Host ('  [Conversion Error] ' + $line) -ForegroundColor Red;" ^
    "            } else {" ^
    "                Write-Host ('  [Info] ' + $line) -ForegroundColor Gray;" ^
    "            }" ^
    "        }" ^
    "    }" ^
    "    Remove-Job $job -Force;" ^
    "};" ^
    "Write-Progress -Activity 'Converting Publisher Files to PowerPoint' -Completed;"

echo.
echo ========================================================
echo Execution complete or stopped. Press any key to exit...
echo ========================================================
pause
