<# : 2>nul
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression $([System.IO.File]::ReadAllText('%~f0'))"
exit /b
#>

# Batch convert Microsoft Publisher .pub files to PDF or ODG using LibreOffice.
# Save as a .bat file including the above, and either run from a CMD shell or double-click in Windows File Explorer.
# The code below is PowerShell code, but the above CMD/DOS code allows the double-click functionality
# Tested in Windows 11 with PowerShell 7 and LibreOffice 26.8 with a default install.
# Requires Windows PowerShell / PowerShell 7

Add-Type -AssemblyName System.Windows.Forms ;

Write-Host "========================================================" -ForegroundColor Cyan ;
Write-Host "     LibreOffice Publisher Converter (PDF & ODG)" -ForegroundColor Cyan ;
Write-Host "========================================================" -ForegroundColor Cyan ;
Write-Host "" ;

# Clean up any lingering soffice background processes on startup
Write-Host "Cleaning up any existing LibreOffice processes..." -ForegroundColor Yellow ;
$null = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /IM soffice.exe /IM soffice.bin /T" -NoNewWindow -Wait -ErrorAction SilentlyContinue ;

# 1. Select Source Directory
Write-Host "Step 1: Select SOURCE folder containing Publisher (.pub) files..." -ForegroundColor Yellow ;
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog ;
$folderBrowser.Description = "Select SOURCE folder containing Publisher (.pub) files" ;
if ($folderBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No source folder selected. Exiting." -ForegroundColor Red ;
    pause ;
    exit ;
}
$srcDir = $folderBrowser.SelectedPath ;
Write-Host "Source Folder: $srcDir`n" -ForegroundColor Green ;

# 2. Select Output Directory
Write-Host "Step 2: Select OUTPUT folder (Cancel = Same as source)..." -ForegroundColor Yellow ;
$outBrowser = New-Object System.Windows.Forms.FolderBrowserDialog ;
$outBrowser.Description = "Select OUTPUT folder (Cancel to mirror inside source)" ;
if ($outBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $outDirRoot =$outBrowser.SelectedPath ;
} else {
    $outDirRoot =$srcDir ;
}
Write-Host "Output Folder: $outDirRoot`n" -ForegroundColor Green ;

# 3. Recursive Prompt (30s Timeout, Default = Y)
$recursive = $true ;
Write-Host "Scan subdirectories recursively? [Y/n] (Auto-Yes in 30s): " -NoNewline -ForegroundColor Yellow ;
$deadline = (Get-Date).AddSeconds(30) ;
$userMadeChoice = $false ;
while ((Get-Date) -lt ($deadline)) {
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true) ;
        if (($key.KeyChar -eq 'n') -or ($key.KeyChar -eq 'N')) {
            $recursive = $false ;
            Write-Host "N" -ForegroundColor Red ;
            $userMadeChoice = $true ;
            break ;
        } elseif (($key.Key -eq 'Enter') -or ($key.KeyChar -eq 'y') -or ($key.KeyChar -eq 'Y')) {
            Write-Host "Y" -ForegroundColor Green ;
            $userMadeChoice = $true ;
            break ;
        }
    }
    Start-Sleep -Milliseconds 100 ;
}
if (-not ($userMadeChoice)) { Write-Host "Y (Timeout Default)" -ForegroundColor Green ; }

# 4. Format Prompt (30s Timeout, Default = Both)
Write-Host "`nSelect target conversion format:" ;
Write-Host "  [1] Both PDF & ODG (Default)" ;
Write-Host "  [2] PDF only" ;
Write-Host "  [3] ODG only" ;
Write-Host "Choice [1/2/3] (Auto-selecting [1] in 30s): " -NoNewline -ForegroundColor Yellow ;

$doPdf =$true ;
$doOdg =$true ;
$deadline = (Get-Date).AddSeconds(30) ;
$userMadeChoice =$false ;
while ((Get-Date) -lt ($deadline)) {
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true) ;
        if ($key.KeyChar -eq '2') { $doOdg =$false ; Write-Host "2 (PDF Only)" -ForegroundColor Green ; $userMadeChoice =$true ; break ; }
        if ($key.KeyChar -eq '3') { $doPdf =$false ; Write-Host "3 (ODG Only)" -ForegroundColor Green ; $userMadeChoice =$true ; break ; }
        if (($key.Key -eq 'Enter') -or ($key.KeyChar -eq '1')) { Write-Host "1 (Both)" -ForegroundColor Green ; $userMadeChoice =$true ; break ; }
    }
    Start-Sleep -Milliseconds 100 ;
}
if (-not ($userMadeChoice)) { Write-Host "1 (Timeout Default)" -ForegroundColor Green ; }

# 5. Find LibreOffice
$soffice =$null ;
$p1 = "$env:ProgramFiles\LibreOffice\program\soffice.exe" ;
$p2 = "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe" ;
$p3 = "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe" ;

if ([System.IO.File]::Exists($p1)) { $soffice =$p1 ; }
elseif ([System.IO.File]::Exists($p2)) { $soffice =$p2 ; }
elseif ([System.IO.File]::Exists($p3)) { $soffice =$p3 ; }

if (-not ($soffice)) {
    Write-Host "`nERROR: LibreOffice (soffice.exe) was not found!" -ForegroundColor Red ;
    pause ;
    exit ;
}

# 6. Logging Setup
$timestamp = Get-Date -Format "yyyy-MM-dd-HHmmss" ;
$logFile = Join-Path ($outDirRoot) ("LibreOffice-PUB-Converter-" + $timestamp + ".log") ;
Set-Content -Path ($logFile) -Value ("Started: " + (Get-Date)) ;

Write-Host "`n--------------------------------------------------" ;
Write-Host "Executable: $soffice" ;
Write-Host "Log File:   $logFile" ;
Write-Host "--------------------------------------------------`n" ;

# 7. Scan Files & Filter System Trash Folders
Write-Host "Indexing Publisher (.pub) files in $srcDir..." -ForegroundColor Yellow ;
$searchOption = if ($recursive) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly } ;
$rawFiles = [System.IO.Directory]::GetFiles($srcDir, "*.pub", $searchOption) ;

$pubFiles = [System.Collections.Generic.List[string]]::new() ;
for ($i = 0; $i -lt ($rawFiles.Count); $i++) {
    $f = $rawFiles[$i] ;
    if (($f -notlike "*\`$RECYCLE.BIN\*") -and ($f -notlike "*\System Volume Information\*")) {
        $pubFiles.Add($f) ;
    }
}

if ($pubFiles.Count -eq 0) {
    Write-Host "No valid .pub files found." -ForegroundColor Red ;
    pause ;
    exit ;
}

Write-Host "Found $($pubFiles.Count) file(s) ready to process.`n" -ForegroundColor Green ;

# 8. Conversion Loop
$pdfCount = 0 ; $odgCount = 0 ; $skipCount = 0 ; $failCount = 0 ;
$maxTimeoutSeconds = 45 ; # Per-file conversion timeout limit

# Temp User Profile Directory for isolated headless execution
$tempProfileDir = Join-Path ($env:TEMP) "LO_PUB_Profile" ;
$profileArg = "-env:UserInstallation=file:///" + ($tempProfileDir.Replace("\", "/")) ;

for ($i = 0; $i -lt ($pubFiles.Count); $i++) {
    $file = $pubFiles[$i] ;
    $index = $i + 1 ;
    $fileItem = Get-Item -LiteralPath ($file) ;
    $parentDir = $fileItem.DirectoryName ;

    # Determine Target Output Subfolder Structure
    if ($outDirRoot.Equals($srcDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        $targetDir = $parentDir ;
    } else {
        $relativePath = $parentDir.Substring($srcDir.Length).TrimStart('\') ;
        $targetDir = Join-Path ($outDirRoot) ($relativePath) ;
    }

    if (-not ([System.IO.Directory]::Exists($targetDir))) {
        $null = New-Item -ItemType Directory -Path ($targetDir) -Force ;
    }

    $pdfPath = Join-Path ($targetDir) ($fileItem.BaseName + ".pdf") ;
    $odgPath = Join-Path ($targetDir) ($fileItem.BaseName + ".odg") ;

    $hasPdf = ([System.IO.File]::Exists($pdfPath)) -and ((New-Object System.IO.FileInfo($pdfPath)).Length -gt 0) ;
    $hasOdg = ([System.IO.File]::Exists($odgPath)) -and ((New-Object System.IO.FileInfo($odgPath)).Length -gt 0) ;

    $pdfNeeded = ($doPdf) -and (-not ($hasPdf)) ;
    $odgNeeded = ($doOdg) -and (-not ($hasOdg)) ;

    if ((-not ($pdfNeeded)) -and (-not ($odgNeeded))) {
        Write-Host "[$index/$($pubFiles.Count)] Skipping (Outputs exist): $($fileItem.Name)" -ForegroundColor DarkGray ;
        Add-Content -Path ($logFile) -Value ("SKIPPED - " + $fileItem.FullName) ;
        $skipCount++ ;
        continue ;
    }

    Write-Host "[$index/$($pubFiles.Count)] Processing: $($fileItem.FullName)" -ForegroundColor Cyan ;

    # Convert to PDF
    if ($pdfNeeded) {
        Write-Host "   -> Converting to PDF..." -NoNewline ;
        $argList = @($profileArg, "--headless", "--invisible", "--nologo", "--nofirststartwizard", "--norestore", "--convert-to", "pdf", "`"$($fileItem.FullName)`"", "--outdir", "`"$targetDir`"") ;
        $proc = Start-Process -FilePath ($soffice) -ArgumentList ($argList) -PassThru -NoNewWindow ;
        
        $finished = $proc.WaitForExit($maxTimeoutSeconds * 1000) ;
        if (-not ($finished)) {
            Write-Host " [TIMED OUT]" -ForegroundColor Red ;
            $null = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /IM soffice.exe /IM soffice.bin /T" -NoNewWindow -Wait -ErrorAction SilentlyContinue ;
            $failCount++ ;
            Add-Content -Path ($logFile) -Value ("TIMEOUT PDF - " + $fileItem.FullName) ;
        } else {
            if (([System.IO.File]::Exists($pdfPath)) -and ((New-Object System.IO.FileInfo($pdfPath)).Length -gt 0)) {
                Write-Host " [OK]" -ForegroundColor Green ;
                $pdfCount++ ;
            } else {
                Write-Host " [FAILED]" -ForegroundColor Red ;
                $failCount++ ;
                Add-Content -Path ($logFile) -Value ("FAILED PDF - " + $fileItem.FullName) ;
            }
        }
    }

    # Convert to ODG
    if ($odgNeeded) {
        Write-Host "   -> Converting to ODG..." -NoNewline ;
        $argList = @($profileArg, "--headless", "--invisible", "--nologo", "--nofirststartwizard", "--norestore", "--convert-to", "odg", "`"$($fileItem.FullName)`"", "--outdir", "`"$targetDir`"") ;
        $proc = Start-Process -FilePath ($soffice) -ArgumentList ($argList) -PassThru -NoNewWindow ;
        
        $finished = $proc.WaitForExit($maxTimeoutSeconds * 1000) ;
        if (-not ($finished)) {
            Write-Host " [TIMED OUT]" -ForegroundColor Red ;
            $null = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /IM soffice.exe /IM soffice.bin /T" -NoNewWindow -Wait -ErrorAction SilentlyContinue ;
            $failCount++ ;
            Add-Content -Path ($logFile) -Value ("TIMEOUT ODG - " + $fileItem.FullName) ;
        } else {
            if (([System.IO.File]::Exists($odgPath)) -and ((New-Object System.IO.FileInfo($odgPath)).Length -gt 0)) {
                Write-Host " [OK]" -ForegroundColor Green ;
                $odgCount++ ;
            } else {
                Write-Host " [FAILED]" -ForegroundColor Red ;
                $failCount++ ;
                Add-Content -Path ($logFile) -Value ("FAILED ODG - " + $fileItem.FullName) ;
            }
        }
    }
}

Write-Host "`n================ SUMMARY RESULTS ================" -ForegroundColor Cyan ;
Write-Host "PDFs Created: $pdfCount" ;
Write-Host "ODGs Created: $odgCount" ;
Write-Host "Skipped:      $skipCount" ;
Write-Host "Failed:       $failCount" ;
Write-Host "Log File:     $logFile" ;

# Clean up temp profile directory on exit
if ([System.IO.Directory]::Exists($tempProfileDir)) {
    try { Remove-Item -Path ($tempProfileDir) -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}

pause ;
