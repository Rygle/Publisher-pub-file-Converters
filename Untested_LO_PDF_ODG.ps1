# Requires PowerShell 7+ (Cross-Platform for Windows, macOS, and Linux)
# WARNING, NOT TESTED!!!! This is Powershell 7 code that may or may not successfully automate LibreOffice on Windows, Mac or Linux.
# I HAVE NOT TESTED THIS!!!!! You will have to install Powershell 7. Save to a ps1 file e.g. 'lopub2pdfodg.ps1' and run 'pwsh lopub2pdfodg.ps1' or similar.

# ---------------------------------------------------------------------------

# Helper Functions for Cross-Platform Compatibility

# ---------------------------------------------------------------------------

function Stop-LibreOfficeProcesses {

    Get-Process -Name "soffice", "soffice.bin" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue ;

}

function Select-FolderDialog ($promptTitle) {

    if ($IsWindows) {

        Add-Type -AssemblyName System.Windows.Forms ;

        $browser = New-Object System.Windows.Forms.FolderBrowserDialog ;

        $browser.Description =$promptTitle ;

        if ($browser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return$browser.SelectedPath ; }

    } elseif ($IsMacOS) {$appleScript = "posix path of (choose folder with prompt `"$promptTitle`")" ;

        $res = osascript -e $appleScript 2>$null ;

        if (-not [string]::IsNullOrWhiteSpace($res)) { return$res.Trim() ; }

    } elseif ($IsLinux) {

        if (Get-Command "zenity" -ErrorAction SilentlyContinue) {

            $res = zenity --file-selection --directory --title="$promptTitle" 2>$null ;

            if (-not [string]::IsNullOrWhiteSpace($res)) { return$res.Trim() ; }

        }

    }

    

    # Fallback Terminal Prompt for headless or unsupported environments

    Write-Host "`n$promptTitle" -ForegroundColor Yellow ;

    $path = Read-Host "Enter directory path (or press Enter to cancel)" ;

    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) { return $null ; }

    return (Get-Item $path).FullName ;

}

function Find-LibreOfficeExecutable {

    if ($IsWindows) {

        $paths = @(

            "$env:ProgramFiles\LibreOffice\program\soffice.exe",

            "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",

            "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe"

        ) ;

        foreach ($p in $paths) { if ([System.IO.File]::Exists($p)) { return $p ; } }

    } elseif ($IsMacOS) {

        $macPath = "/Applications/LibreOffice.app/Contents/MacOS/soffice" ;

        if ([System.IO.File]::Exists($macPath)) { return $macPath ; }

    } elseif ($IsLinux) {

        $cmd = Get-Command "soffice" -ErrorAction SilentlyContinue ;

        if ($cmd) { return $cmd.Source ; }

        $linuxPaths = @("/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice") ;

        foreach ($p in $linuxPaths) { if ([System.IO.File]::Exists($p)) { return $p ; } }

    }

    return $null ;

}

# ---------------------------------------------------------------------------

# Main Execution Flow

# ---------------------------------------------------------------------------

Write-Host "========================================================" -ForegroundColor Cyan ;

Write-Host "  LibreOffice Publisher Converter (Cross-Platform)" -ForegroundColor Cyan ;

Write-Host "========================================================" -ForegroundColor Cyan ;

Write-Host "" ;

# Startup Cleanup

Write-Host "Cleaning up any existing LibreOffice processes..." -ForegroundColor Yellow ;

Stop-LibreOfficeProcesses ;

# 1. Select Source Directory

$srcDir = Select-FolderDialog "Select SOURCE folder containing Publisher (.pub) files" ;

if (-not $srcDir) {

    Write-Host "No valid source folder selected. Exiting." -ForegroundColor Red ;

    exit ;

}

Write-Host "Source Folder: $srcDir`n" -ForegroundColor Green ;

# 2. Select Output Directory

$outDirRoot = Select-FolderDialog "Select OUTPUT folder (Cancel = Same as source)" ;

if (-not $outDirRoot) { $outDirRoot =$srcDir ; }

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

$soffice = Find-LibreOfficeExecutable ;

if (-not ($soffice)) {

    Write-Host "`nERROR: LibreOffice (soffice) executable was not found on this system!" -ForegroundColor Red ;

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

# 7. Scan Files Phase with Progress Bar

Write-Host "Phase 1: Indexing Publisher (.pub) files in $srcDir..." -ForegroundColor Yellow ;

Write-Progress -Activity "Phase 1/2: Indexing Files" -Status "Scanning directory tree..." -PercentComplete 0 ;

$searchOption = if ($recursive) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly } ;

$rawFiles = [System.IO.Directory]::GetFiles($srcDir, "*.pub", $searchOption) ;

$pubFiles = [System.Collections.Generic.List[string]]::new() ;

$totalRaw = $rawFiles.Count ;

for ($i = 0; $i -lt ($totalRaw); $i++) {

    $f = $rawFiles[$i] ;

    $indexNum = $i + 1 ;

    $percent = [math]::Floor(($indexNum / [math]::Max(1, $totalRaw)) * 100) ;

    Write-Progress -Activity "Phase 1/2: Indexing Files" -Status "Filtering file $indexNum of $totalRaw ($percent%)" -PercentComplete $percent ;

    # Ignore Trash and System folders across Windows, macOS, and Linux

    if (($f -notlike "*\`$RECYCLE.BIN\*") -and ($f -notlike "*\System Volume Information\*") -and ($f -notlike "*/.Trashes/*") -and ($f -notlike "*/.Trash-*/*")) {

        $pubFiles.Add($f) ;

    }

}

Write-Progress -Activity "Phase 1/2: Indexing Files" -Completed ;

if ($pubFiles.Count -eq 0) {

    Write-Host "No valid .pub files found." -ForegroundColor Red ;

    exit ;

}

Write-Host "Found $($pubFiles.Count) valid file(s) ready to process.`n" -ForegroundColor Green ;

# 8. Conversion Phase with Progress Bar

$pdfCount = 0 ; $odgCount = 0 ; $skipCount = 0 ; $failCount = 0 ;

$maxTimeoutSeconds = 45 ;

$tempProfileDir = Join-Path ([System.IO.Path]::GetTempPath()) "LO_PUB_Profile" ;

$profileArg = "-env:UserInstallation=file:///" + ($tempProfileDir.Replace("\", "/")) ;

$totalPub = $pubFiles.Count ;

for ($i = 0; $i -lt ($totalPub); $i++) {

    $file = $pubFiles[$i] ;

    $index = $i + 1 ;

    $percent = [math]::Floor(($index / [math]::Max(1, $totalPub)) * 100) ;

    $fileItem = Get-Item -LiteralPath ($file) ;

    $parentDir = $fileItem.DirectoryName ;

    Write-Progress -Activity "Phase 2/2: Converting Files" -Status "Processing file $index of $totalPub ($percent% complete)" -CurrentOperation $fileItem.Name -PercentComplete $percent ;

    # Determine Target Output Subfolder Structure

    if ($outDirRoot.Equals($srcDir, [System.StringComparison]::OrdinalIgnoreCase)) {

        $targetDir = $parentDir ;

    } else {

        $relativePath = $parentDir.Substring($srcDir.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar) ;

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

        Write-Host "[$index/$totalPub] Skipping (Outputs exist): $($fileItem.Name)" -ForegroundColor DarkGray ;

        Add-Content -Path ($logFile) -Value ("SKIPPED - " + $fileItem.FullName) ;

        $skipCount++ ;

        continue ;

    }

    Write-Host "[$index/$totalPub] Processing ($percent%): $($fileItem.FullName)" -ForegroundColor Cyan ;

    # Convert to PDF

    if ($pdfNeeded) {

        Write-Host "   -> Converting to PDF..." -NoNewline ;

        $argList = @($profileArg, "--headless", "--invisible", "--nologo", "--nofirststartwizard", "--norestore", "--convert-to", "pdf", "`"$($fileItem.FullName)`"", "--outdir", "`"$targetDir`"") ;

        $proc = Start-Process -FilePath ($soffice) -ArgumentList ($argList) -PassThru -NoNewWindow ;

        

        $finished = $proc.WaitForExit($maxTimeoutSeconds * 1000) ;

        if (-not ($finished)) {

            Write-Host " [TIMED OUT]" -ForegroundColor Red ;

            Stop-LibreOfficeProcesses ;

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

            Stop-LibreOfficeProcesses ;

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

Write-Progress -Activity "Phase 2/2: Converting Files" -Completed ;

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
