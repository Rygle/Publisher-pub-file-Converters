pt<# :

@echo off

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"

exit /b

#>

# Batch convert Microsoft Publisher files to PDF.
# Save as a .bat file including the above, and either run from a CMD shell or double-click in Windows File Explorer.
# Tested in Windows 11 with PowerShell 7 and Publisher 365 pre-13-10-2026 EOL date.
# May work as is or with minor modifications with earlier versions of Publisher after the EOL date.

param(

    [string]$RootFolder,

    [string]$OutputFolder,

    [int]$TimeoutSeconds = 60  # Maximum time allowed per file in seconds

)

function Select-Folder {

    param([string]$Title)

    $shell = New-Object -ComObject Shell.Application

    try {

        $folder = $shell.BrowseForFolder(0, $Title, 0, 0)

        if ($null -eq $folder) { return $null }

        return $folder.Self.Path

    }

    finally {

        if ($shell -ne $null) {

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)

        }

    }

}

# 1. Prompt for Source Folder

if ([string]::IsNullOrWhiteSpace($RootFolder)) {

    $RootFolder = Select-Folder -Title 'Choose the SOURCE folder containing your Microsoft Publisher files'

}

if ([string]::IsNullOrWhiteSpace($RootFolder)) {

    Write-Host 'No source folder was selected. Operation cancelled.' -ForegroundColor Yellow

    exit

}

if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) {

    Write-Error "The source folder does not exist: $RootFolder"

    exit 1

}

# 2. Prompt for Destination Folder

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {

    $OutputFolder = Select-Folder -Title 'Choose the DESTINATION folder where PDFs should be saved'

}

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {

    Write-Host 'No destination folder was selected. Operation cancelled.' -ForegroundColor Yellow

    exit

}

# Ensure destination directory exists

if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {

    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

}

# --- AUTOMATIC SECURITY SUPPRESSION & TRUSTED LOCATION REGISTRATION ---

try {

    $officeVersions = @('16.0', '15.0', '14.0', '12.0')

    foreach ($ver in $officeVersions) {

        # 1. Configure Publisher Security Keys

        $pubSecPath = "HKCU:\Software\Microsoft\Office\$ver\Publisher\Security"

        if (-not (Test-Path $pubSecPath)) { New-Item -Path $pubSecPath -Force | Out-Null }

        

        Set-ItemProperty -Path $pubSecPath -Name "VBAWarnings" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $pubSecPath -Name "AllowAutomation" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $pubSecPath -Name "DisableAllSecurityWarnings" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        # 2. Configure Global Office Security Keys

        $commonSecPath = "HKCU:\Software\Microsoft\Office\$ver\Common\Security"

        if (-not (Test-Path $commonSecPath)) { New-Item -Path $commonSecPath -Force | Out-Null }

        

        Set-ItemProperty -Path $commonSecPath -Name "VBAWarnings" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $commonSecPath -Name "UIGrasshopper" -Value 0 -Type DWord -ErrorAction SilentlyContinue

        # 3. Add Selected Source Folder as an Office Trusted Location

        $trustedPath = "HKCU:\Software\Microsoft\Office\$ver\Publisher\Security\Trusted Locations\BatchConverterSource"

        if (-not (Test-Path $trustedPath)) { New-Item -Path $trustedPath -Force | Out-Null }

        

        Set-ItemProperty -Path $trustedPath -Name "Path" -Value $RootFolder -Type String -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $trustedPath -Name "AllowSubfolders" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $trustedPath -Name "Description" -Value "Batch PDF Converter Source Folder" -Type String -ErrorAction SilentlyContinue

    }

    Write-Host "Source folder dynamically registered as an Office Trusted Location." -ForegroundColor Gray

}

catch {

    Write-Warning "Unable to set registry security overrides: $($_.Exception.Message)"

}

# 0. Kill lingering Publisher processes from previous runs/hangs

Get-Process -Name "MSPUB" -ErrorAction SilentlyContinue | Stop-Process -Force

# SINGLE LOG FILE in the ROOT of the Source Directory

$LogFile = Join-Path $RootFolder ('Publisher-to-PDF-{0}.log' -f (Get-Date -Format 'yyyy-MM-dd-HHmmss'))

# Recurse through all subfolders

$pubFiles = Get-ChildItem -LiteralPath $RootFolder -Filter '*.pub' -File -Recurse

Write-Host ''

Write-Host "Starting conversion..." -ForegroundColor Green

Write-Host "Source:      $RootFolder"

Write-Host "Destination: $OutputFolder"

Write-Host "Timeout:     $TimeoutSeconds seconds per file"

Write-Host "Log file:    $LogFile"

Write-Host "Files found: $($pubFiles.Count)"

Write-Host '--------------------------------------------------'

"Started: $(Get-Date)" | Out-File -LiteralPath $LogFile -Encoding utf8

"Source folder: $RootFolder" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Destination folder: $OutputFolder" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Timeout threshold: $TimeoutSeconds seconds" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Publisher files found: $($pubFiles.Count)" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

'--------------------------------------------------' | Out-File -LiteralPath $LogFile -Append -Encoding utf8

'' | Out-File -LiteralPath $LogFile -Append -Encoding utf8

$converted = 0

$skipped = 0

$failed = 0

foreach ($file in $pubFiles) {

    if ($file.DirectoryName.Length -gt $RootFolder.Length) {

        $relativePath = $file.DirectoryName.Substring($RootFolder.Length).TrimStart('\', '/')

    } else {

        $relativePath = ""

    }

    $targetDir = Join-Path $OutputFolder $relativePath

    if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {

        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null

    }

    $pdfFileName = [System.IO.Path]::ChangeExtension($file.Name, '.pdf')

    $pdfPath = Join-Path $targetDir $pdfFileName

    # Skip existing valid PDFs (> 0 bytes)

    if (Test-Path -LiteralPath $pdfPath) {

        $existingFile = Get-Item -LiteralPath $pdfPath

        if ($existingFile.Length -gt 0) {

            Write-Host "Skipping existing PDF: $pdfPath" -ForegroundColor Yellow

            "SKIPPED | $(Get-Date -Format 'HH:mm:ss') | $($file.FullName) | Valid PDF already exists" |

                Out-File -LiteralPath $LogFile -Append -Encoding utf8

            $skipped++

            continue

        } else {

            Write-Host "Overwriting partial/0-byte PDF: $pdfPath" -ForegroundColor Cyan

            Remove-Item -LiteralPath $pdfPath -Force -ErrorAction SilentlyContinue

        }

    }

    Write-Host "Converting: $($file.FullName)" -ForegroundColor Cyan

    # Background scriptblock for conversion AND popup dismissing

    $conversionScript = {

        param($filePath, $targetPdfPath)

        $PdfFormat = 2 # pbFixedFormatTypePDF

        $publisher = $null

        $document = $null

        # Sub-job to automatically click through Font Substitution dialogs if they pop up

        $popupWatcher = Start-Job -ScriptBlock {

            $wshell = New-Object -ComObject WScript.Shell

            for ($i = 0; $i -lt 60; $i++) {

                Start-Sleep -Milliseconds 500

                

                # Check for Font Substitution / Publisher Alert dialog windows

                $win = Get-Process -Name "MSPUB" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne "" }

                if ($win -and ($win.MainWindowTitle -match "Font" -or $win.MainWindowTitle -match "Substitution" -or $win.MainWindowTitle -match "Publisher")) {

                    $wshell.AppActivate($win.Id)

                    $wshell.SendKeys("{ENTER}") # Presses OK to accept default font substitution

                }

            }

        }

        try {

            # Unblock file streams

            Unblock-File -LiteralPath $filePath -ErrorAction SilentlyContinue

            $publisher = New-Object -ComObject Publisher.Application

            

            # Disable GUI alerts and warning popups

            try { $publisher.Option("ShowServiceWarnings") = $false } catch {}

            # Open document read-only

            $document = $publisher.Open($filePath, $true, $false)

            $document.ExportAsFixedFormat($PdfFormat, $targetPdfPath)

            return "SUCCESS"

        }

        catch {

            $msg = $_.Exception.Message

            if ($msg -match "Publisher has detected a problem" -or $msg -match "protect your computer") {

                return "BLOCKED_BY_OFFICE_SECURITY: Protected View blocked this legacy or corrupted file."

            }

            return "ERROR: $msg"

        }

        finally {

            Stop-Job $popupWatcher -ErrorAction SilentlyContinue

            Remove-Job $popupWatcher -Force -ErrorAction SilentlyContinue

            if ($document -ne $null) {

                try { $document.Close() } catch {}

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)

            }

            if ($publisher -ne $null) {

                try { $publisher.Quit() } catch {}

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($publisher)

            }

            [GC]::Collect()

            [GC]::WaitForPendingFinalizers()

        }

    }

    # Start conversion in a background job

    $job = Start-Job -ScriptBlock $conversionScript -ArgumentList $file.FullName, $pdfPath

    $completed = Wait-Job $job -Timeout $TimeoutSeconds

    if ($null -eq $completed) {

        # TIMEOUT OCCURRED

        Stop-Job $job -ErrorAction SilentlyContinue

        Remove-Job $job -Force -ErrorAction SilentlyContinue

        

        # Kill stuck MSPUB.exe process to release handle locks

        Get-Process -Name "MSPUB" -ErrorAction SilentlyContinue | Stop-Process -Force

        # Clean up partial PDF if created

        if (Test-Path -LiteralPath $pdfPath) {

            Remove-Item -LiteralPath $pdfPath -Force -ErrorAction SilentlyContinue

        }

        Write-Warning "TIMED OUT: Exceeded $TimeoutSeconds seconds (Font popup or corruption hang). Skipping: $($file.Name)"

        "FAILED_TIMEOUT | $(Get-Date -Format 'HH:mm:ss') | $($file.FullName) | Exceeded $TimeoutSeconds seconds threshold" |

            Out-File -LiteralPath $LogFile -Append -Encoding utf8

        $failed++

    }

    else {

        # JOB COMPLETED WITHIN TIMEOUT

        $result = Receive-Job $job

        Remove-Job $job -Force -ErrorAction SilentlyContinue

        if ($result -eq "SUCCESS" -and (Test-Path -LiteralPath $pdfPath)) {

            Write-Host "Created: $pdfPath" -ForegroundColor Green

            "OK | $(Get-Date -Format 'HH:mm:ss') | $($file.FullName) | $pdfPath" |

                Out-File -LiteralPath $LogFile -Append -Encoding utf8

            $converted++

        }

        elseif ($result -match "BLOCKED_BY_OFFICE_SECURITY") {

            Write-Warning "BLOCKED: Publisher security stopped $($file.Name) from opening."

            "FAILED_BLOCKED | $(Get-Date -Format 'HH:mm:ss') | $($file.FullName) | $result" |

                Out-File -LiteralPath $LogFile -Append -Encoding utf8

            $failed++

        }

        else {

            Write-Warning "Failed: $($file.FullName)"

            Write-Warning $result

            "FAILED_ERROR | $(Get-Date -Format 'HH:mm:ss') | $($file.FullName) | $result" |

                Out-File -LiteralPath $LogFile -Append -Encoding utf8

            $failed++

        }

    }

}

'' | Out-File -LiteralPath $LogFile -Append -Encoding utf8

'--------------------------------------------------' | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Finished: $(Get-Date)" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Converted: $converted" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Skipped: $skipped" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

"Total Failed/Blocked/Timed Out: $failed" | Out-File -LiteralPath $LogFile -Append -Encoding utf8

Write-Host '--------------------------------------------------'

Write-Host 'Finished.' -ForegroundColor Green

Write-Host "Converted: $converted"

Write-Host "Skipped:   $skipped"

Write-Host "Failed:    $failed"

Write-Host "Log file:  $LogFile"

Write-Host ''

Read-Host -Prompt "Press Enter to exit"
