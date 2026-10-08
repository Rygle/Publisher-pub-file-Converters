; AutoIt v3 script to click Enable on Microsoft Publisher Security Notices for batch conversion
; Use with https://www.autoitscript.com/
; Compile by right-clicking and select either just Compile Script (x64) or 'Show more options' -> Compile Script (x64)
; Double-click the compiled file to run in the background while batch converting .pub files.

#NoTrayIcon
Opt("WinTitleMatchMode", 1) ; Match title prefix
Opt("WinDetectHiddenText", 1)

Local $sTitle = "Microsoft Publisher Security Notice"

While 1
    ; Wait for the specific Publisher Security Notice window
    If WinExists($sTitle) Then
        ; Target and activate the window
        WinActivate($sTitle)
        WinWaitActive($sTitle, "", 2)

        ; Send Alt+E (Triggers the underlined 'E' on the Enable button)
        Send("!e")

        ; Brief pause to prevent duplicate triggers on the same window
        Sleep(1000)
    EndIf

    ; Poll every 250ms to keep CPU usage negligible
    Sleep(250)
WEnd