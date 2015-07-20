#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---
#include <Header.au3>

Func _UpdateWS()

	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	If @OSArch = "X64" Then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf

	WinWait("Confirm","Update",5)
	Local $updateText = StringSplit(WinGetText("Confirm"), "[]", 02)
	Local $versionArray = StringSplit($updateText[1], ".", 2)
	;_ArrayDisplay($versionArray)
	ControlClick("Confirm","&Yes","Button1")
	Sleep(2500)
	Local $browserTitle = WinGetTitle("[REGEXPTITLE:(Opening)|(Google Chrome)|(Internet Explorer)]")
	Local $browserDefault
	If StringInStr($browserTitle, "Opening") Then
		$browserDefault = "mozilla"
	ElseIf StringInStr($browserTitle, "Google") Then
		$browserDefault = "chrome"
	ElseIf StringInStr($browserTitle, "Explorer") Then
		$browserDefault = "explorer"
	EndIf
	If $browserDefault <> "explorer" Then
		Run ("cmd.exe")
		WinWait("[CLASS:ConsoleWindowClass]")
	EndIf
	Switch $browserDefault
		Case "explorer"
			WinActivate("View Downloads - Internet Explorer")
			Send("{Enter}")
		Case "mozilla"
			WinActivate("Opening")
			Sleep(100)
			Send("{Left}{Enter}")
		;Case "chrome"
			;WinActivate("Google Chrome")
			;Send("^j{Tab 3}{Enter}")
		Case Else
			MsgBox(0, "Uh Oh", "Something went wrong with the automatic update.")
	EndSwitch

	;You will have to change the file name of the downloaded update, ie, 79_0_0 or 80_0_0
	Sleep(20000) ;time for the download, will need to change later
	If $browserDefault <> "explorer" Then
		WinActivate("[CLASS:ConsoleWindowClass]")
		Send("cd..{Enter}cd..{Enter}cd Downloads{Enter}")
		ConsoleWrite("cmw_qaupgrade_" & $versionArray[0] & "_" & $versionArray[1] & "_" & $versionArray[2] & ".exe" & @CRLF)
	EndIf

	Exit
	WinWait("cmwupdater Setup","",10)
	if WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup")
		ControlClick("cmwupdater Setup","OK","Button1")
		MsgBox($MB_DEFBUTTON2,"Workstation Updated","          Workstation Updated Successful          ")

	EndIf
	MsgBox($MB_DEFBUTTON2,"Workstation Updated", "No Update Available")
EndFunc


Func _Update()
;CLEAR OUT your download directory before executing and check this path
;FileDelete("C:\Users\Work-LT\Downloads\CMW*.exe")
;Depending on 64 vs 32 you might have to change workstation path
if @OSArch = "X64" then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
EndIf
_UpdateWS()
EndFunc

_Update()
_OpenWS($defaultLoginFile)




