#include <Header.au3>

Func _UpdateWS()
	Local Const $autoUpdateINI = FileOpen("C:\Users\Paul\Desktop\AutoIT Stuff\cmwautoupdate.ini")
	Local $settings = FileReadToArray($autoUpdateINI)
	Local $cmwType = StringSplit($settings[0], "=", 2)
	Local $upgradePrefix

	Switch $cmwType[1]
		Case "qa"
			$upgradePrefix = "qaupgrade"
		Case Else
			MsgBox(0, "Error", "Invalid CMW Type (check cmwautoupdate.ini)")
	EndSwitch

	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	If @OSArch = "X64" Then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf

	WinWait("Confirm", "Update", 5)
	If WinExists("Log in", "OK") Then
		MsgBox(0, "No Update Required", "This workstation does not need to be updated.")
		Exit
	EndIf

	Local $updateText = StringSplit(WinGetText("Confirm"), "[]", 02)
	Local $versionArray = StringSplit($updateText[1], ".", 2)
	Local $executable = "cmw_" & $upgradePrefix & "_" & $versionArray[0] & "_" & $versionArray[1] & "_" & $versionArray[2] & ".exe"
	Local $dlURL = "http://car-partbidmate.com/cmw/" & $executable

	ConsoleWrite($executable & @CRLF)

	;FileDelete("C:\Users\Paul\Downloads\CMW*.exe")
	FileDelete("C:\cmw_*.exe")

	WinActivate("Confirm")
	ControlClick("Confirm", "Version update", "Button2")

	#comments-start
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

		Switch $browserDefault
		Case "explorer"
		WinActivate("View Downloads - Internet Explorer")
		Send("{Enter}")
		While FileExists("C:\Users\Paul\Downloads\" & $executable & ".*.partial")
		Sleep(1000)
		WEnd
		Case "mozilla"
		WinActivate("Opening")
		Sleep(250)
		Send("{Left}{Enter}")
		Sleep(1000)
		While FileExists("C:\Users\Paul\Downloads\" & $executable & ".part")
		Sleep(1000)
		WEnd
		Case "chrome"
		WinActivate("Google Chrome")
		Send("^j{Tab 3}{Enter}")
		While FileExists("C:\Users\Paul\Downloads\Unconfirmed*.crdownload")
		Sleep(1000)
		WEnd
		Case Else
		MsgBox(0, "Uh Oh", "Something went wrong with the automatic update.")
		EndSwitch

		;You will have to change the file name of the downloaded update, ie, 79_0_0 or 80_0_0
		;$begin = TimerInit()

		;	While TimerDiff($begin) < 300000
		;		 If FileExists("C:\Users\Paul\Downloads\" & $executable) Then
		;			 ConsoleWrite("Exited Loop" & @CRLF)
		;			 ExitLoop
		;		 EndIf
		;		 Sleep(1000)
		;	WEnd



		If Not FileExists("C:\Users\Paul\Downloads\" & $executable) Then
		MsgBox(0, "Aborting", "Couldn't find downloaded file.")
		Exit
		EndIf
	#comments-end


	InetGet($dlURL, "C:\" & $executable, 0, 1)
	MsgBox(0, "Please wait", "CMW is downloading in the background, and will install when done.")
	;Run("C:\Users\Paul\Downloads\" & $executable)
	$begin = TimerInit()

	While TimerDiff($begin) < 300000
		If FileExists("C:\" & $executable) Then
			ConsoleWrite("Exited Loop" & @CRLF)
			ExitLoop
		EndIf
		Sleep(1000)
	WEnd

	If Not FileExists("C:\Users\Paul\Downloads\" & $executable) Then
		MsgBox(0, "File not Found", "Download did not complete or took longer than five minutes. Please retry.")
		Exit
	EndIf

	Run("C:\" & $executable)

	WinWait("cmwupdater Setup", "Update is finished", 10)
	If WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup", "Update is finished")
		ControlClick("cmwupdater Setup", "OK", "Button1")
		MsgBox($MB_DEFBUTTON2, "Workstation Updated", "          Workstation Updated Successfully          ", 5)
	EndIf
		;MsgBox($MB_DEFBUTTON2, "Workstation Updated", "No Update Available")
;### Tidy Error -> "endfunc" is closing previous "while" on line 115
	EndFunc   ;==>_UpdateWS


	Func _Update()
		;CLEAR OUT your download directory before executing and check this path
		;FileDelete("C:\Users\Work-LT\Downloads\CMW*.exe")
		;Depending on 64 vs 32 you might have to change workstation path
		If @OSArch = "X64" Then
			Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
		Else
			Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
		EndIf
		_UpdateWS()
	EndFunc   ;==>_Update

	_Update()
	_OpenWS($defaultLoginFile)





;### Tidy Error -> func is never closed in your script.
