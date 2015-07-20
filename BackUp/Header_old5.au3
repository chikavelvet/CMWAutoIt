#comments-start

	--- NOTES ---
	This has a bunch of global variables and functions that most scripts will
	probably need to function, and therefore should probably be included in them.
	Some of this may seem extraneous; that's because some of it is. Some things
	may be taken out of this and placed elsewhere in the future, but until it
	becomes significantly more bloated, it should be fine to leave as-is.

#comments-end

#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GuiButton.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include <FileConstants.au3>

;Password of the Day
;TO-DO: Make a better way to get/update this (it isn't used)
Global $sPOD = "XTJ"

;quick ways to reference the various windows
Global 	$g_wMain = "[CLASS:TfrmMain_CMW]",; the main CMW window (everything else is inside here)
		$g_wCustomer = "[CLASS:TfrmCustomerUpdate]",; the window for customer creation/updating
		$g_wFind = "[CLASS:TfrmFind_Customer]",; the window for searching for a customer
		$g_wExtra = "[CLASS:Tfrm_FindandSell_ExtraSales]",; the window for extra sales
		$g_wPay = "[CLASS:TfrmPaymentInputBox]",; the window for partial payments and deposits
		$g_wGadget = "[CLASS:TfrmDashboardSettings]",; the window for selecting gadgets in dashboard
		$g_wMessage = "[CLASS:TAdvMessageForm]",; CMW message windows
		$g_wPassword = "[CLASS:TfrmPassword; TITLE:Log in to Yard]"; the initial log-in window

;this is the default file with log-in information
;TO-DO: Maybe make an AutoIT GUI set-up to create this file and input all relevant information
Global $sDefaultLoginFile = "C:\Users\Paul\Desktop\AutoIT Stuff\CMWLogin.csv"

;defined here for global status
;TO-DO: encapsulating these would be better programming practice
Global $iYardNumber
Global $sUserName
Global $sUserPwd

AutoItSetOption("WinTitleMatchMode", 2)

Func _endProg($bFlag = 0)
	Sleep(5000)
	If $bFlag = 1 Then
		WinClose($g_wMain)
	EndIf
	MsgBox(0, "Script Execution Notification", "Script has ended.")
	Exit
EndFunc   ;==>_endProg

Func _UpdateWS()
	Local Const $autoUpdateINI = FileOpen("C:\Users\Paul\Desktop\AutoIT Stuff\cmwautoupdate.ini")
	;Local Const $autoUpdateINI = FileOpen("C:\*\cmwautoupdate.ini")
	Local $fiSettings = FileReadToArray($autoUpdateINI)
	Local $asCMWType = StringSplit($fiSettings[0], "=", 2)
	Local $sUpgradePrefix

	Switch $asCMWType[1]
		Case "qa"
			$sUpgradePrefix = "qaupgrade"
		Case Else
			MsgBox(0, "Error", "Invalid CMW Type (check cmwautoupdate.ini)")
	EndSwitch

	WinWait("Confirm", "Update", 5)
	If WinExists("Log in", "OK") Then
		MsgBox(0, "No Update Required", "This workstation does not need to be updated.")
		_OpenWS() ;TO-DO: this could cause an infinite loop, see what you can do
	EndIf

	Local $asUpdateText = StringSplit(WinGetText("Confirm"), "[]", 02)
	Local $aiVersionArray = StringSplit($asUpdateText[1], ".", 2)
	Local $sExecutable = "cmw_" & $sUpgradePrefix & "_" & $aiVersionArray[0] & "_" & $aiVersionArray[1] & "_" & $aiVersionArray[2] & ".exe"
	Local $dlURL = "http://car-partbidmate.com/cmw/" & $sExecutable

	ConsoleWrite($sExecutable & @CRLF)

	FileDelete("C:\cmw_*.exe")

	WinActivate("Confirm")
	ControlClick("Confirm", "Version update", "Button2")


	InetGet($dlURL, "C:\" & $sExecutable, 0, 1)
	MsgBox(0, "Please wait", "CMW is downloading in the background, and will install when done.", 5)
	$begin = TimerInit()

	While TimerDiff($begin) < 300000
		If FileExists("C:\" & $sExecutable) Then
			ExitLoop
		EndIf
		Sleep(1000)
	WEnd

	If Not FileExists("C:\" & $sExecutable) Then
		MsgBox(0, "File not Found", "Download did not complete or took longer than five minutes. Please retry.")
		_endProg(1)
	EndIf

	Run("C:\" & $sExecutable)

	WinWait("cmwupdater Setup", "Update is finished, starting up Checkmate Workstation")
	If WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup", "Update is finished")
		ControlClick("cmwupdater Setup", "OK", "Button1")
		MsgBox(0, "Workstation Updated", "          Workstation Updated Successfully          ", 5)
	EndIf
EndFunc   ;==>_UpdateWS

Func _OpenWS($fiLoginFile = $sDefaultLoginFile, $asApps = Null)
	Local Const $fileOpen = FileOpen($fiLoginFile)
	Local $openInfo = FileReadToArray($fileOpen)
	Local $logininfo = StringSplit($openInfo[0], ",")

	If $asApps = Null Then
		$asAppsToStart = StringSplit($openInfo[1], ",")
	Else
		$asAppsToStart = $asApps
	EndIf

	$iYardNumber = $logininfo[1]
	$sUserName = $logininfo[2]
	$sUserPwd = $logininfo[3]

	;Check to see if CMW.exe is running, if so kill it.
	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	If @OSArch = "X64" Then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf

	Local $hFirstWindow = WinGetTitle(WinWaitActive("[REGEXPTITLE:(Select Yard Number)|(Confirm)|(Log in to Yard)]"))

	;ConsoleWrite($firstWindow & @CRLF)

	Switch $hFirstWindow
		Case "Select Yard Number"
			;ConsoleWrite("Select Yard Number Case" & @CRLF)
			WinActivate("Select Yard Number")
			ControlSetText("[CLASS:TfrmSelectYardNumber]", "", "Edit1", $iYardNumber)
			Send("{Enter}")
		Case "Confirm"
			;ConsoleWrite("Confirm Case" & @CRLF)
			_UpdateWS()
			_OpenWS()
	EndSwitch

	;ConsoleWrite("About to wait for the password screen" & @CRLF)
	WinWait($g_wPassword, "")
	;ConsoleWrite("Found the password screen" & @CRLF)
	If WinExists($g_wPassword, "OK") Then
		ControlFocus($g_wPassword, "", "Edit1")
		;Send($sUserName&"{TAB}"&$sUserPwd&"{Enter}")
		ControlSetText($g_wPassword, "", "Edit1", $sUserName)
		ControlSetText($g_wPassword, "", "TEdit1", $sUserPwd)
		ControlClick($g_wPassword, "", "TBitBtn2")
	Else
		MsgBox(0, "Warning", "CMW did not start or not seeing OK button - Ending Script Execution")
		_endProg(1)
	EndIf
	WinActivate($g_wMain)
	WinWait($g_wMain, "eBay")


	For $i In $asAppsToStart
		_OpenApp($i)
	Next

EndFunc   ;==>_OpenWS

Func _OpenApp($sApp)
	Local $sAppName
	Local $wAppWindow

	Switch $sApp
		Case "pro"
			$sAppName = "TAdvGlowButton2"
			$wAppWindow = "frm_FindandSell_Main"
		Case "dash"
			$sAppName = "TAdvGlowButton11"
			$wAppWindow = "DBTAB"
		Case "trak"
			$sAppName = "TAdvGlowButton13"
			$wAppWindow = "OTTAB"
		Case "term"
			$sAppName = "TAdvGlowButton12"
			$wAppWindow = "TRMTAB1"
		Case Else
			$sAppName = Null
			$wAppWindow = Null
			;TODO have error here
	EndSwitch
	ControlClick("[CLASS:TfrmMain_CMW]", "", $sAppName)
	WinWait("[CLASS:TfrmMain_CMW]", $wAppWindow)
EndFunc   ;==>_OpenApp



