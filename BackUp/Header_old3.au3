#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GuiButton.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include <FileConstants.au3>

Local $totalTimer = TimerInit()
Local $POD = "XTJ"
Local $main = "[CLASS:TfrmMain_CMW]"
Local $cuswind = "[CLASS:TfrmCustomerUpdate]"
Local $find = "[CLASS:TfrmFind_Customer]"
Local $extra = "[CLASS:Tfrm_FindandSell_ExtraSales]"
Local $pay = "[CLASS:TfrmPaymentInputBox]"
Local $gadget = "[CLASS:TfrmDashboardSettings]"
Local $message = "[CLASS:TAdvMessageForm]"
Local $pwdwind = "[CLASS:TfrmPassword; TITLE:Log in to Yard]"
Local $defaultLoginFile = "C:\Users\Paul\Desktop\AutoIT Stuff\Login.csv"

Local $yardNumber
Local $userName
Local $userPwd

AutoItSetOption("WinTitleMatchMode", 2)

Func _endProg()
	Sleep(5000)
	;Send("!f")
	;Send("x")
	WinClose($main)
	MsgBox($MB_OK,"SCRIPT EXECUTION NOTIFICATION","SCRIPT HAS ENDED.")
	Exit
EndFunc   ;==>_endProg
Func _WinWait($classInput, $textInput, $logMess)
	Local $timer = TimerInit()
	If $textInput = "" Then
		Local $wwRetunCode = WinWait($classInput, "", 300)
		If $wwRetunCode = 0 Then
			_logIt($logMess & " - Screen Not Found or Timed-out. Ending Script")
			_endProg()
		Else
			Local $timerDiff = TimerDiff($timer) / 1000
			_logIt($logMess & " - Seconds: " & Round($timerDiff, 2))
		EndIf
	Else
		Local $wwRetunCode = WinWait($classInput, $textInput, 300)
		If $wwRetunCode = 0 Then
			_logIt($logMess & " - Screen Not Found or Timed-out. Ending Script")
			_endProg()
		Else
			Local $timerDiff = TimerDiff($timer) / 1000
			_logIt($logMess & " - Seconds: " & Round($timerDiff, 2))
		EndIf
	EndIf
	Sleep(1500)
	$wwRetunCode = 0
	$timerDiff = 0
	$timer = 0
	$classInput = ""
	$textInput = ""
	$logMess = ""
EndFunc   ;==>_WinWait

Func _sendInput($siClass, $siControl, $siText)
	_logIt("Sending: " & $siText)
	ControlFocus($siClass, "", $siControl)
	Send($siText)
EndFunc   ;==>_sendInput

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

	FileDelete("C:\cmw_*.exe")

	WinActivate("Confirm")
	ControlClick("Confirm", "Version update", "Button2")


	InetGet($dlURL, "C:\" & $executable, 0, 1)
	MsgBox(0, "Please wait", "CMW is downloading in the background, and will install when done.", 5)
	$begin = TimerInit()

	While TimerDiff($begin) < 300000
		If FileExists("C:\" & $executable) Then
			ExitLoop
		EndIf
		Sleep(1000)
	WEnd

	If Not FileExists("C:\" & $executable) Then
		MsgBox(0, "File not Found", "Download did not complete or took longer than five minutes. Please retry.")
		Exit
	EndIf

	Run("C:\" & $executable)

	WinWait("cmwupdater Setup", "Update is finished, starting up Checkmate Workstation", 5)
	If WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup", "Update is finished")
		ControlClick("cmwupdater Setup", "OK", "Button1")
		MsgBox($MB_DEFBUTTON2, "Workstation Updated", "          Workstation Updated Successfully          ", 5)
	EndIf
	;MsgBox($MB_DEFBUTTON2, "Workstation Updated", "No Update Available")
EndFunc

Func _OpenWS($loginFile = $defaultLoginFile, $apps = Null)
	Local Const $fileOpen = FileOpen($loginFile)
	Local $openInfo = FileReadToArray($fileOpen)
	Local $logininfo = StringSplit($openInfo[0], ",")
	Local $appsToStart

	If $apps = Null Then
		$appsToStart = StringSplit($openInfo[1], ",")
		$appst = $appsToStart
	Else
		$appst = $apps
	EndIf

	$yardNumber = $logininfo[1]
	$userName = $logininfo[2]
	$userPwd = $logininfo[3]

	;Check to see if CMW.exe is running, if so kill it.
	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	If @OSArch = "X64" Then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf

	Local $firstWindow = WinGetTitle(WinWaitActive("[REGEXPTITLE:(Select Yard Number)|(Confirm)|(Log in to Yard)]"))

	;ConsoleWrite($firstWindow & @CRLF)

	Switch $firstWindow
		Case "Select Yard Number"
			;ConsoleWrite("Select Yard Number Case" & @CRLF)
			WinActivate("Select Yard Number")
			ControlSetText("[CLASS:TfrmSelectYardNumber]", "", "Edit1", $yardNumber)
			Send("{Enter}")
		Case "Confirm"
			;ConsoleWrite("Confirm Case" & @CRLF)
			_UpdateWS()
			_OpenWS()
			_endProg()
	EndSwitch

	;ConsoleWrite("About to wait for the password screen" & @CRLF)
	WinWait($pwdwind, "")
	;ConsoleWrite("Found the password screen" & @CRLF)
	If WinExists($pwdwind, "OK") Then
		ControlFocus($pwdwind, "", "Edit1")
		;Send($username&"{TAB}"&$userPwd&"{Enter}")
		ControlSetText($pwdwind, "", "Edit1", $userName)
		ControlSetText($pwdwind, "", "TEdit1", $userPwd)
		ControlClick($pwdwind, "", "TBitBtn2")
	Else
		MsgBox(0, "Warning", "CMW did not start or not seeing OK button - Ending Script Execution")
		_endProg()
	EndIf
	WinActivate($main)
	WinWait($main, "eBay")


	For $i In $appst
		_OpenApp($i)
	Next

EndFunc   ;==>_OpenWS

Func _OpenApp($app)
	Local $appName
	Local $appWindow

	Switch $app
		Case "pro"
			$appName = "TAdvGlowButton2"
			$appWindow = "frm_FindandSell_Main"
		Case "dash"
			$appName = "TAdvGlowButton11"
			$appWindow = "DBTAB"
		Case "trak"
			$appName = "TAdvGlowButton13"
			$appWindow = "OTTAB"
		Case "term"
			$appName = "TAdvGlowButton12"
			$appWindow = "TRMTAB1"
		Case Else
			$appName = Null
			$appWindow = Null
			;TODO have error here
	EndSwitch
	ControlClick("[CLASS:TfrmMain_CMW]", "", $appName)
	WinWait("[CLASS:TfrmMain_CMW]", $appWindow)
EndFunc   ;==>_OpenApp



