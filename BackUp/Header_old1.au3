#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GuiButton.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include <FileConstants.au3>

Local $logFileName = "C:\AutoITScript.log"
Local $totalTimer = TimerInit()
Local $POD = "XTJ"
Local $main = "[CLASS:TfrmMain_CMW]"
Local $cuswind = "[CLASS:TfrmCustomerUpdate]"
Local $find = "[CLASS:TfrmFind_Customer]"
Local $extra = "[CLASS:Tfrm_FindandSell_ExtraSales]"

Local $yardNumber
Local $userName
Local $userPwd

#region --- Internal functions Au3Recorder Start ---
Func _Au3RecordSetup()
	Opt('WinWaitDelay',1000)
	Opt('WinDetectHiddenText',1)
	Opt('MouseCoordMode',0)
	Opt('SendKeyDelay',100)
	Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
	If $aResult[1] <> '00000409' Then
		MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(00000409->' & $aResult[1] & ')')
	EndIf
	if FileExists($logFileName) Then
		FileDelete($logFileName)
	EndIf
EndFunc
Func _logIt($logInput)
	_FileWriteLog($logFileName,$logInput,-1)
EndFunc
Func _endProg()
	sleep(5000)
	Send("!f")
	Send("x")
	_logIt("Ending CMW Program.")
	_logIt("End of Execution - Total Seconds: " & Round(TimerDiff($totalTimer)/1000,2))
	;MsgBox($MB_OK,"SCRIPT EXECUTION NOTIFICATION","SCRIPT HAS ENDED.")
	Exit
EndFunc
Func _WinWait($classInput, $textInput, $logMess)
	Local $timer = TimerInit()
	if $textInput = "" Then
		local $wwRetunCode = WinWait($classInput,"",300)
		if $wwRetunCode = 0 Then
			_logIt($logMess & " - Screen Not Found or Timed-out. Ending Script")
			_endProg()
		Else
			Local $timerDiff = TimerDiff($timer)/1000
			_logIt($logMess & " - Seconds: " & Round($timerDiff,2))
		EndIf
	Else
		Local $wwRetunCode = WinWait($classInput, $textInput,300)
		if $wwRetunCode = 0 Then
			_logIt($logMess & " - Screen Not Found or Timed-out. Ending Script")
			_endProg()
		Else
			Local $timerDiff = TimerDiff($timer)/1000
			_logIt($logMess & " - Seconds: " & Round($timerDiff,2))
		EndIf
	EndIf
	Sleep(1500)
	$wwRetunCode=0
	$timerDiff=0
	$timer=0
	$classInput=""
	$textInput=""
	$logMess=""
EndFunc
Func _sendInput($siClass,$siControl,$siText)
	_logIt("Sending: " & $siText)
	ControlFocus($siClass,"",$siControl)
	Send($siText)
EndFunc
Func _OpenWS($fileName)
	Local Const $fileOpen = FileOpen($fileName)

	Local $loginline = FileReadLine($fileOpen)
	ConsoleWrite($loginline & @CRLF)
	Local $logininfo = StringSplit($loginline, ",")
	ConsoleWrite(FileReadLine($fileOpen) & @CRLF)

	$yardNumber = $logininfo[1]
	$userName = $logininfo[2]
	$userPwd = $logininfo[3]

	_AU3RecordSetup()
	_FileWriteLog($logFileName,"Beginning of Log File",1)
	;Check to see if CMW.exe is running, if so kill it.
	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	;ControlClick("[CLASS:#32770]", "Edit1",2)
	if @OSArch = "X64" then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf
	;Sleep(5000)
	if WinExists("C:\Program Files\Car-Part\Checkm~1\CMW.exe") Then
		WinActivate("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
		;ControlClick("C:\Program Files\Car-Part\Checkm~1\CMW.exe","",2)
		Send($POD&"{ENTER}")
	EndIf
	if WinExists("Confirm") Then
		_logit("CMW needs updating! Please run Update script!")
		_endProg()
	EndIF
	if WinExists("Select Yard Number") Then
		WinActivate("Select Yard Number")
		ControlClick("[CLASS:TfrmSelectYardNumber]","Edit1",2)
		Send($yardNumber&"{Enter}")
	EndIf
	WinWait("[CLASS:TfrmPassword]","")
	if WinExists("[CLASS:TfrmPassword]","OK") Then
		Local $pwdwind = "[CLASS:TfrmPassword]"
		ControlFocus($pwdwind,"","Edit1")
		;Send($username&"{TAB}"&$userPwd&"{Enter}")
		ControlSetText($pwdwind, "", "Edit1", $userName)
		ControlSetText($pwdwind, "", "TEdit1", $userPwd)
		ControlClick($pwdwind, "", "TBitBtn2")
	Else
		_logIt("CMW did NOT start or not seeing OK button - Ending Script Execution")
		_endProg()
	EndIf
	WinActivate("[CLASS:TfrmMain_CMW]")
	WinWait("[CLASS:TfrmMain_CMW]","eBay")
	ControlClick ("[CLASS:TfrmMain_CMW]","","TAdvGlowButton2")
	WinWait("[CLASS:TfrmMain_CMW]","frm_FindandSell_Main")
	if WinExists("[CLASS:TfrmMain_CMW]","frm_FindandSell_Main") Then
		_logIt("Found Find-N-Sell")
	Else
		_logIt("Did not find Find-N-Sell - Stopping Execution!")
		_endProg()
	EndIf
EndFunc
