#comments-start

	--- NOTES ---
	This has a bunch of global variables and functions that most scripts will
	probably need to function, and therefore should probably be included in them.
	Some of this may seem extraneous; that's because some of it is. Some things
	may be taken out of this and placed elsewhere in the future, but until it
	becomes significantly more bloated, it should be fine to leave as-is.

#comments-end

#include-once

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
Global $g_wMain = "[CLASS:TfrmMain_CMW]",; the main CMW window (everything else is inside here)
$g_wCustomer = "[CLASS:TfrmCustomerUpdate]",; the window for customer creation/updating
$g_wFind = "[CLASS:TfrmFind_Customer]",; the window for searching for a customer
$g_wExtra = "[CLASS:Tfrm_FindandSell_ExtraSales]",; the window for extra sales
$g_wPay = "[CLASS:TfrmPaymentInputBox]",; the window for partial payments and deposits
$g_wGadget = "[CLASS:TfrmDashboardSettings]",; the window for selecting gadgets in dashboard
$g_wMessage = "[CLASS:TAdvMessageForm]",; CMW message windows
$g_wPassword = "[CLASS:TfrmPassword; TITLE:Log in to Yard]"; the initial log-in window

;this is the default file with log-in information
;TO-DO: Maybe make an AutoIT GUI set-up to create this file and input all relevant information
Global $g_fiDefaultLoginFile = "C:\Users\Paul\Desktop\AutoIT Stuff\CMWLogin.csv"

;defined here for global status
;TO-DO: encapsulating these would be better programming practice
Global $g_iYardNumber
Global $g_sUserName, $g_sUserPwd

;this option makes title-matching not have to start at the beginning of the title
;makes some things easier, but could cause problems, so change if necessary
;(I don't believe anything done so far requires this)
AutoItSetOption("WinTitleMatchMode", 2)

#comments-start

	-- _EndProg --
	Extension of the native Exit command.
	Pauses a few seconds, then displays a message saying the script has ended and exits.

	- $bFlag: 0 or 1 -
	0 : don't close CMW after script finish (default)
	1 : close CMW after script finish

#comments-end
Func _EndProg($bFlag = 0)
	Sleep(5000); pause a few seconds
	If $bFlag = 1 Then; if the flag is set to 1
		WinClose($g_wMain); close the main CMW window
	EndIf
	MsgBox(0, "Script Execution Notification", "Script has ended."); notify that the script is ending
	Exit; native exit
EndFunc   ;==>_EndProg

#comments-start

	-- _UpdateWS --
	Function to Update CMW. This is called (if applicable) by the _OpenWS() function.
	An .ini file called "cmwautoupdate.ini" has a setting "cmwtype".
	This type determines what upgrade executable to download/install.
	Currently the only type supported is QA, but that's easy to change.
	Uses a built-in function to download the update executable in the background
	directly from the car-partbidmate.com website.
	Waits 5 minutes for the file to download before throwing an error (modifiable)
	Continues to open CMW after completion
	Returns 0 if successful, -1 (or halts) if something's wrong
	TO-DO: make above statement actually true, and maybe change the way it handles things being wrong

#comments-end
Func _UpdateWS()
	;determing CMW type (QA, general, etc)
	Local Const $autoUpdateINI = FileOpen("C:\Users\Paul\Desktop\AutoIT Stuff\cmwautoupdate.ini")
	Local $fiSettings = FileReadToArray($autoUpdateINI)
	Local $asCMWType = StringSplit($fiSettings[0], "=", 2)
	Local $sUpgradePrefix

	;determining prefix to executable based on CMW type
	Switch $asCMWType[1]
		Case "qa"
			$sUpgradePrefix = "qaupgrade"
		Case Else
			MsgBox(0, "Error", "Invalid CMW Type (check cmwautoupdate.ini)")
			_EndProg(1)
	EndSwitch

	;wait for update window
	WinWait("Confirm", "Update", 5)

	;if it goes to log-in, then no update necessary
	If WinExists("Log in", "OK") Then
		MsgBox(0, "No Update Required", "This workstation does not need to be updated.")
		_OpenWS() ;TO-DO: this could cause an infinite loop, see what you can do
	EndIf

	;determine executable file name
	Local $asUpdateText = StringSplit(WinGetText("Confirm"), "[]", 02)
	Local $aiVersionArray = StringSplit($asUpdateText[1], ".", 2)
	Local $sExecutable = "cmw_" & $sUpgradePrefix & "_" & $aiVersionArray[0] & "_" & $aiVersionArray[1] & "_" & $aiVersionArray[2] & ".exe"

	;determine URL to download update
	Local $dlURL = "http://car-partbidmate.com/cmw/" & $sExecutable

	;write determined executable file name to console (testing purposes)
	ConsoleWrite($sExecutable & @CRLF)

	;delete all existing update executables to ensure no duplicates
	FileDelete("C:\cmw_*.exe")

	;decline and close the update window (download will be done by other means)
	WinActivate("Confirm")
	ControlClick("Confirm", "Version update", "Button2")

	;download the update file in the background, show message while doing it
	;TO-DO: find a way to automatically close the message box when dl completes,
	;instead of having to wait for the timeout anyway
	InetGet($dlURL, "C:\" & $sExecutable, 0, 1)
	MsgBox(0, "Please wait", "CMW is downloading in the background, and will install when done.", 5)

	;set-up for timer
	$begin = TimerInit()
	;for 5 minutes, check every second if the executable exists (dl complete)
	;if it is, exit the loop
	While TimerDiff($begin) < 300000
		If FileExists("C:\" & $sExecutable) Then
			ExitLoop
		EndIf
		Sleep(1000)
	WEnd

	;if the executable isn't there after the 5 minute loop, throw error
	If Not FileExists("C:\" & $sExecutable) Then
		MsgBox(0, "File not Found", "Download did not complete or took longer than five minutes. Please retry.")
		_EndProg(1)
	EndIf

	;run the executable
	Run("C:\" & $sExecutable)

	;wait for the upgrade to finish
	WinWait("cmwupdater Setup", "Update is finished, starting up Checkmate Workstation")
	;display message that the automatic update completed
	If WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup", "Update is finished")
		ControlClick("cmwupdater Setup", "OK", "Button1")
		MsgBox(0, "Workstation Updated", "          Workstation Updated Successfully          ", 5)
		Return 0
	EndIf
	Return -1
EndFunc   ;==>_UpdateWS

#comments-start

	-- _OpenWS --
	Automatically opens and signs in to CMW.
	Closes any existing CMW.exe process before it does so.
	Determines which ASM architecture to refer to based on OS
	 - NOTE: x86-64 AutoIT scripts may not run completely correctly on x86 machines
			 and vice-versa. May be best to compile two separate scripts for each task.
	Checks for three things on launch:
	 - Update window
	 - Yard selection window
	 - Log-in window
	If update windows appears, it automatically updates via _UpdateWS()
	If yard selection window appears, it enters in a yard number based on CMWLogin.csv and continues
	If log-in window appears, it logs in based on CMWLogin.csv
	After opening, it opens up apps in CMW based on CMWLogin.csv or parameters

	- $fiLoginFile: File location (string) -
	CSV (comma-delimited array) file to use for Logging In.
	Default is $g_fiDefaultLoginFile, the default location for the Log-in File

	- $asApps: Array of Strings -
	String-array containing the names of apps, in order, to open on start-up.
	The name codes are generally self-explanatory, and are listed in _OpenApp().
	Default is Null, which will open apps based on the second line of CMWLogic.cvs
	(the parameter will override apps listed in CMWLogin.csv)

#comments-end
Func _OpenWS($fiLoginFile = $g_fiDefaultLoginFile, $asApps = Null)
	;determine log-in information based on log-in file
	Local Const $fileOpen = FileOpen($fiLoginFile)
	Local $asOpenInfo = FileReadToArray($fileOpen)
	Local $asLoginInfo = StringSplit($asOpenInfo[0], ",")

	;if apps parameter is null (default), get it from log-in file, else use the parameter
	If $asApps = Null Then
		$asAppsToStart = StringSplit($asOpenInfo[1], ",")
	Else
		$asAppsToStart = $asApps
	EndIf

	;assign global variables for yard number, username, and password based on log-in file
	$g_iYardNumber = $asLoginInfo[1]
	$g_sUserName = $asLoginInfo[2]
	$g_sUserPwd = $asLoginInfo[3]

	;Check to see if CMW.exe is running; if so, kill it.
	If ProcessExists("CMW.exe") Then
		ProcessClose("CMW.exe")
	EndIf

	;determines CMW's location based on OS architecture
	;see NOTE above for warning on x86 and x86-64 AutoIT scripting
	If @OSArch = "X64" Then
		Run("C:\Program Files (x86)\Car-Part\Checkmate Workstation\CMW.exe")
	Else
		Run("C:\Program Files\Car-Part\Checkm~1\CMW.exe")
	EndIf

	;wait for the first window (of yard select, update, and log-in) and get its handle
	Local $hFirstWindow = WinGetTitle(WinWaitActive("[REGEXPTITLE:(Select Yard Number)|(Confirm)|(Log in to Yard)]"))

	;choose what to do based on what window appeared
	Switch $hFirstWindow
		Case "Select Yard Number"
			;ConsoleWrite("Select Yard Number Case" & @CRLF)
			WinActivate("Select Yard Number")
			ControlSetText("[CLASS:TfrmSelectYardNumber]", "", "Edit1", $g_iYardNumber)
			Send("{Enter}")
		Case "Confirm"
			;ConsoleWrite("Confirm Case" & @CRLF)
			_UpdateWS()
			_OpenWS()
	EndSwitch

	;wait for log-in screen to appear
	WinWait($g_wPassword, "")
	;if it finds it, log in, else display error message and halt script
	If WinExists($g_wPassword, "OK") Then
		ControlFocus($g_wPassword, "", "Edit1")\
		ControlSetText($g_wPassword, "", "Edit1", $g_sUserName)
		ControlSetText($g_wPassword, "", "TEdit1", $g_sUserPwd)
		ControlClick($g_wPassword, "", "TBitBtn2")
	Else
		MsgBox(0, "Warning", "CMW did not start or not seeing OK button - Ending Script Execution")
		_EndProg(1)
	EndIf

	;wait for CMW main window to appear, and make sure it is active
	WinWait($g_wMain, "eBay")
	WinActivate($g_wMain)

	;loop through and open apps from array created earlier
	For $i In $asAppsToStart
		_OpenApp($i)
	Next

EndFunc   ;==>_OpenWS

#comments-start

	-- _OpenApp --
	Opens a single CMW application based on string name passed in.

	- $sApp: String -
	String containing the keyword name of an app to open
	Keywords:
	 - "pro"  : Checkmate Sales Pro
	 - "dash" : Dashboard
	 - "trak" : OrderTrakker
	 - "term" : Alphacom Terminal
	TO-DO: 	Add the rest of the keywords
			Possibly find better method than switch case that
				still allows easy calling (associative array via hashing?)

#comments-end
Func _OpenApp($sApp)
	Local $sAppName, $wAppWindow

	Switch $sApp
		Case "pro"
			$sAppName = "TAdvGlowButton2", $wAppWindow = "frm_FindandSell_Main"
		Case "dash"
			$sAppName = "TAdvGlowButton11", $wAppWindow = "DBTAB"
		Case "trak"
			$sAppName = "TAdvGlowButton13", $wAppWindow = "OTTAB"
		Case "term"
			$sAppName = "TAdvGlowButton12", $wAppWindow = "TRMTAB1"
		Case Else
			$sAppName = Null, $wAppWindow = Null
			;TO-DO: have error or something here
	EndSwitch
	ControlClick("[CLASS:TfrmMain_CMW]", "", $sAppName)
	WinWait("[CLASS:TfrmMain_CMW]", $wAppWindow)
EndFunc   ;==>_OpenApp



