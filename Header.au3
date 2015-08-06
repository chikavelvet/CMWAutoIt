#comments-start
	All script by Trey Gonsoulin unless otherwise noted.

	--- NOTES ---
	This has a bunch of global variables and functions that most scripts will
	probably need to function, and therefore should probably be included in them.
	Some of this may seem extraneous; that's because some of it is. Some things
	may be taken out of this and placed elsewhere in the future, but until it
	becomes significantly more bloated, it should be fine to leave as-is.

	References for Windows:
	- $g_wMain 		: The main CMW window; everything is inside of it
	- $g_wCustomer 	: The window for creating/editing a customer account
	- $g_wFind		: The window for finding a customer via search
	- $g_wExtra		: The window for extra sales
	- $g_wPay		: The window for payment, when entering a partial payment or deposit
	- $g_wGadget	: The window for selecting gadgets in Dashboard
	- $g_wMessage	: Any message window that CMW creates (e.g., do you want to close this tab?)
	- $g_wPassword	: The initial log-in window
	- $g_wTerminal	: An AlphaCom terminal window
	- $g_wDate		: The date change window (specifically from Dashboard)
	- $g_wDashSett	: The Dashboard settings window
	- $g_wPrint		: A print preview window

	These can be used as the "title" parameter for many AutoIt functions, to
	select a specific window. If they need modification, it can be done all at once.
	Currently they only hold a CLASS definition, but could also include title text
	and Regular Expressions, if necessary.

	Oh I got a comment now.s

#comments-end

#include-once

#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GuiButton.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include <FileConstants.au3>

Opt("WinDetectHiddenText", 1); 0: don't detect, 1: detect

;Password of the Day
;TO-DO: Make a better way to get/update this (it isn't used)
Global $g_sPOD; = "BTA"

;quick ways to reference the various windows, listed in NOTES above
Global 	$g_wMain = "[CLASS:TfrmMain_CMW]", 					$g_wCustomer = "[CLASS:TfrmCustomerUpdate]", _
		$g_wFind = "[CLASS:TfrmFind_Customer]", 			$g_wExtra = "[CLASS:Tfrm_FindandSell_ExtraSales]", _
		$g_wPay = "[CLASS:TfrmPaymentInputBox]", 			$g_wGadget = "[CLASS:TfrmDashboardSettings]", _
		$g_wMessage = "[CLASS:TAdvMessageForm]", 			$g_wPassword = "[CLASS:TfrmPassword]", _
		$g_wTerminal = "AlphaCom", 							$g_wDate = "[CLASS:TfrmEnterDateRange]", _
		$g_wDashSett = "[CLASS:TfrmDashboardSettings]", 	$g_wPrint = "[CLASS:TPrintpreview]"

;this is the default file with log-in information
;TO-DO: Maybe make an AutoIt GUI set-up to create this file and input all relevant information
Global $g_fiDefaultLoginFile = @AppDataDir & "\AutoIt\CMWLogin.csv"

;defined here for global status
;TO-DO: encapsulating these would be better programming practice
Global $g_iYardNumber
Global $g_sUserName, $g_sUserPwd

Global $g_iDefScreenCaptures = 0
Global $g_iErrorCount = 0

Global $g_asNonDefaultLogin

Global $g_fiDefaultConnectionFile = @ProgramFilesDir & "\OmniCom\AlphaCom\CHECKMATE.cft"

Global $g_tDate = @YEAR & @MON & @MDAY
Global $g_tTime = @HOUR & @MIN & @SEC
Global $g_fiScreenCaptureFolder = @AppDataDir & "\AutoIt\Screen Captures\" & $g_tDate & " " & $g_tTime

;this option makes title-matching not have to start at the beginning of the title
;makes some things easier, but could cause problems, so change if necessary
;(I don't believe anything done so far requires this)
AutoItSetOption("WinTitleMatchMode", 2)
HotKeySet("{END}", Terminate)

Func Terminate()
	Exit
EndFunc

HotKeySet("{Esc}", "_EndProg")

#comments-start

	-- _EndProg --
	This function is set up by OnAutoItExitRegister to run when AutoIt exits.
	A normal Exit command (@exitCode of 0) will simply cause a display message
	to come up stating the script has ended.
	Exit 2 (@exitCode of 2) will cause the same message box to appear, but also
	will close the CMW Main window.

#comments-end
Func _EndProg()
	If @exitCode = 2 Then; if the flag is set to 1
		WinClose($g_wMain); close the main CMW window
	EndIf
	MsgBox(0, "Script Execution Notification", "Script has ended."); notify that the script is ending
EndFunc   ;==>_EndProg

OnAutoItExitRegister("_EndProg")

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
	Local Const $autoUpdateINI = FileOpen(@AppDataDir & "\AutoIt\cmwautoupdate.ini")
	Local $fiSettings = FileReadToArray($autoUpdateINI)
	Local $asCMWType = StringSplit($fiSettings[0], "=", 2)
	Local $sUpgradePrefix

	ConsoleWrite("Here1" & @CRLF)
	;determining prefix to executable based on CMW type
	Switch $asCMWType[1]
		Case "qa"
			$sUpgradePrefix = "qaupgrade"
		Case Else
			MsgBox(0, "Error", "Invalid CMW Type (check cmwautoupdate.ini)")
			Exit 2
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

	Local $sExecutablePath = @UserProfileDir & "\Downloads\" & $sExecutable

	;delete any existing update executable with the same name to ensure no duplicates
	FileDelete($sExecutablePath)

	;decline and close the update window (download will be done by other means)
	WinActivate("Confirm")
	ControlClick("Confirm", "Version update", "Button2")

	;download the update file in the background, show message while doing it
	;TO-DO: find a way to automatically close the message box when dl completes,
	;instead of having to wait for the timeout anyway
	InetGet($dlURL, $sExecutablePath, 0, 1)
	MsgBox(0, "Please wait", "CMW is downloading in the background, and will install when done.", 5)

	;wait 5 minutes for the file to download
	Local $tBegin = TimerInit()
	While TimerDiff($tBegin) < 300000
		If FileExists($sExecutablePath) Then
			ExitLoop
		EndIf
		Sleep(1000)
	WEnd
	;ConsoleWrite("Found exec" & @CRLF & $sExecutablePath & @CRLF)

	;if is isn't done after 5 minutes, error
	If Not FileExists($sExecutablePath) Then
		MsgBox(0, "File not Found", "Download did not complete or took longer than five minutes. Please retry.")
		Exit 2
	EndIf

	;ConsoleWrite("About to run exec" & @CRLF)
	Run($sExecutablePath)
	;ConsoleWrite("Ran exec" & @CRLF)

	;wait for the upgrade to finish
	WinWait("cmwupdater Setup", "Update is finished")
	;display message that the automatic update completed
	If WinExists("cmwupdater Setup") Then
		WinActivate("cmwupdater Setup", "Update is finished")
		ControlClick("cmwupdater Setup", "OK", "Button1")
		MsgBox(0, "Workstation Updated", "          Workstation Updated Successfully          ", 5)
		Return 0
	EndIf
	Return -1
EndFunc   ;==>_UpdateWS

Func LogIn($fiLoginFile = $g_fiDefaultLoginFile)
	;wait for the first window (of yard select, update, and log-in) and get its handle
	Local $hFirstWindow = WinGetTitle(WinWaitActive("[REGEXPTITLE:(Select Yard Number)|(Confirm)|(Log in to Yard)|(Security Lock)]"))

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

	WinActivate($hFirstWindow)
	;wait for log-in screen to appear
	WinWait($g_wPassword, "")
	;if it finds it, log in, else display error message and halt script
	If WinExists($g_wPassword, "OK") Then
		ControlFocus($g_wPassword, "", "Edit1")
		ControlSetText($g_wPassword, "", "Edit1", $g_sUserName)
		ControlSetText($g_wPassword, "", "TEdit1", $g_sUserPwd)
		ControlClick($g_wPassword, "", "TBitBtn2")
	Else
		MsgBox(0, "Warning", "CMW did not start or not seeing OK button - Ending Script Execution")
		Exit 2
	EndIf
EndFunc

Func _ProgramFiles32Dir()
    Local $fiProgramFileDir
    Switch @OSArch
        Case "X32"
            $fiProgramFileDir = "Program Files"
        Case "X64"
            $fiProgramFileDir = "Program Files (x86)"
    EndSwitch
    Return @HomeDrive & "\" & $fiProgramFileDir
EndFunc   ;==>_ProgramFilesDirh

#comments-start

	-- _OpenWS --
	Automatically opens and signs in to CMW.
	Closes any existing CMW.exe process before it does so.
	Determines which ASM architecture to refer to based on OS
	- NOTE: x86-64 AutoIt scripts may not run completely correctly on x86 machines
	and vice-versa. May be best to compile two separate scripts for each task.
	Checks for three things on launch:
	- Update window
	- Yard selection window
	- Log-in window
	If update windows appears, it automatically updates via _UpdateWS
	If yard selection window appears, it enters in a yard number based on CMWLogin.csv and continues
	If log-in window appears, it logs in based on CMWLogin.csv
	After opening, it opens up apps in CMW based on CMWLogin.csv or parameters

	- $fiLoginFile: File location (string) -
	CSV (comma-delimited array) file to use for Logging In.
	Default is $g_fiDefaultLoginFile, the default location for the Log-in File

	- $asApps: Array of Strings -
	String-array containing the names of apps, in order, to open on start-up.
	The name codes are generally self-explanatory, and are listed in _OpenApp.
	Default is Null, which will open apps based on the second line of CMWLogic.cvs
	(the parameter will override apps listed in CMWLogin.csv)

#comments-end
Func _OpenWS($fiLoginFile = $g_fiDefaultLoginFile, $asApps = Null, $bKillOpen = True)
	;determine log-in information based on log-in file
	Local Const $fileOpen = FileOpen($fiLoginFile)
	Local $asOpenInfo = FileReadToArray($fileOpen)
	If $fiLoginFile <> $g_fiDefaultLoginFile Then
		$g_asNonDefaultLogin = $asOpenInfo
	EndIf
	Local $asLoginInfo = StringSplit($asOpenInfo[0], ",")
	Local $fiProgramFilesDir = _ProgramFiles32Dir()

;	ConsoleWrite("Here2" & @CRLF)

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
	$g_sPOD = $asLoginInfo[4]

	;Check to see if CMW.exe is running; if so, kill it.
	If $bKillOpen And ProcessExists("CMW.exe") Then ProcessClose("CMW.exe")

;	ConsoleWrite("About to run" & @CRLF)
;	ConsoleWrite($fiProgramFilesDir & "\Car-Part\Checkmate Workstation\CMW.exe")
	Run($fiProgramFilesDir & "\Car-Part\Checkmate Workstation\CMW.exe", "", @SW_MAXIMIZE)
;	ConsoleWrite("Ran" & @CRLF)
	LogIn($fiLoginFile)

	;wait for CMW main window to appear, and make sure it is active
	If $bKillOpen Then
		WinWait($g_wMain, "eBay")
		WinActivate($g_wMain)
	Else
		WinWait("[CLASS:TfrmMain_CMW; INSTANCE:2]", "eBay")
		WinActivate("[CLASS:TfrmMain_CMW; INSTANCE:2]")
	EndIf


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
	TO-DO:
	- Add the rest of the keywords
	- Possibly find better method than switch case that
	still allows easy calling (associative array via hashing?)

#comments-end
Func _OpenApp($sApp, $hWnd=$g_wMain)
	Local $sAppButton, $wAppWindow
	Local $bFlag = False

	Switch $sApp
		Case "pro"
			$sAppButton = "TAdvGlowButton2"
			$wAppWindow = "frm_FindandSell_Main"
		Case "dash"
			$sAppButton = "TAdvGlowButton11"
			$wAppWindow = "DBTAB"
			$bFlag = True
		Case "trak"
			$sAppButton = "TAdvGlowButton13"
			$wAppWindow = "OTTAB"
			$bFlag = True
		Case "term"
			$sAppButton = "TAdvGlowButton12"
			$wAppWindow = "TRMTAB1"
		Case "image"
			$sAppButton = "TAdvGlowButton10"
			$wAppWindow = "IMGTAB"
			$bFlag = True
		Case "report"
			$sAppButton = "TAdvGlowButton4"
			$wAppWindow = "REPORTTAB"
			$bFlag = True
		Case "ebay"
			$sAppButton = "TAdvGlowButton9"
			$wAppWindow = "EBTAB"
			$bFlag = True
		Case Else
			$sAppButton = Null
			$wAppWindow = Null
			;TO-DO: have error or something here
	EndSwitch

	If $bFlag Then
		If Not StringInStr(WinGetText($hWnd), $wAppWindow) Then
			$bFlag = False
		EndIf
	EndIf

	If $bFlag Then
		ConsoleWrite($sApp & " already opened." & @CRLF)
	Else
		WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 5)
		If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
			Return -1
		EndIf
		ControlClick($hWnd, "", $sAppButton)
		If $wAppWindow = "TRMTAB1" Then
			WinWait("Checkmate Workstation", "Terminal Path is not set", 2)
			If WinExists("Checkmate Workstation", "Terminal Path is not set") Then
				Exit
			EndIf
		EndIf
		WinWait($hWnd, $wAppWindow)
		Return 0
	EndIf
EndFunc   ;==>_OpenApp

#comments-start

	-- TermTextWait --
	Functions similarly to WinWait, but with terminal applications (specifically through AlphaCom).
	Rather than waiting for a window via the window's title and text, TermTextWait waits for a specific
	terminal screen through (250ms) periodic parsing of the screen's text and locating of key words.
	This is most likely a lot less efficient than WinWait, but the period is relatively long, and
	a single terminal screen is not an overwhelming amount of text, so it should be fine.

	- $hTitle: Handle -
	This is the title of the window in which the terminal resides. Generally this will be $g_wTerminal,
	for AlphaCom. The way the terminal embeds into CMW is kind of weird, so it's possible that $g_wMain
	may also work, or perhaps be the only working title.

	- $hText: String -
	This is any relevant text not in the terminal directly that can be used to certainly locate the
	terminal screen. In AlphaCom, the "Ready" at the bottom left is generally useful for this purpose.

	- $hControlID: String -
	This is the Control ID for the terminal screen. Note this is distinct from the window as a whole

	- $sTextToSearch: String -
	This is the text within the screen that is being waited on. An example is "Find and Sell" for
	waiting for a successful log in.

	- $tTimeout: time -
	This is an optional parameter denoting the amount of time to wait before giving up on the text
	ever being found. Default value is 0, which indicates no time out (it will loop indefinitely
	until finding the text). Note that the timeout time is in seconds, for ease of use, but can
	easily and naturally be changed (switched back) to milliseconds, if need be ($tTimeout can
	be decimal or fractional values, so this shouldn't be a huge issue.)


#comments-end
Func TermTextWait($hTitle, $hText, $hControlID, $sTextToSearch, $tTimeout = 0)

	ControlFocus($hTitle, $hText, $hControlID)

	Local $tBegin = TimerInit()
	Local $bWorked = False

	While $bWorked = False And ($tTimeout = 0 Or TimerDiff($tBegin) < ($tTimeout * 1000))
		Sleep(1000)
		Send("!es")
		If StringInStr(ClipGet(), $sTextToSearch) Then
			;ConsoleWrite(ClipGet() & @CRLF)
			$bWorked = True
			Return 0
		EndIf
	WEnd

	MsgBox(0, "Error", "Could not find text before time out.")

EndFunc   ;==>TermTextWait

#comments-start

	-- CaptureScreen --
	This is a modification on the native function _ScreenCapture_CaptureWnd. It takes a screen shot of a
	window and saves it to a directory with a name. These things are all variable. Note that all parameters
	are optional, primarily for ease of use.

	- $hTitle: Handle -
	This is the title of the window to be captured. Default value is the CMW main window. Note that
	if the main window is the window to capture, the capture function used is of the whole screen,
	not just the window (this function assumes the CMW main window to be generally maximized; if it
	isn't, it will still work, just potentially provide more information than necessary.)
	TO-DO: Rather than basing it on if it is the main window or not, base full-screen vs window
	capturing on whether the window is maximized or not.

	- $sName: String -
	This is a string used to give a custome name to the screen shot. This name does not include a
	leading backslash, nor does it include the file extension (which is always .png). The file extension
	default can be changed. If no name is specified, the default naming system is "screenshot" followed
	by an enumeration. Note that this suffixed number only changes when using the default naming system.

	- $fiAltDir: File Directory -
	By default, a new directory is created in the AppData\AutoIt\Screen Captures folder with its name
	consisting of the date and time in yyyymmdd hhmmss format. If desired, an alternate location may
	be specified to store the saved screenshot.

#comments-end
Func CaptureScreen($hTitle = $g_wMain, $sName = Null, $fiSubDir = Null, $fiAltDir = Null)

	Sleep(500)

	If Not FileExists(@AppDataDir & "\AutoIt\Screen Captures") Then
		DirCreate(@AppDataDir & "\AutoIt\Screen Captures")
	EndIf

	If Not FileExists($g_fiScreenCaptureFolder) Then
		DirCreate($g_fiScreenCaptureFolder)
	EndIf

	If $fiSubDir <> Null Then
		If Not FileExists($g_fiScreenCaptureFolder & "\" & $fiSubDir) Then
			DirCreate($g_fiScreenCaptureFolder & "\" & $fiSubDir)
		EndIf
	EndIf

	Local $sScreenCaptureName

	If $sName = Null Then
		If $fiSubDir = Null Then
			If $fiAltDir = Null Then
				$sScreenCaptureName = $g_fiScreenCaptureFolder & "\screenshot" & $g_iDefScreenCaptures & ".png"
			Else
				$sScreenCaptureName = $fiAltDir & "\screenshot" & $g_iDefScreenCaptures & ".png"
			EndIf
			$g_iDefScreenCaptures += 1
		Else
			$sScreenCaptureName = $g_fiScreenCaptureFolder & "\" & $fiSubDir & "\screenshot" & $g_iDefScreenCaptures & ".png"
			$g_iDefScreenCaptures += 1
		EndIf
	Else
		If $fiSubDir = Null Then
			If $fiAltDir = Null Then
				$sScreenCaptureName = $g_fiScreenCaptureFolder & "\" & $sName & ".png"
			Else
				$sScreenCaptureName = $fiAltDir & "\" & $sName & ".png"
			EndIf
		Else
			$sScreenCaptureName = $g_fiScreenCaptureFolder & "\" & $fiSubDir & "\" & $sName & ".png"
		EndIf
	EndIf

	ConsoleWrite($sScreenCaptureName & @CRLF)

	If $hTitle <> $g_wMain Then
		_ScreenCapture_CaptureWnd($sScreenCaptureName, $hTitle, 0, 0, -1, -1, False)
	Else
		_ScreenCapture_Capture($sScreenCaptureName, 0, 0, -1, -1, False)
	EndIf


EndFunc   ;==>CaptureScreen

#comments-start

	-- CheckError --
	This is a work-in-progress function that checks for the CMW error message
	window, and if it finds it, it immediately exits. At some point, this or
	a similar function could be implemented to immediately end the script and
	display/log/send a message saying that something went wrong. This would
	be especially helpful in the CMW Test script, as it would know if something
	it did broke Checkmate Workstation.

	TO-DO: make this work as intended

#comments-end
Func CheckError()
	If WinExists("[CLASS:madExceptWndClass]") Then
		MsgBox(0, "Error Encountered", "Ending Script")
		Exit
	EndIf
EndFunc   ;==>CheckError


