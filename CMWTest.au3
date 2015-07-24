#comments-start
	All script by Trey Gonsoulin unless otherwise noted.

	--- NOTES ---
	This is the main run file for the Checkmate Workstation test script.
	It goes through the list of tasks provided in the CMW Test Plan
	and completes the actions required to perform the testing. Any test
	that requires some sort of verification of correctness (i.e., most
	of them) ultimately result in screen captures being taken and saved
	to their unique folders. Each test has its own folder. This file
	includes a single function for each test, as well as a few other
	functions used in the tests. By editing the bottom section, or
	including this file in another script, you can run each test
	individually. The test functions return a 'complete' message if they
	complete successfully.

	Reference (CMW Test Plan):
	http://wiki01.car-part.com:8090/display/CMDEV/CMWTestPlan

#comments-end

#include <Header.au3>
;#include <DashboardTest.au3>
#include <Date.au3>
#include <IE.au3>


#comments-start

	-- MinusSeven --
	This function subtracts seven days from the date given. The date
	comes in an array that is MDY formatted. To do the date subtraction,
	which can be complicated around month and year beginnings, the
	date is first converted into its Julian Day Value, the seven days
	are subtracted, and then it is converted back into a YMD format.

	- $tDateArray: Date (Array, MDY format) -
	The input is a reference to the array, so note that it modifies the
	parameter directly; this is done for convenience (and slight memory
	efficiency) rather than necessity, so it can be changed if required.
	Also note that the resulting date array is in global date standard
	YMD format,	despite	the input being in MDY form (which comes unchanged
	from CMW). This is also easily modifiable.

#comments-end
Func MinusSeven(ByRef $tDateArray)
	Local $M = $tDateArray[0]
	Local $D = $tDateArray[1]
	Local $Y = $tDateArray[2]
	Local $sJulDate = _DateToDayValue($Y, $M, $D)
	$sJulDate = _DayValueToDate($sJulDate - 7, $Y, $M, $D)
	$tDateArray[0] = $Y
	$tDateArray[1] = $M
	$tDateArray[2] = $D
EndFunc   ;==>MinusSeven


Local $tDateToday = [@MON, @MDAY, @YEAR]
Local $tDateMinusSeven = $tDateToday
MinusSeven($tDateMinusSeven)
$tDateToday = $tDateToday[2] & "-" & $tDateToday[0] & "-" & $tDateToday[1]
$tDateMinusSeven = $tDateMinusSeven[0] & "-" & $tDateMinusSeven[1] & "-" & $tDateMinusSeven[2]
ConsoleWrite($tDateToday & @CRLF & $tDateMinusSeven & @CRLF)

#comments-start

	-- TerminalOpenLogin --
	This is a function of convenience that opens up the terminal
	application and signs into it using the CMWTest.csv login
	information. It finishes when it finds the Find and Sell screen,
	and if it doesn't, it errors and exits the script.

	Note: this may not work on all terminals, including any Checkmate
	Junior yard (or any terminal that does not begin at the Find and
	Sell screen.)
	TO-DO: confirm/modify to work on all terminal log ins (if necessary)

#comments-end
Func TerminalOpenLogin($hWnd = "[CLASS:AfxFrameOrView80]")
	_OpenApp("term")

	ControlFocus($g_wTerminal, "Ready", $hWnd)
	Send("!es")
	While Not StringInStr(ClipGet(), "[2J")
		Send("{Enter}")
		Sleep(250)
		Send("!es")
	WEnd

	Send($g_sUserPwd & "{Enter}")

	If TermTextWait($g_wTerminal, "Ready", $hWnd, "Find and Sell") = -1 Then
		MsgBox(0, "No Find and Sell Found", "Something went wrong.")
		Exit
	EndIf
EndFunc   ;==>TerminalOpenLogin

#Region --- SETTINGS TEST FUNCTION ---

Func AccessCMWToolbar($i, $j, $k = 0)
	Local $iXOff
	Switch $i
		Case 0
			$iXOff = 30
		Case 1
			$iXOff = 100
		Case 2
			$iXOff = 170
		Case Else
			$iXOff = 0
	EndSwitch

	ControlClick($g_wMain, "", "TAdvToolBar1", "primary", 1, $iXOff, 20)

	Send("{Down " & $j & "}")
	If $i = 1 And $j = 1 Then
		Send("{Right}{Down " & $k & "}")
	EndIf

	Send("{Enter}")
EndFunc   ;==>AccessCMWToolbar

Func TestSettings()
	#Region -- Settings Test 1 - Edit security settings so only yard owner has access to all dashboard gadgets
	While Not WinExists("Security for Dashboard")
		AccessCMWToolbar(1, 1, 0)
		WinWait("Security for Dashboard", "", 5)
	WEnd

	Send("{Down 3}{Up}1")
	For $i = 0 To 24
		Send("{Down}1")
	Next
	ControlClick("Security for Dashboard", "", "TBitBtn1", "primary")
	;Exit
	#EndRegion -- Settings Test 1 - Edit security settings so only yard owner has access to all dashboard gadgets


	#Region -- Settings Test 2 - Set settings to eBay so only YO can list parts, but anyone with sales can view tab

	While Not WinExists("Security for eBay")
		AccessCMWToolbar(1, 1, 1)
		WinWait("Security for eBay", "", 5)
	WEnd

	Send("{Down 3}{Up}1")
	Send("{Down}2,3")
	ControlClick("Security for eBay", "", "TBitBtn1", "primary")

	#EndRegion -- Settings Test 2 - Set settings to eBay so only YO can list parts, but anyone with sales can view tab

	#Region -- Settings Test 3 - Give imaging rights so that only inventory can add images

	While Not WinExists("Security for CMIS")
		AccessCMWToolbar(1, 1, 2)
		WinWait("Security for CMIS", "", 5)
	WEnd

	ControlClick("Security for CMIS", "", "TAdvStringGrid1", "primary", 1, 180, 30)
	Send("5,6")
	ControlClick("Security for CMIS", "", "TBitBtn1", "primary")

	#EndRegion -- Settings Test 3 - Give imaging rights so that only inventory can add images

	#Region -- Settings Test 4 - Login as a user with only inventory rights, verify imaging works
	_OpenWS(@AppDataDir & "\AutoIt\CMWTestInvLogin.csv")
	ImageTest123("SettingsTest")

	#EndRegion -- Settings Test 4 - Login as a user with only inventory rights, verify imaging works

	#Region -- Settings Test 5 - Try to open dashboard, get security error
	_OpenApp("dash")
	WaitForUpdating()
	CaptureScreen($g_wMain, "NotYODashboard", "SettingsTest")
	#EndRegion -- Settings Test 5 - Try to open dashboard, get security error

	#Region -- Settings Test 6 - Try to open eBay, get security error
	ControlClick($g_wMain, "", "TAdvGlowButton9", "primary")
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "You do not have security to", 5)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]", "You do not have security to") Then
		CaptureScreen($g_wMain, "NotYOeBay", "SettingsTest")
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "You do not have security to", "Button1", "primary")
	Else
		MsgBox(0, "Security box not appearing correctly", "Security settings may not be working properly.")
	EndIf
	#EndRegion -- Settings Test 6 - Try to open eBay, get security error

	#Region -- Settings Test 7 - Login as yard owner, should be able to see all gadgets, and access ebay
	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")
	_OpenApp("dash")
	WaitForUpdating()
	CaptureScreen($g_wMain, "YODashboard", "SettingsTest")

	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	_OpenApp("ebay")
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 10)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
		CaptureScreen($g_wMain, "ErrorImage" & $g_iErrorCount)
	EndIf
	CaptureScreen($g_wMain, "YOeBay", "SettingsTest")

	Sleep(10000)

	ControlClick($g_wMain, "", "TAdvGlowButton25")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	Sleep(1000)
	#EndRegion -- Settings Test 7 - Login as yard owner, should be able to see all gadgets, and access ebay

	#Region -- Settings Test 8 - (as YO), Should not be able to add images to images
	ControlClick($g_wMain, "", "TAdvGlowButton10", "primary")
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 15)
	CaptureScreen($g_wMain, "YOImaging", "SettingsTest")
	ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
	#EndRegion -- Settings Test 8 - (as YO), Should not be able to add images to images

	#Region -- Settings Test 9 - Change lockout setting to 3 minutes, let PC sit and verify ws prompt appears
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	Send("^{Tab 5}")
	ControlSetText("Setup", "", "TAdvSpinEdit1", "1")
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)
	;WinWaitNotActive($g_wMain, "", 350)
	;ConsoleWrite("Not active" & @CRLF)
	;WinActivate($g_wPassword)
	;WinSetState($g_wPassword, "", @SW_SHOW)
	;ConsoleWrite(Binary(WinGetState($g_wMain)) & @CRLF)
	;ConsoleWrite("Title: " & WinGetTitle(WinGetHandle("")) & @CRLF & "Text: " & WinGetText(WinGetHandle("")) & @CRLF)
	;ConsoleWrite("Waited" & @CRLF)
	;WinActivate($g_wMain)
	;LogIn(@AppDataDir & "\AutoIt\CMWTest.csv")

	#EndRegion -- Settings Test 9 - Change lockout setting to 3 minutes, let PC sit and verify ws prompt appears

	#Region -- Settings Test 10 - Set different tools to open automatically, verify they do so
	While Not WinExists("Security for CMIS")
		AccessCMWToolbar(1, 1, 2)
		WinWait("Security for CMIS", "", 5)
	WEnd

	ControlClick("Security for CMIS", "", "TAdvStringGrid1", "primary", 1, 180, 30)
	Send("1,5,6")
	ControlClick("Security for CMIS", "", "TBitBtn1", "primary")

	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd

	Send("^{Tab 4}")
	For $i = 0 To 6
		Send("{Tab}")
		If Random(0, 1, 1) = 1 Then
			Send("{Space}")
		EndIf
	Next
	CaptureScreen($g_wMain, "RandomTabsToOpen1", "SettingsTest")
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")

	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 30)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
		CaptureScreen($g_wMain, "ErrorImage" & $g_iErrorCount)
	EndIf
	WinActivate($g_wMain)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	CaptureScreen($g_wMain, "RandomTabsOpened1", "SettingsTest")

	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	WinActivate("Setup")
	Send("^{Tab 4}")
	For $i = 0 To 6
		Sleep(5000)
		Send("{Tab}")
		Send("{Space}")
	Next
	CaptureScreen($g_wMain, "RandomTabsToOpen2", "SettingsTest")
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")

	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 30)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
		CaptureScreen($g_wMain, "ErrorImage" & $g_iErrorCount)
	EndIf
	While Not WinExists("Setup")
		WinActivate($g_wMain)
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	CaptureScreen($g_wMain, "RandomTabsOpened2", "SettingsTest")

	#EndRegion -- Settings Test 10 - Set different tools to open automatically, verify they do so

	#Region -- Settings Test 11 - Edit printer settings
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	WinActivate("Setup")
	Send("^{Tab 3}")
	;ControlClick("Setup", "", "Edit5", "primary")
	ControlSend("Setup", "", "Edit5", "{Down}")
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)
	#EndRegion -- Settings Test 11 - Edit printer settings

	Exit
EndFunc   ;==>TestSettings

#EndRegion --- SETTINGS TEST FUNCTION ---

#Region --- DASHBOARD TEST FUNCTION ---


#comments-start

	-- WaitForNewGadget --
	This function is based somewhat off the WinWait function, except
	it waits for a Dashboard gadget instead. Additionally, it waits
	for a new gadget to appear, by taking in the initial current status
	bar	message (unique for each gadget), and periodically comparing it
	to the new status bar message. When it first detects a change
	(usually, though not always, a change to "Updating"), it runs the
	WaitForUpdating function. After it knows the status has changed,
	and is also not "Updating", it updates the initial current status,
	to be used when it's called next.

	- $sCurStat: String -
	This is the initial current status. A refence directly to it is
	passed in rather than its value so that it may be updated when the
	function ends. All checks are done against this status.

	- $wWindow: Window -
	This is the window that the function checks for the status. This is
	always the main CMW window during these tests, but is included as
	and optional parameter for modularity's sake.

	- $idStatBar: Control ID -
	This is the ID of the status bar used to get the new status to check
	against the initial one. This defaults to the only status bar used
	during testing, TStatusBar1.

	Note: This function needs more work to be nice, but it works as-is.
	One problem is that ControlGetText only returns the first gadget's
	status when used on TStatusBar1, so it will not work if the first
	gadget doesn't change, even if others do. (this is not, however,
	vital to the testing process).

#comments-end
;TO-DO: make this work for multiple gadgets
Func WaitForNewGadget(ByRef $sCurStat, $wWindow = $g_wMain, $idStatBar = "TStatusBar1")
	Local $sNewStatus = $sCurStat
	While $sNewStatus = $sCurStat
		$sNewStatus = ControlGetText($wWindow, "", $idStatBar)
		Sleep(250)
	WEnd
	WaitForUpdating()
	$sCurStat = ControlGetText($wWindow, "", $idStatBar)
EndFunc   ;==>WaitForNewGadget

Func WaitForUpdating($tExtraWait = .1, $wWindow = $g_wMain)
	While StringInStr(WinGetText($wWindow), "Updating")
		Sleep(250)
	WEnd
	Sleep($tExtraWait * 1000)
EndFunc   ;==>WaitForUpdating

Func TestDashboard()
	WinSetState($g_wMain, "", @SW_MAXIMIZE)
	_OpenApp("dash")
	Local $sCurrentStatus

	ControlFocus($g_wMain, "Dashboard", "TAdvToolBar2")

	#Region -- Dashboard Test 1 - Turn on all gadgets -done
	Sleep(1000)
	$sCurrentStatus = ControlGetText($g_wMain, "", "TStatusBar1")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)
	ControlClick($g_wDate, "", "TButton1")
	WaitForUpdating()
	Sleep(200)
	CaptureScreen($g_wMain, "GadgetTest1", "DashboardTest")
	For $i = 0 To 24
		ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
		WinWait($g_wGadget)
		;Sleep(150)

		For $j = 1 To $i
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		Next

		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")

		;remove if hidden gadget bug is fixed
		If $i = 15 Or $i = 17 Then
			$i += 1
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		EndIf

		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		Sleep(1000)
		ControlClick($g_wGadget, "", "TBitBtn2")

		WaitForNewGadget($sCurrentStatus)
		Sleep(500)
		CaptureScreen($g_wMain, "GadgetTest" & ($j + 1), "DashboardTest")
		Sleep(500)
	Next
	;Exit
	#EndRegion -- Dashboard Test 1 - Turn on all gadgets -done

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 25}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")

	#Region -- Dashboard Test 2 - Set date for 7 days -done
	;TO-DO: Skip gadgets that don't use date range
	WaitForNewGadget($sCurrentStatus)
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)

	ControlClick($g_wDate, "Date Range", "TButton1", "primary")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)

	ControlCommand($g_wDate, "Date Range", "TCheckBox1", "Check", "")

	Local $tDateArrayMDY = StringSplit(ControlGetText($g_wDate, "Date Range", "TDateTimePicker2") & @CRLF, "/", 2)
	MinusSeven($tDateArrayMDY)
	ControlClick($g_wDate, "Date Range", "TDateTimePicker1", "primary")
	ControlSend($g_wDate, "Date Range", "TDateTimePicker1", "{Right 2}" & $tDateArrayMDY[2])
	;Exit
	ControlClick($g_wDate, "", "TBitBtn1", "primary")

	WaitForUpdating()
	Sleep(100)

	CaptureScreen($g_wMain, "Gadget7DaysTest", "DashboardTest")

	For $i = 0 To 24
		ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
		WinWait($g_wGadget)
		;Sleep(1000)

		For $j = 1 To $i
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		Next

		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")

		;remove if hidden gadget bug is fixed
		If $i = 15 Or $i = 17 Then
			$i += 1
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		EndIf

		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		Sleep(1000)
		ControlClick($g_wGadget, "", "TBitBtn2")

		WaitForNewGadget($sCurrentStatus)
		Sleep(250)
		CaptureScreen($g_wMain, "Gadget7DaysTest" & ($j + 1), "DashboardTest")
		Sleep(250)
	Next

	#EndRegion -- Dashboard Test 2 - Set date for 7 days -done

	;Exit
	Sleep(250)
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	Sleep(200)
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 25}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")
	Sleep(150)
	WaitForUpdating()

	#Region -- Dashboard Test 4 - Display Only WO for "NameLoggedIn" -done

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")
	WaitForUpdating(.5)

	CaptureScreen($g_wMain, "WOs_ALL", "DashboardTest")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WaitForUpdating(.5)
	CaptureScreen($g_wMain, "WOs_(" & $g_sUserName & ")", "DashboardTest")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WaitForUpdating(.5)

	#EndRegion -- Dashboard Test 4 - Display Only WO for "NameLoggedIn" -done

	#Region -- Dashboard Test 5 - Expand and retract gadgets -done
	;TO-DO: make sure all gadget opens wait for the window afterward
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)

	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")

	ControlClick($g_wGadget, "", "TBitBtn2")
	WaitForNewGadget($sCurrentStatus)
	WaitForUpdating(1)

	CaptureScreen($g_wMain, "ExpandGadget(Before)", "DashboardTest")
	ControlClick($g_wMain, "", "TButton1")
	Sleep(250)
	CaptureScreen($g_wMain, "ExpandGadget(After)1", "DashboardTest")
	ControlClick($g_wMain, "", "TButton1")
	Sleep(250)
	ControlClick($g_wMain, "", "TButton2")
	Sleep(250)
	CaptureScreen($g_wMain, "ExpandGadget(After)2", "DashboardTest")
	ControlClick($g_wMain, "", "TButton1")
	Sleep(250)
	;Exit
	#EndRegion -- Dashboard Test 5 - Expand and retract gadgets -done

	#Region -- Dashboard Test 6 - Change columns to display -done
	WaitForUpdating(.5)
	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
	Sleep(100)
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
	WinWait($g_wDashSett)
	For $n = 0 To 2
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wDashSett, "", "TBitBtn2", "primary")
	WaitForUpdating(1)
	CaptureScreen($g_wMain, "DisplayColumnsChanged", "DashboardTest")
	Sleep(100)
	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
	Sleep(100)
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
	WinWait($g_wDashSett)
	For $n = 0 To 2
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wDashSett, "", "TBitBtn2", "primary")
	WaitForUpdating(1)


	#EndRegion -- Dashboard Test 6 - Change columns to display -done

	#Region -- Dashboard Test 7 - Show details for WO -done

	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
	Sleep(200)
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 7}")
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "Error: WO Number", 5)

	Local $i = 1
	While WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]", "Error: WO Number")
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "Error: WO Number", "Button1", "primary")
		ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
		Sleep(200)
		ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 5}")
		ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
		WinWait("Enter Date for Active WOs")
		ControlSend("Enter Date for Active WOs", "", "TDateTimePicker2", "{Right}" & @MDAY - $i)
		$i += 1
		ControlClick("Enter Date for Active WOs", "", "TBitBtn1", "primary")
		WaitForUpdating()
		ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
		Sleep(200)
		ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 7}")
		ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
		WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "Error: WO Number", 5)
	WEnd

	WinWait($g_wPrint)
	WinSetState($g_wPrint, "", @SW_MAXIMIZE)
	ControlClick($g_wPrint, "", "TButton4", "primary")
	Sleep(100)
	CaptureScreen($g_wPrint, "WODetails", "DashboardTest")
	ControlClick($g_wPrint, "", "TButton6", "primary")

	#EndRegion -- Dashboard Test 7 - Show details for WO -done

	#Region -- Dashboard Test 3 - Verify resizing -done

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	For $i = 0 To 12
		If Random(0, 1, 1) = 1 Then
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		EndIf
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next


	ControlClick($g_wGadget, "", "TBitBtn2")
	Sleep(250)
	WaitForUpdating()
	Sleep(100)
	CaptureScreen($g_wMain, "Resize", "DashboardTest")

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	For $i = 0 To 12
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wGadget, "", "TBitBtn2")

	WaitForUpdating()
	Sleep(100)

	CaptureScreen($g_wMain, "Resize2", "DashboardTest")

	#EndRegion -- Dashboard Test 3 - Verify resizing -done

	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	Return "Dashboard Test Complete"

EndFunc   ;==>TestDashboard

#EndRegion --- DASHBOARD TEST FUNCTION ---

#Region --- TERMINAL TEST FUNCTION ---

Func CheckTermSettings($fiConnectionFile = $g_fiDefaultConnectionFile)
	ControlSend($g_wMain, "", "TAdvOfficePager1", "!sw")
	WinWait("[CLASS:TfrmSetup_CMW; TITLE:Setup]")
	WinActivate("[CLASS:TfrmSetup_CMW; TITLE:Setup]")
	ControlSend("[CLASS:TfrmSetup_CMW; TITLE:Setup]", "", "TPageControl1", "^{Tab 2}")
	Local $fiCurConnectionFile = ControlGetText("[CLASS:TfrmSetup_CMW; TITLE:Setup]", "", "TLabeledEdit1")
	If $fiCurConnectionFile <> $fiConnectionFile Then
		MsgBox(0, "Changing Connection File", "Current connection file: " & $fiCurConnectionFile & @CRLF & "New connection file: " & $fiConnectionFile & @CRLF, 3)
		ControlSetText("[CLASS:TfrmSetup_CMW; TITLE:Setup]", "", "TLabeledEdit1", $fiConnectionFile)
	EndIf
	WinClose("[CLASS:TfrmSetup_CMW; TITLE:Setup]")
EndFunc   ;==>CheckTermSettings

Func TestTerminal()
	WinSetState($g_wMain, "", @SW_MAXIMIZE)
	#Region -- Terminal Test 1 - set up terminal settings and verify alphacom opens -done
	CheckTermSettings($g_asNonDefaultLogin[4])
	TerminalOpenLogin()
	CaptureScreen($g_wMain, "AlphaComOpened", "TerminalTest")
	#EndRegion -- Terminal Test 1 - set up terminal settings and verify alphacom opens -done

	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
	EndIf

	#Region -- Terminal Test 2 - verify an open alphacom screen outside of CMW will be focused when opening terminal app in CMW -done
	Run(@ProgramFilesDir & "\OmniCom\AlphaCom\Alpha.exe")
	WinWait("AlphaCom")
	WinActivate($g_wMain)
	Sleep(5000)
	ControlClick($g_wMain, "", "TAdvGlowButton12")
	CaptureScreen("AlphaCom", "AlphaComFocused", "TerminalTest")
	WinClose("AlphaCom")
	#EndRegion -- Terminal Test 2 - verify an open alphacom screen outside of CMW will be focused when opening terminal app in CMW -done

	#Region -- Terminal Test 3 - verify that alphacom opens in both of 2 separate CMWs -done

	WinActivate($g_wMain)
	TerminalOpenLogin()
	Local $hNdl = ControlGetHandle($g_wTerminal, "", "[CLASS:AfxFrameOrView80]")
	WinMove($g_wMain, "", 0, 0, @DesktopWidth / 2, @DesktopHeight)
	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv", Null, False)
	WinMove("[CLASS:TfrmMain_CMW; INSTANCE:2]", "", @DesktopWidth / 2, 0, @DesktopWidth / 2, @DesktopHeight)
	WinActivate("[CLASS:TfrmMain_CMW; INSTANCE:2]")
	TerminalOpenLogin()
	CaptureScreen($g_wMain, "TwoCMWAlphaComs", "TerminalTest")
	#EndRegion -- Terminal Test 3 - verify that alphacom opens in both of 2 separate CMWs -done

	WinClose($g_wMain)
	WinSetState($g_wMain, "", @SW_MAXIMIZE)
	;ConsoleWrite("Found the Find and Sell; No Problemo 2" & @CRLF)
	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
	EndIf

	Return "Terminal Test Complete"
EndFunc   ;==>TestTerminal

#EndRegion --- TERMINAL TEST FUNCTION ---

#Region --- ORDER TRAKKER TEST FUNCTION ---

#Region -- Order Trakker Test Vars and Funcs --
Global $aiOTTabPos[15] = [30]
Global $aiOTTabNum[15]

Global $asOTNameIndex[15] = [ _
		"Dispatch", _
		"Warehouse", "Dismantling", "Yard", "Brokered", _
		"Arrived", "Void", "RdySnd", _
		"CPU", "Truck", "LTL", "FedEx/UPS", _
		"Returned", "Delivered", "Restocked" _
		]

Global $aiOTTabBaseWidths[15] = [80, 94, 100, 59, 78, 66, 59, 68, 60, 60, 60, 93, 79, 83, 88]
Global $aiOTTabCurrentWidths[15] = [80, 94, 100, 59, 78, 66, 59, 68, 60, 60, 60, 93, 79, 83, 88]
Global $sLastActiveTab = "Restocked"

Func OTGetActiveTab()
	Local $sOTText = WinGetText($g_wMain)
	Local $sOTTextStripped = StringStripCR(StringStripWS($sOTText, 8))
	;	ConsoleWrite("Last Active Tab: " & $sLastActiveTab & @CRLF & "Text: " & $sOTTextStripped & @CRLF)
	Local $asOTTextMatch
	If $sLastActiveTab == "CPU" Or $sLastActiveTab == "TCL" Then
		$asOTTextMatch = StringRegExp($sOTTextStripped, "Dispatch(?s)(.*)" & StringLeft($sLastActiveTab, 3), 1)
	Else
		$asOTTextMatch = StringRegExp($sOTTextStripped, "Dispatch(?s)(.*)" & StringLeft($sLastActiveTab, 4), 1)
	EndIf
	Local $sActiveTab = $asOTTextMatch[0]
	;	ConsoleWrite("GETTING ACTIVE TAB: " & $sActiveTab & @CRLF & @CRLF)
	Return $sActiveTab
EndFunc   ;==>OTGetActiveTab

Func OTGetActiveTabIndex()
	Local $sActiveTab = OTGetActiveTab()
	Local $asActiveInfoSplit = StringSplit($sActiveTab, "()")
	Local $iTabIndex = _ArraySearch($asOTNameIndex, $asActiveInfoSplit[1])
	Return $iTabIndex
EndFunc   ;==>OTGetActiveTabIndex

Func OTSetActiveTabNum()
	Local $sActiveTab = OTGetActiveTab()
	Local $asActiveInfoSplit = StringSplit($sActiveTab, "()")
	Local $iTabIndex = _ArraySearch($asOTNameIndex, $asActiveInfoSplit[1])
	;	ConsoleWrite("ActiveInfoSplit[0]: " & $asActiveInfoSplit[0] & @CRLF)
	;	ConsoleWrite("ActiveInfoSplit[1]: " & $asActiveInfoSplit[1] & @CRLF)
	If $asActiveInfoSplit[0] == 1 Then
		;		ConsoleWrite("ArraySearch: " & $iTabIndex & @CRLF)
		$aiOTTabNum[$iTabIndex] = 0
	Else
		Local $iTabNum = $asActiveInfoSplit[2]
		$aiOTTabNum[$iTabIndex] = $iTabNum
		;		Local $iOffset = 0
		;		If $aiOTTabNum[$iTabIndex] < 10 Then
		;			$iOffset = 28
		;		Else
		;			$iOffset = 36
		;		EndIf
		;		$aiOTTabCurrentWidths[$iTabIndex] = $aiOTTabBaseWidths[$iTabIndex] + $iOffset
	EndIf
EndFunc   ;==>OTSetActiveTabNum

Func OTSwitchToTab($iTargetTabIndex)
	;ConsoleWrite("Tab Index to Switch to: " & $iTabIndex & @CRLF)
	Local $iActiveTabIndex = OTGetActiveTabIndex()
	If $iActiveTabIndex <> $iTargetTabIndex Then
		Local $tempTab = OTGetActiveTab()
		;		ConsoleWrite("Target x-pos: " & $aiOTTabPos[$iTargetTabIndex] & @CRLF)
		ControlClick($g_wMain, "", "TPageControl1", "primary", 1, $aiOTTabPos[$iTargetTabIndex], 10)
		$sLastActiveTab = $tempTab
	EndIf
EndFunc   ;==>OTSwitchToTab

Func OTSetInitialTabPos()
	OTSetInitialNums()
	OTUpdatePos()
EndFunc   ;==>OTSetInitialTabPos

Func OTSetInitialNums()
	;	Local $iCurTabWidth
	For $i = 0 To UBound($aiOTTabPos) - 2
		OTSwitchToTab($i)
		OTSetActiveTabNum()
		OTUpdatePos()
		;$iCurTabWidth = $aiOTTabCurrentWidths[$i]
		;$aiOTTabPos[$i + 1] = $aiOTTabPos[$i] + $iCurTabWidth
	Next
EndFunc   ;==>OTSetInitialNums

Func OTUpdatePos()
	Local $iOffset
	For $i = 0 To UBound($aiOTTabPos) - 2
		$iOffset = 0
		If $aiOTTabNum[$i] <> 0 Then
			If $aiOTTabNum[$i] < 10 Then
				$iOffset = 28
			Else
				$iOffset = 36
			EndIf
		EndIf

		$aiOTTabPos[$i + 1] = $aiOTTabPos[$i] + $aiOTTabBaseWidths[$i] + $iOffset
		$aiOTTabCurrentWidths[$i] = $aiOTTabBaseWidths[$i] + $iOffset

		;		ConsoleWrite(@CRLF & "$i: " & $i & @CRLF & "$iOffset: " & $iOffset & @CRLF & "$aiOTTabNum[$i]: " _
		;			& $aiOTTabNum[$i] & @CRLF & "$aiOTTabPos[$i]: " & $aiOTTabPos[$i] & @CRLF & _
		;			"$aiOTTabPos[$i+1]: " & $aiOTTabPos[$i+1] & @CRLF & "$aiOTTabBaseWidths[$i]: "  &$aiOTTabBaseWidths[$i] & @CRLF & @CRLF)
	Next
EndFunc   ;==>OTUpdatePos

Func OTSendPartToTab($iTargetTabIndex)
	Local $iActiveTabIndex = OTGetActiveTabIndex()
	ControlClick($g_wMain, "", "TColorButton" & (UBound($aiOTTabPos) - 1) - $iTargetTabIndex, "primary")

	$aiOTTabNum[$iActiveTabIndex] -= 1
	$aiOTTabNum[$iTargetTabIndex] += 1
	OTUpdatePos()
EndFunc   ;==>OTSendPartToTab

#EndRegion -- Order Trakker Test Vars and Funcs --

Func TestTrakker()
	_OpenApp("trak")
	While Not WinExists("Setup")
		WinActivate($g_wMain)
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	Local $posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)


	OTSetInitialTabPos()

	_ArrayDisplay($aiOTTabCurrentWidths)

	;Sleep(500)

	;Assumes set up and parts sales are already completed
	;TO-DO: do this instead of assuming it's done

	;the next two tests may not be fully working yet

	#Region -- Order Trakker Test 1 - verify that when you sell each part they're put in the correct tab
	;Check Warehouse Tab
	OTSwitchToTab(1)
	CaptureScreen($g_wMain, "WarehouseTab1", "OrderTrakkerTest")
	OTSendPartToTab(0)

	;Check Brokered Tab
	OTSwitchToTab(4)
	CaptureScreen($g_wMain, "BrokeredTab1", "OrderTrakkerTest")
	OTSendPartToTab(0)

	;Check Yard Tab
	OTSwitchToTab(3)
	CaptureScreen($g_wMain, "YardTab1", "OrderTrakkerTest")
	OTSendPartToTab(0)
	Sleep(500)
	#EndRegion -- Order Trakker Test 1 - verify that when you sell each part they're put in the correct tab


	#Region -- Order Trakker Test 2 - Move each part into every tab
	OTSwitchToTab(0)
	CaptureScreen($g_wMain, "", "Part012Dispatch")
	For $i = 0 To 2
		OTSwitchToTab(0)
		Sleep(500)
		For $j = 1 To UBound($aiOTTabPos) - 2
			OTSendPartToTab($j)
			Sleep(500)
			OTSwitchToTab($j)
			CaptureScreen($g_wMain, "Part" & $i & "Tab" & $j)
		Next
	Next
	#EndRegion -- Order Trakker Test 2 - Move each part into every tab

	Exit
	;TO-DO: figure out how to make Test 1 and 2 work

	#Region -- Order Trakker Test 3 - Verify that the history is correct

	ControlClick($g_wMain, "", "TPageControl1", "primary", 1, $aiOTTabPos[0], 10)
	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary", 1, 20, 40)
	Send("{Down 7}{Enter}")
	CaptureScreen($g_wMain, "PartHistory", "OrderTrakkerTest")

	#EndRegion -- Order Trakker Test 3 - Verify that the history is correct
	Exit
	Return "Order Trakker Test Complete"
EndFunc   ;==>TestTrakker



#EndRegion --- ORDER TRAKKER TEST FUNCTION ---

#Region --- IMAGING TEST FUNCTION ---

Func ImageTest123($sTestSubDir = "ImagingTest")
	#Region -- Imaging Test 1 - Look up a stock number -done
	Local $sLoginFileImagingLine = $g_asNonDefaultLogin[2]
	Local $asImagingLineArray = StringSplit($sLoginFileImagingLine, ",", 2)
	Global $iStockNumber = $asImagingLineArray[0]
	Global $sPartCode = $asImagingLineArray[1]
	;ConsoleWrite($iStockNumber & @CRLF)
	;ConsoleWrite("1" & @CRLF)
	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 5)
	ControlSetText($g_wMain, "Imaging", "TAdvEdit4", $iStockNumber)
	ControlClick($g_wMain, "Imaging", "TAdvGlowButton15", "primary")
	WinWait($g_wMain, $iStockNumber)

	;ConsoleWrite("3" & @CRLF)

	CaptureScreen($g_wMain, "StockNumber" & $iStockNumber & "Lookup", $sTestSubDir)

	#EndRegion -- Imaging Test 1 - Look up a stock number -done

	#Region -- Imaging Test 2 - Look up a specific part -done

	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 5)
	ControlSetText($g_wMain, "Imaging", "TAdvEdit2", $sPartCode)
	ControlClick($g_wMain, "Imaging", "TAdvGlowButton15", "primary")
	WinWait($g_wMain, $sPartCode)

	CaptureScreen($g_wMain, "SpecificPart" & $sPartCode, $sTestSubDir)

	#EndRegion -- Imaging Test 2 - Look up a specific part -done

	#Region -- Imaging Test 3 - Add single image using browse image search window -done

	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 150)
	ControlClick($g_wMain, "Imaging", "TDirectoryListBoxEx1", "primary", 2, 1, 1)
	Local $asPicture = StringSplit($g_fiScreenCaptureFolder, "\")
	;_ArrayDisplay($asPicture)
	;Opt("SendKeyDelay", 100)
	;	For $i = 2 To $asPicture[0]
	;		Sleep(100)
	;		ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Down}")
	;		Send(StringLower(StringLeft($asPicture[$i], 1)))
	;		ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Enter}")
	;		Sleep(150)
	;	Next
	$j = 0
	For $i In $asPicture
		If $j >= 2 Then
			Sleep(500)
			ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Down 2}")
			Send($i)
			ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Enter}")
			;Sleep(5000))
		EndIf
		$j += 1
	Next
	ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Down 2}")
	Send($sTestSubDir)
	ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Enter}")
	WinWait($g_wMain, "c: [OS]", 2)

	ControlFocus($g_wMain, "Imaging", "TAdvStringGrid1")
	ControlSend($g_wMain, "Imaging", "TAdvStringGrid1", "{Space}")
	Global $aiImageFromPos = ControlGetPos($g_wMain, "Imaging", "TImageEnMView2")
	Global $aiImageToPos = ControlGetPos($g_wMain, "Imaging", "TImageEnMView1")
	MouseClickDrag("primary", $aiImageFromPos[0] + 175, $aiImageFromPos[1] + 50, $aiImageToPos[0] + 50, $aiImageToPos[1] + 50)
	WinWait("Confirm", "Filename being copied", 5)
	While WinExists("Confirm", "Filename being copied")
		ControlClick("Confirm", "Filename being copied", "Button1", "primary")
		WinWait("Confirm", "Filename being copied", 2.5)
	WEnd
	Sleep(2500)
	CaptureScreen($g_wMain, "AddSinglePartFromBrowse", $sTestSubDir)

	#EndRegion -- Imaging Test 3 - Add single image using browse image search window -done
EndFunc   ;==>ImageTest123


Func TestImaging()
	WinSetState($g_wMain, "", @SW_SHOWNORMAL)
	WinSetState($g_wMain, "", @SW_MAXIMIZE)
	_OpenApp("image")

	ImageTest123()

	#Region -- Imaging Test 4 - Add multiple images using Windows Explorer
	Opt("SendKeyDelay", 5)
	ControlFocus($g_wMain, "Imaging", "TAdvStringGrid1")
	ControlSend($g_wMain, "Imaging", "TAdvStringGrid1", "{Space}")

	Send("{LWIN}")
	Sleep(1000)
	Send($g_fiScreenCaptureFolder & "\ImagingTest\{Enter}")
	WinWait("ImagingTest")
	WinActivate("ImagingTest")
	WinMove("ImagingTest", "", 0, 0, @DesktopWidth / 2, @DesktopHeight - 100)
	Local $idExplorerControl = ControlGetHandle("ImagingTest", "", "DirectUIHWND3")
	ControlSend("ImagingTest", "", $idExplorerControl, "!vr{Enter}")
	Sleep(250)
	ControlSend("ImagingTest", "", $idExplorerControl, "{Right}+{Left}")
	$aiImageFromPos = ControlGetPos("ImagingTest", "", $idExplorerControl)
	MouseClickDrag("primary", $aiImageFromPos[0] + 100, $aiImageFromPos[1] + 50, @DesktopWidth / 2 + 100, $aiImageToPos[1] + 50)


	WinWait("Confirm", "Filename being copied", 2.5)
	While WinExists("Confirm", "Filename being copied")
		ControlClick("Confirm", "Filename being copied", "Button1", "primary")
		WinWait("Confirm", "Filename being copied", 2.5)
	WEnd
	WinActivate($g_wMain)

	;TO-DO: fix/change this
	;WinWait("[CLASS:TAdvSmoothMessageForm]", "", 5)
	;WinWaitClose("[CLASS:TAdvSmoothMessageForm]", "", 5)
	Sleep(7500)

	CaptureScreen($g_wMain, "AddMultiplePartsFromExplorer", "ImagingTest")

	#EndRegion -- Imaging Test 4 - Add multiple images using Windows Explorer

	#Region -- Imaging Test 5 - Delete images -done

	ControlClick($g_wMain, "Imaging", "TImageEnMView1", "secondary", 1, 50, 50)
	ControlSend($g_wMain, "Imaging", "TImageEnMView1", "r")
	;ControlSend($g_wMain, "Imaging", "TImageEnMView1", "m")
	;WinWait("[CLASS:TMessageForm; TITLE:Confirm]")
	;ControlSend($g_wMain, "Imaging", "TImageEnMView1", "y")

	Sleep(5000)

	CaptureScreen($g_wMain, "DeleteImage", "ImagingTest")

	#EndRegion -- Imaging Test 5 - Delete images -done

	#Region -- Imaging Test 6 - Verify import images works

	CaptureScreen($g_wMain, Null, "ImagingTest")
	Send("!fi")
	ConsoleWrite($g_sPOD & @CRLF)
	WinWait("Password of the Day")
	ControlSetText("Password of the Day", "", "TEdit1", $g_sPOD)
	ControlClick("Password of the Day", "", "TBitBtn2", "primary")

	WinWait("CMW - Image Import")
	Local $fiCurImageDestPath = ControlGetText("CMW - Image Import", "", "TAdvEdit1")
	ControlSetText("CMW - Image Import", "", "TAdvEdit2", $g_fiScreenCaptureFolder & "\ImagingTest")
	ControlClick("CMW - Image Import", "", "TAdvGlowButton3", "primary")
	Sleep(1000)
	ControlClick("CMW - Image Import", "", "TAdvGlowButton1", "primary")
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "Import Process Finished")
	ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "Import Process Finished", "Button1", "primary")
	WinClose("CMW - Image Import")
	WinWaitActive($g_wMain)
	Send("{LWin}")
	Sleep(100)
	Send($fiCurImageDestPath & "\" & @YEAR & "{Enter}")
	WinWait(@YEAR, "Address")
	CaptureScreen(@YEAR, "ImportedLocation", "ImagingTest")
	WinClose(@YEAR)

	#EndRegion -- Imaging Test 6 - Verify import images works
	#comments-start
		#Region -- Imaging Test 7 - Change location of imported images

		WinClose(@YEAR)
		WinActivate($g_wMain)

		Send("!fi")
		WinWait("Password of the Day")
		ControlSetText("Password of the Day", "", "TEdit1", $g_sPOD)
		ControlClick("Password of the Day", "", "TBitBtn2", "primary")
		WinWait("CMW - Image Import")
		ControlSend("CMW - Image Import", "", "TAdvOfficePager1", "^{Tab}")
		ControlClick("CMW - Image Import", "", "TAdvGlowButton2", "primary")
		WinWait("Confirm", "&No")
		ControlClick("Confirm", "", "TButton2", "primary")

		WinWait("Browse for Folder", "Select a folder")
		Local $hTreeView = ControlGetHandle("Browse for Folder", "Select a folder", "SysTreeView321")
		;Local $fiNewImageDestPath = $fiCurImageDestPath & "NewCMImages\"
		Local $fiNewImageDestPath = $fiCurImageDestPath
		Local $sReplaced = StringReplace($fiNewImageDestPath, "\", "|")
		DirCreate($fiNewImageDestPath & "\NewCMImages")
		ConsoleWrite($sReplaced & @CRLF)
		$sReplaced = StringLeft($sReplaced, StringLen($sReplaced) - 1)
		;Exit
		$sReplaced = "#0" & StringRight($sReplaced, StringLen($sReplaced) - 2)
		ConsoleWrite($sReplaced & @CRLF)
		;ControlTreeView("", "", $hTreeView, "Select", "Desktop|Computer|" & StringLeft($sReplaced,StringLen($sReplaced)-1))
		ControlTreeView("", "", $hTreeView, "Expand", "Desktop|Computer|#0")
		Sleep(750)
		;ConsoleWrite("Directory is: " & ControlTreeView("", "", $hTreeView, "GetText", "#0|#4|#0|#6") & @CRLF)
		ConsoleWrite("Desktop|Computer|" & $sReplaced & "|NewCMImages" & @CRLF)
		ControlTreeView("", "", $hTreeView, "Expand", "Desktop|Computer|" & $sReplaced)
		Sleep(250)
		ControlTreeView("", "", $hTreeView, "Select", "Desktop|Computer|" & $sReplaced & "|NewCMImages")
		ControlClick("Browse for Folder", "", "Button1", "primary")

		WinWait("Confirm", "&No")
		ControlClick("Confirm", "", "TButton2", "primary")

		Exit

		#EndRegion -- Imaging Test 7 - Change location of imported images
	#comments-end
	#Region -- Imaging Test 8 - Look up the images in alphacom and verify link

	TerminalOpenLogin()

	;TO-DO: make this work
	;Send("!es")
	;Local $sTerminalScreen = ClipGet()
	;ConsoleWrite($sTerminalScreen & @CRLF)
	;Local $asTerminalScreenArray = StringSplit($sTerminalScreen, ":", 2)
	;_ArrayDisplay($asTerminalScreenArray)
	;_ArraySearch($asTerminalScreenArray, ""

	Send("{Down 4}")
	Send($sPartCode)
	Send("{Enter}")
	Send("{Down 2}")
	Send($iStockNumber)
	Send("{ESC}")

	TermTextWait($g_wTerminal, "Ready", "AfxFrameOrView801", "Stock #: " & $iStockNumber)
	Send("{Left}")
	Sleep(100)
	CaptureScreen($g_wMain, "ImageLink", "ImagingTest")
	;Local $aiAlphaPos = ControlGetPos("AlphaCom", "", "AfxFrameOrView801")
	;Sleep(6000)
	;_ArrayDisplay($aiAlphaPos)
	;ConsoleWrite($aiAlphaPos & @CRLF)
	;Exit
	;MouseMove(@DesktopWidth * .625, @DesktopHeight * .5926)
	MouseClick("primary", @DesktopWidth * .625, @DesktopHeight * .5926)
	WinWait("Image Viewer")
	WinActivate("Image Viewer")
	Sleep(5000)
	CaptureScreen("Image Viewer", "VerifyLink", "ImagingTest")
	Sleep(500)

	#EndRegion -- Imaging Test 8 - Look up the images in alphacom and verify link
	#comments-start
		#Region -- Imaging Test 9 - Print the images from the viewer

		Local $aiViewerWindPos = WinGetPos("Image Viewer")
		Local $aiViewerPos = ControlGetPos("Image Viewer", "", "TImageEnMView1")
		MouseClick("primary", $aiViewerWindPos[0] + $aiViewerPos[0] + 50, $aiViewerWindPos[1] + $aiViewerPos[1] + 75, 2)
		ControlClick("Image Viewer", "", "TAdvPanel1", "primary", 1, 445, 10)
		WinWait("Print")
		ControlClick("Print", "", "TBitBtn2", "primary")
		Sleep(1000)

		#EndRegion -- Imaging Test 9 - Print the images from the viewer

		#Region -- Imaging Test 10 - Email an image using the viewer and verify

		;TO-DO: work this with outlook actually set up
		;	ControlClick("Image Viewer", "", "TAdvPanel1", "primary", 1, 380, 10)
		;	WinWait("E-mail")
		;	ControlClick("Print", "", "TBitBtn2", "primary")
		Sleep(1000)

		#EndRegion -- Imaging Test 10 - Email an image using the viewer and verify

		#Region -- Imaging Test 11 - Verify eBay button works from viewer

		;TO-DO: work this on a part on eBay
		;	ControlClick("Image Viewer", "", "TAdvGlowButton1", "primary")

		#EndRegion -- Imaging Test 11 - Verify eBay button works from viewer
	#comments-end
	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	If WinExists("Image Viewer") Then
		WinClose("Image Viewer")
	EndIf
	;Exit
	Return "Imaging Test Complete"
EndFunc   ;==>TestImaging

#EndRegion --- IMAGING TEST FUNCTION ---

#Region --- REPORTS TEST FUNCTION ---

Func TestReports()
	WinSetState($g_wMain, "", @SW_MAXIMIZE)
	_OpenApp("report")
	If ProcessExists("cView.exe") Then
		ProcessClose("cView.exe")
	EndIf

	#Region -- Reports Test 1 - Verify reports are listed and descriptions show -done

	Local $aiSizes = [16, 21, 15, 6, 6, 15, 2]

	ControlSend($g_wMain, "Launch Report", "TListBox1", "{Tab 2}")

	Local $k = 0
	For $i = 0 To UBound($aiSizes) - 1
		ControlClick($g_wMain, "Launch Report", "TAdvOfficePager2", "primary", 1, 75, 35 + $k)
		For $j = 0 To $aiSizes[$i] - 1
			ControlSend($g_wMain, "Launch Report", "TListBox1", "{Down}")
			;Sleep(200)
			CaptureScreen($g_wMain, "Report" & $i & "," & $j, "ReportsTest")
		Next
		$k += 50
	Next

	#EndRegion -- Reports Test 1 - Verify reports are listed and descriptions show -done

	#Region -- Reports Test 2 - Verify report opens in cView and runs correctly -done
	ControlClick($g_wMain, "Launch Report", "TAdvOfficePager2", "primary", 1, 75, 35)
	ControlSend($g_wMain, "Launch Report", "TListBox1", "{Tab}{Up 15}")
	ControlClick($g_wMain, "Launch Report", "TAdvGlowButton14", "primary")
	WinWait("Enter Values")

	WinMove("Enter Values", "", 0, 0, @DesktopWidth, @DesktopHeight)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 60)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $g_iYardNumber & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 190)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateMinusSeven & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 310)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateToday & "{Enter}")

	Opt("SendKeyDelay", 100)
	MouseClick("primary", 150, 500)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $g_sUserName)
	Opt("SendKeyDelay", 5)

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 340, 480)
	Sleep(250)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, @DesktopWidth - 160, 650)

	WinWait("AdvancedPurchaseOrderReport")
	WinActivate("AdvancedPurchaseOrderReport")
	WinSetState("AdvancedPurchaseOrderReport", "", @SW_MAXIMIZE)
	Sleep(250)
	While StringInStr(WinGetText("AdvancedPurchaseOrderReport"), " / 0")
		Sleep(250)
	WEnd
	CaptureScreen("AdvancedPurchaseOrderReport", "AdvancedPOReport", "ReportsTest")
	WinClose("AdvancedPurchaseOrderReport")

	#EndRegion -- Reports Test 2 - Verify report opens in cView and runs correctly -done

	#Region -- Reports Test 3 - Test a few different reports -done

	;Different Report 1
	;ControlClick($g_wMain, "Launch Report", "TAdvOfficePager2", "primary", 1, 75, 35)
	ControlSend($g_wMain, "Launch Report", "TListBox1", "{Down}")
	ControlClick($g_wMain, "Launch Report", "TAdvGlowButton14", "primary")
	WinWait("Enter Values")

	WinMove("Enter Values", "", 0, 0, @DesktopWidth, @DesktopHeight)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 60)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $g_iYardNumber & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, @DesktopWidth - 160, 140)

	WinWait("DailyAgeReceivable", "100%")
	WinActivate("DailyAgeReceivable")
	WinSetState("DailyAgeReceivable", "", @SW_MAXIMIZE)
	Sleep(250)
	While StringInStr(WinGetText("DailyAgeReceivable"), " / 0")
		Sleep(250)
	WEnd
	CaptureScreen("DailyAgeReceivable", "AgedReceivablesReport", "ReportsTest")
	WinClose("DailyAgeReceivable")
	;Exit

	;Different Report 2
	;ControlClick($g_wMain, "Launch Report", "TAdvOfficePager2", "primary", 1, 75, 35)
	ControlSend($g_wMain, "Launch Report", "TListBox1", "{Down}")
	ControlClick($g_wMain, "Launch Report", "TAdvGlowButton14", "primary")
	WinWait("Enter Values")

	WinMove("Enter Values", "", 0, 0, @DesktopWidth, @DesktopHeight)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 90)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateMinusSeven & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 210)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateToday & "{Enter}")

	Opt("SendKeyDelay", 100)
	MouseClick("primary", 150, 390)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $g_iYardNumber)
	Opt("SendKeyDelay", 5)

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 340, 380)
	Sleep(250)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, @DesktopWidth - 160, 550)

	WinWait("Invoice Credit Report")
	WinActivate("Invoice Credit Report")
	WinSetState("Invoice Credit Report", "", @SW_MAXIMIZE)
	Sleep(250)
	While StringInStr(WinGetText("Invoice Credit Report"), " / 0")
		Sleep(250)
	WEnd
	CaptureScreen("Invoice Credit Report", "InvoiceCreditReport", "ReportsTest")
	WinClose("Invoice Credit Report")

	;Different Report 3
	;ControlClick($g_wMain, "Launch Report", "TAdvOfficePager2", "primary", 1, 75, 35)
	ControlSend($g_wMain, "Launch Report", "TListBox1", "{Down 3}")
	ControlClick($g_wMain, "Launch Report", "TAdvGlowButton14", "primary")
	WinWait("Enter Values")

	WinMove("Enter Values", "", 0, 0, @DesktopWidth, @DesktopHeight)
	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 65)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $g_iYardNumber & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 190)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateMinusSeven & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, 150, 310)
	ControlSend("Enter Values", "", "Internet Explorer_Server1", $tDateToday & "{Enter}")

	ControlClick("Enter Values", "", "Internet Explorer_Server1", "primary", 1, @DesktopWidth - 160, 385)

	WinWait("Invoices by Customer")
	WinActivate("Invoices by Customer")
	WinSetState("Invoices by Customer", "", @SW_MAXIMIZE)
	Sleep(250)
	While StringInStr(WinGetText("Invoices by Customer"), " / 0")
		Sleep(250)
	WEnd
	CaptureScreen("Invoices by Customer", "InvoicesbyCustomer", "ReportsTest")
	WinClose("Invoices by Customer")

	#EndRegion -- Reports Test 3 - Test a few different reports -done

	ControlClick($g_wMain, "", "TAdvGlowButton15")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	Return "Reports Test Complete"

EndFunc   ;==>TestReports



#EndRegion --- REPORTS TEST FUNCTION ---

#Region --- EBAY TEST FUNCTION ---

Func TestEbay()
	_OpenApp("ebay")

	Local $sLoginFileEbayLine = $g_asNonDefaultLogin[5]
	Local $asEbayLineArray = StringSplit($sLoginFileEbayLine, ",", 2)
	Local $iStockNumber = $asEbayLineArray[0]
	Local $sPartCode = $asEbayLineArray[1]

	#Region -- eBay Test 1 - Edit eBay settings

	ControlSend($g_wMain, "", "TAdvOfficePager1", "!fc")
	WinWaitActive("eBay User Configuration")
	WinActivate("eBay User Configuration")
	Sleep(5000)
	ControlClick("eBay User Configuration", "", "TMemo1", "primary")
	Send("^{TAB}")
	If ControlCommand("eBay User Configuration", "", "TComboBox3", "GetCurrentSelection") <> "United States" Then
		ControlCommand("eBay User Configuration", "", "TComboBox3", "SelectString", "United States")
	Else
		ControlCommand("eBay User Configuration", "", "TComboBox3", "SelectString", "Canada")
	EndIf
	Send("^{TAB}")
	ControlCommand("eBay User Configuration", "", "TComboBox3", "SelectString", "Standard Shipping")
	ControlClick("eBay User Configuration", "", "TAdvGlowButton3", "primary")

	#EndRegion -- eBay Test 1 - Edit eBay settings

	#Region -- eBay Test 2 - Edit eBay template
	;TO-DO: make this not a constant timer
	Sleep(10000)
	Local $posStock = ControlGetPos($g_wMain, "", "TAdvEdit3")
	MouseClick("primary", $posStock[0] + 5, $posStock[1] + 5)
	;WinWaitActive($g_wMain)

	ControlSend($g_wMain, "", "TAdvOfficePager1", "!fo")
	WinWait("eBay Template Editor")
	ControlSend("eBay Template Editor", "", "TAdvMemo1", "{PGDN 30}")
	ControlSend("eBay Template Editor", "", "TAdvMemo1", "{Enter 3}")
	ControlSend("eBay Template Editor", "", "TAdvMemo1", "!s")
	ControlClick("eBay Template Editor", "", "TAdvGlowButton5", "primary")
	WinWait("[CLASS:TMessageForm; TITLE:Confirm]")
	ControlClick("Confirm", "", "TButton2", "primary")
	WinClose("eBay Template Editor")

	#EndRegion -- eBay Test 2 - Edit eBay template

	#Region -- eBay Test 3 - Look up stock number
	WinWaitActive($g_wMain)
	ControlSetText($g_wMain, "", "TAdvEdit3", $iStockNumber)
	ControlClick($g_wMain, "", "TAdvGlowButton16", "primary")
	;TO-DO: make this not a constant timer
	Sleep(20000)
	#EndRegion -- eBay Test 3 - Look up stock number

	#Region -- eBay Test 4 - Look up part code
	ControlSetText($g_wMain, "", "Edit1", $sPartCode)
	ControlClick($g_wMain, "", "TAdvGlowButton16", "primary")
	;TO-DO: this one too
	Sleep(20000)
	#EndRegion -- eBay Test 4 - Look up part code

	#Region -- eBay Test 5 - Add part to items to send, edit listing

	$posStock = ControlGetPos($g_wMain, "", "TAdvStringGrid4")
	MouseClick("primary", $posStock[0] + 10, $posStock[1] + 50)
	ControlClick($g_wMain, "", "TAdvGlowButton24", "primary")
	Sleep(1000)
	ControlClick($g_wMain, "", "TAdvStringGrid3", "primary", 1, 55, 35)


	WinWait("eBay Listing Editor")
	ControlCommand("eBay Listing Editor", "", "TCheckBox1", "Check")
	ControlSetText("eBay Listing Editor", "", "TAdvEdit1", "15")
	ControlClick("eBay Listing Editor", "", "TButton1", "primary")
	ControlClick("eBay Listing Editor", "", "TBitBtn2", "primary")

	#EndRegion -- eBay Test 5 - Add part to items to send, edit listing

	Exit

EndFunc   ;==>TestEbay

#EndRegion --- EBAY TEST FUNCTION ---

#Region --- RUNNING TEST CODE ---

_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")
WinActivate($g_wMain)
;ConsoleWrite(TestDashboard() & @CRLF)
;ConsoleWrite(TestSettings() & @CRLF)
;ConsoleWrite(TestTerminal() & @CRLF)
;ConsoleWrite(TestImaging() & @CRLF)
;ConsoleWrite(TestReports() & @CRLF)
ConsoleWrite(TestTrakker() & @CRLF)
;ConsoleWrite(TestEbay() & @CRLF)


Exit 2

#EndRegion --- RUNNING TEST CODE ---


