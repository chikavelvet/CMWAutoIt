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

	--- BEFORE STARTING TEST MAKE SURE THE FOLLOWING IS DONE ---
	- No tabs are set to open on start-up
	- Only the first gadget (Uninventoried Vehicles) is checked in Dashboard
	- No lock timeout is set up
	- The CMWTest.csv file has:
		- A valid Yard Owner username and password
		- The correct Password of the Day
		- A valid stock number and part code
		- A valid CFT file path for Alphacom
		- A valid stock number and part code (can be same as or different to first one)
	- The CMWTestInvLogin.csv file has:
		- A valid Inventory-only username and password
		- The correct Password of the Day
		- 'image' in the second line
		- A valid stock number and part code (can be same as or different to first two

	(This information is also in "PRE-TEST CHECKLIST.txt")

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

;initializing today's date (MDY) and copying it to another variable (to be -7'd)
Local $tDateToday = [@MON, @MDAY, @YEAR]
Local $tDateMinusSeven = $tDateToday

;subtracting seven days from today's date
MinusSeven($tDateMinusSeven)

;putting dates into formatted strings (from arrays) for use in Reports test
;TO-DO: this overlaps with things done in Dashboard test, so try to somehow combine them
$tDateToday = $tDateToday[2] & "-" & $tDateToday[0] & "-" & $tDateToday[1]
$tDateMinusSeven = $tDateMinusSeven[0] & "-" & $tDateMinusSeven[1] & "-" & $tDateMinusSeven[2]

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

#comments-start

-- AccessCMWToolbar --
This function uses a single controlled click and regular input to access any
of the menus from the main toolbar in CMW. This includes File, Settings, and Help,
and any of the subitems in these three categories. The function uses a zero-based
three dimensional indexing system, though the third dimension is optional, and
only used in Settings->Security, which then has more options to choose from.

- $i, $j, $k: Integers -
These three integer inputs form a three dimensional coordinate that determines
what specific toolbar function to access. The first value determines whether
to open File, Settings, or Help initially; the second value determines which
subitem to go to and open; and the third (optional) value is used if the subitem
in turn presents multiple subitems to access (only currently found in Settings->Security).

Note: the $k value is defaulted to zero, but also not used unless $i and $j both
are equal to 1 (Setting->Security case).

#comments-end
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

	;Open up Security for Dashboard menu (try again every 5 seconds if it doesn't work)
	While Not WinExists("Security for Dashboard")
		AccessCMWToolbar(1, 1, 0)
		WinWait("Security for Dashboard", "", 5)
	WEnd

	;Down 3 Up 1 puts the selection highlight on the first input box
	;set first input box to 1
	Send("{Down 3}{Up}1")

	;Set all Dashboard gadget securities to 1 (Yard Owner)
	For $i = 0 To 24
		Send("{Down}1")
	Next
	;OK button
	ControlClick("Security for Dashboard", "", "TBitBtn1", "primary")

	;re-open security menu and take screenshot to verify settings saved
	While Not WinExists("Security for Dashboard")
		AccessCMWToolbar(1, 1, 0)
		WinWait("Security for Dashboard", "", 5)
	WEnd
	CaptureScreen($g_wMain, "DashboardSecurity", "SettingsTest")
	ControlClick("Security for Dashboard", "", "TBitBtn1", "primary")

	#EndRegion -- Settings Test 1 - Edit security settings so only yard owner has access to all dashboard gadgets


	#Region -- Settings Test 2 - Set settings to eBay so only YO can list parts, but anyone with sales can view tab

	;Open up Security for eBay menu (try again every 5 seconds if it doesn't work)
	While Not WinExists("Security for eBay")
		AccessCMWToolbar(1, 1, 1)
		WinWait("Security for eBay", "", 5)
	WEnd

	;Down 3 Up 1 puts the selection highlight on the first input box
	;set first input box to 1
	Send("{Down 3}{Up}1")
	;set second input box to 2,3 (sales, sales manager)
	Send("{Down}2,3")
	;OK button
	ControlClick("Security for eBay", "", "TBitBtn1", "primary")

	;re-open security menu and take screenshot to verify settings saved
	While Not WinExists("Security for eBay")
		AccessCMWToolbar(1, 1, 1)
		WinWait("Security for eBay", "", 5)
	WEnd
	CaptureScreen($g_wMain, "eBaySecurity", "SettingsTest")

	#EndRegion -- Settings Test 2 - Set settings to eBay so only YO can list parts, but anyone with sales can view tab

	#Region -- Settings Test 3 - Give imaging rights so that only inventory can add images

	;Open up Security for Imaging menu (try again every 5 seconds if it doesn't work)
	While Not WinExists("Security for CMIS")
		AccessCMWToolbar(1, 1, 2)
		WinWait("Security for CMIS", "", 5)
	WEnd

	;Imaging security menu only has one box, so Down 3 Up 1 doesn't work
	;use controlclick to select the first/only input box
	;TO-DO: see if there's a more stable way to select this (via keyboard)
	ControlClick("Security for CMIS", "", "TAdvStringGrid1", "primary", 1, 180, 30)
	;set first/only input box to 5,6 (inventory, inventory manager)
	Send("5,6")
	;OK Button
	ControlClick("Security for CMIS", "", "TBitBtn1", "primary")

	;re-open security menu and take screenshot to verify settings saved
	While Not WinExists("Security for CMIS")
		AccessCMWToolbar(1, 1, 2)
		WinWait("Security for CMIS", "", 5)
	WEnd
	CaptureScreen($g_wMain, "ImagingSecurity", "SettingsTest")

	#EndRegion -- Settings Test 3 - Give imaging rights so that only inventory can add images

	#Region -- Settings Test 4 - Login as a user with only inventory rights, verify imaging works
	;use the CMWTestInvLogin csv file to reopen CMW and log in as a user with only inventory rights
	;TO-DO: (maybe) make a user with these rights via script rather than forcing tester to have one ready
	;TO-DO: combine CMWTestInvLogin.csv with CMWTest.csv (so everything's in one file)
	_OpenWS(@AppDataDir & "\AutoIt\CMWTestInvLogin.csv")

	;run the first 3 tests in the Imaging Test (sufficient for checking if user privileges is working)
	ImageTest123("SettingsTest")

	#EndRegion -- Settings Test 4 - Login as a user with only inventory rights, verify imaging works

	#Region -- Settings Test 5 - Try to open dashboard, get no data
	;Open the Dashboard tab
	_OpenApp("dash")
	;wait for any open dashboard gadgets to update
	;Note: Settings Test should be performed after Dashboard Test so that the resulting randomly opened
	;gadgets from the dashboard test work to sufficiently test user privileges in the Settings Test
	;Note 2: WaitForUpdating has a bug and may not work completely as intended (only waits for first gadget)
	;Note 3: Even if user does not have correct privileges, CMW still allows them to open up the Dashboard
	;tab, but every gadget's information is blank. I am unsure if this is intended behavior or not
	WaitForUpdating()
	CaptureScreen($g_wMain, "NotYODashboard", "SettingsTest")
	#EndRegion -- Settings Test 5 - Try to open dashboard, get no data

	#Region -- Settings Test 6 - Try to open eBay, get security error
	;open eBay tab (doesn't use OpenApp() because it wouldn't finish due to the tab never fully opening)
	;TO-DO: (potentially) modify OpenApp to not cause this problem
	ControlClick($g_wMain, "", "TAdvGlowButton9", "primary")
	;wait for security error window
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "You do not have security to", 10)
	CaptureScreen($g_wMain, "NotYOeBay", "SettingsTest")

	#EndRegion -- Settings Test 6 - Try to open eBay, get security error

	#Region -- Settings Test 7 - Login as yard owner, should be able to see all gadgets, and access ebay
	;re-open workstation and log in using the normal CMWTest.csv (Yard Owner) account
	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")
	;open Dashboard tab
	_OpenApp("dash")
	;see notes above (in Settings Test 5 region) on WaitForUpdating()
	WaitForUpdating()
	CaptureScreen($g_wMain, "YODashboard", "SettingsTest")

	;close Dashboard tab
	;TO-DO: make a CloseTab() function that does this stuff)
	ControlClick($g_wMain, "", "TAdvGlowButton14")
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
		Sleep(100)
	EndIf

	;open eBay app
	_OpenApp("ebay")
	;wait for a recurring eBay error I've been getting (cannot connect to ebay)
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 10)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
		CaptureScreen($g_wMain, "ErrorImage" & $g_iErrorCount)
	EndIf
	CaptureScreen($g_wMain, "YOeBay", "SettingsTest")

	Sleep(10000)

	;close eBay tab
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
	;open Imaging app (this does not use OpenApp(), see Settings Test 6 for details)
	ControlClick($g_wMain, "", "TAdvGlowButton10", "primary")

	;wait for security error and if/when it shows up, take a screenshot close the window
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 15)
	CaptureScreen($g_wMain, "YOImaging", "SettingsTest")
	ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
	#EndRegion -- Settings Test 8 - (as YO), Should not be able to add images to images

	#Region -- Settings Test 9 - Change lockout setting to 3 minutes, let PC sit and verify ws prompt appears
	;open up the CMW setup menu (try again every 5 minutes if it doesn't open)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd

	;Ctrl+Tab to the 'additional' settings menu
	Send("^{Tab 5}")

	;set the timeout to 3 minutes
	ControlSetText("Setup", "", "TAdvSpinEdit1", "3")

	;save settings
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	;this will wait for the timeout (350 seconds, though it will stop waiting upon 3 minute timeout)
	;it will then re-log back in to CMW
	;the first part works fine, the second part has some issues
	;	WinWaitNotActive($g_wMain, "", 350)
	;	ConsoleWrite("Not active" & @CRLF)
	;	WinActivate($g_wPassword)
	;	WinSetState($g_wPassword, "", @SW_SHOW)
	;	WinActivate($g_wMain)
	;Requires user input here
	;TO-DO: make it not require user input
	;LogIn(@AppDataDir & "\AutoIt\CMWTest.csv")
	#EndRegion -- Settings Test 9 - Change lockout setting to 3 minutes, let PC sit and verify ws prompt appears

	#Region -- Settings Test 10 - Set different tools to open automatically, verify they do so
	;open Security for Imaging (try again every 5 seconds if it doesn't work)
	While Not WinExists("Security for CMIS")
		AccessCMWToolbar(1, 1, 2)
		WinWait("Security for CMIS", "", 5)
	WEnd

	;set imaging security up to work for Yard Owner again, so it doesn't give security error when it opens
	ControlClick("Security for CMIS", "", "TAdvStringGrid1", "primary", 1, 180, 30)
	Send("1,5,6")
	ControlClick("Security for CMIS", "", "TBitBtn1", "primary")

	;open up CMW Setup menu (try again every 5 seconds if it doesn't work)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd

	;Ctrl+Tab to Tab Setup menu
	Send("^{Tab 4}")

	;randomly check tabs to start up
	For $i = 0 To 6
		Send("{Tab}")
		If Random(0, 1, 1) = 1 Then
			Send("{Space}")
		EndIf
	Next
	CaptureScreen($g_wMain, "RandomTabsToOpen1", "SettingsTest")
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	;open WS to see if proper tabs show up
	_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")
	;wait for and deal with eBay error message
	WinWait("[CLASS:#32770; TITLE:Checkmate Workstation]", "", 30)
	If WinExists("[CLASS:#32770; TITLE:Checkmate Workstation]") Then
		ControlClick("[CLASS:#32770; TITLE:Checkmate Workstation]", "", "Button1", "primary")
		CaptureScreen($g_wMain, "ErrorImage" & $g_iErrorCount)
	EndIf
	WinActivate($g_wMain)

	;'waits' until it can open setup menu (all apps are finished loading)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	$posSetupWin = WinGetPos("Setup")
	MouseClick("primary", $posSetupWin[0] + 75, $posSetupWin[1] + 550)

	CaptureScreen($g_wMain, "RandomTabsOpened1", "SettingsTest")

	;open up Setup menu, go to tab setup, and invert all checkboxes
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

	;open WS to see if proper tabs show up, deal with eBay error if it happens, wait until setup menu can open
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
	;open CMW setup menu (try again every 5 seconds if it doesn't open)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
	WinActivate("Setup")
	;Ctrl+Tab over to printer settings, change the RPT printer to be one down from what it is
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

#comments-start

	-- WaitForUpdating --
	This function waits for all active dashboard gadgets to
	finish updating. It does this by looking for the string
	"Updating" in the window.

	- $tExtraWait: Integer -
	This parameter introduces an additional amount of constant
	wait time to the end of the function. The input value is
	in seconds (for convenience). The default wait time is .1
	seconds.

	- $wWindow: Window -
	This parameter specifies a specific window to look in when
	looking for "Updating". By default (and in all current use
	cases) it is set to the main CMW window.

	Note: This function has a bug because it only looks at the
	first gadget's updating text. So if the first gadget loads
	before any of the other on-screen gadgets, the function will
	not work as intended. There may be a way to fix this via more
	in-depth string analysis (and potentially regex) but with the
	extra wait time functionality it hasn't become too much of
	a problem and I have yet to look into it fully.

	Note 2: There are multiple cases in these tests (and more
	potential future cases) in which pausing the script while
	a certain string is present in the window is useful. Therefore
	this function could be modified (or a new one created) to
	more generally search a window for a specified string and
	wait for the string to not be present.
	(modifying this function would take a change of only a few words)
	It also may be useful to make a second function that does
	the opposite, waiting until the string is present (already
	somewhat implemented in TermTextWait.)

#comments-end
Func WaitForUpdating($tExtraWait = .1, $wWindow = $g_wMain)
	While StringInStr(WinGetText($wWindow), "Updating")
		Sleep(250)
	WEnd
	Sleep($tExtraWait * 1000)
EndFunc   ;==>WaitForUpdating

Func TestDashboard()
	;maximize window (does nothing if already maximized)
	WinSetState($g_wMain, "", @SW_MAXIMIZE)

	;open dashboard
	_OpenApp("dash")

	Local $sCurrentStatus

	ControlFocus($g_wMain, "Dashboard", "TAdvToolBar2")

	#Region -- Dashboard Test 1 - Turn on all gadgets -done
	Sleep(1000)
	$sCurrentStatus = ControlGetText($g_wMain, "", "TStatusBar1")

	;first, ensure the date is a single day by setting the date to today
	;open up date settings menu
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)
	;click 'today' button
	ControlClick($g_wDate, "", "TButton1")
	WaitForUpdating()
	Sleep(200)
	CaptureScreen($g_wMain, "GadgetTest1", "DashboardTest")
	;open up gadget settings menu and activate each gadget, one at a time
	;take screenshots after the gadget loads
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

	;reset open gadgets to just be "uninventoried vehicles"
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 25}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")

	#Region -- Dashboard Test 2 - Set date for 7 days -done
	;TO-DO: Skip gadgets that don't use date range
	WaitForNewGadget($sCurrentStatus)
	;open up date settings menu
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)

	;click "today" to ensure it's set to today's date initially
	ControlClick($g_wDate, "Date Range", "TButton1", "primary")

	;reopen date settings menu
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	WinWait($g_wDate)

	;check the date range checkbox
	ControlCommand($g_wDate, "Date Range", "TCheckBox1", "Check", "")

	;calculate today's date (formatted) and the date seven days ago
	Local $tDateArrayMDY = StringSplit(ControlGetText($g_wDate, "Date Range", "TDateTimePicker2") & @CRLF, "/", 2)
	MinusSeven($tDateArrayMDY)
	;set the date fields to match the calculated values
	ControlClick($g_wDate, "Date Range", "TDateTimePicker1", "primary")
	ControlSend($g_wDate, "Date Range", "TDateTimePicker1", "{Right 2}" & $tDateArrayMDY[2])
	;Exit
	ControlClick($g_wDate, "", "TBitBtn1", "primary")

	;wait for gadgets to update with new date range
	WaitForUpdating()
	Sleep(100)

	CaptureScreen($g_wMain, "Gadget7DaysTest", "DashboardTest")

	;cycle through every gadget again, this time with the new date range
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
	;reset open gadgets back to only "uninventoried vehicles"
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

	;open a couple WO related gadgets
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

	;change setting to only view WOs by active user
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WaitForUpdating(.5)
	CaptureScreen($g_wMain, "WOs_(" & $g_sUserName & ")", "DashboardTest")
	;change it back
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WaitForUpdating(.5)

	#EndRegion -- Dashboard Test 4 - Display Only WO for "NameLoggedIn" -done

	#Region -- Dashboard Test 5 - Expand and retract gadgets -done
	;TO-DO: make sure all gadget opens wait for the window afterward
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)

	;open some more gadgets (first two)
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")

	ControlClick($g_wGadget, "", "TBitBtn2")
	WaitForNewGadget($sCurrentStatus)
	WaitForUpdating(1)

	;expand and contract the first two gadgets on the screen
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

	;right click on grid, go down to column display options
	;TO-DO: Change this to a Ctrl+F10 statement
	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
	Sleep(100)
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
	WinWait($g_wDashSett)

	;deselect the first 3 columns to display
	For $n = 0 To 2
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wDashSett, "", "TBitBtn2", "primary")

	;update grid and screenshot
	WaitForUpdating(1)
	CaptureScreen($g_wMain, "DisplayColumnsChanged", "DashboardTest")
	Sleep(100)

	;re-open column display options
	ControlClick($g_wMain, "", "TAdvStringGrid1", "secondary")
	Sleep(100)
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wMain, "", "TAdvStringGrid1", "{Enter}")
	WinWait($g_wDashSett)

	;re-select first 3 columns to display
	For $n = 0 To 2
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wDashSett, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wDashSett, "", "TBitBtn2", "primary")

	;update grid
	WaitForUpdating(1)


	#EndRegion -- Dashboard Test 6 - Change columns to display -done

	#Region -- Dashboard Test 7 - Show details for WO -done
	;this works more or less but seems to fail depending on the workstation environment
	;(it works on a normal-type workstation with data, but not on some QA-types)
	#comments-start
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
	#comments-end
	#EndRegion -- Dashboard Test 7 - Show details for WO -done

	#Region -- Dashboard Test 3 - Verify resizing -done

	;open gadget settings menu
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)

	;within the first 12 gadgets, randomly select/deselect them
	For $i = 0 To 12
		If Random(0, 1, 1) = 1 Then
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		EndIf
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wGadget, "", "TBitBtn2")
	Sleep(250)

	;wait for updating/resizing
	WaitForUpdating()
	Sleep(100)
	CaptureScreen($g_wMain, "Resize", "DashboardTest")

	;re-open gadget settings menu
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	WinWait($g_wGadget)

	;within the first 12 gadgets, invert each gadget's selection
	For $i = 0 To 12
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wGadget, "", "TBitBtn2")

	;wait for updating/resizing
	WaitForUpdating()
	Sleep(100)

	CaptureScreen($g_wMain, "Resize2", "DashboardTest")

	#EndRegion -- Dashboard Test 3 - Verify resizing -done

	;close dashboard
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

#comments-start

	-- CheckTermSettings --
	This function is designed to check the terminal settings tab in the CMW setup
	menu, and make sure the connection file is correct. By default, it checks it
	against the global $g_fiDefaultConnectionFile, but a parameter can be passed
	in to make it check against the parameter. In the CMWTest, the file to use
	as the connection file is specified in the CMWTest.csv.

	- $fiConnectionFile: File Path -
	This specifies a file path to use for the AlphaCom connection file (should be
	a valid connection file.) Default is the $g_fiDefaultConnectionFile.

#comments-end
Func CheckTermSettings($fiConnectionFile = $g_fiDefaultConnectionFile)
	While Not WinExists("Setup")
		AccessCMWToolbar(1, 0)
		WinWait("Setup", "", 5)
	WEnd
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
	WinWait($g_wMessage, "Never ask again", 5)
	If WinExists($g_wMessage) Then
		ControlClick($g_wMessage, "", "TAdvOfficeRadioButton2")
		ControlClick($g_wMessage, "", "TAdvGlowButton1")
	EndIf
	Sleep(500)
	WinSetState($g_wMain, "", @SW_SHOWNORMAL)
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
;NOTE: These variables and functions are deprecated.

;an array of valid pixel positions to click to access each tab (updates)
Global $aiOTTabPos[15] = [30]
;an array of the number of parts in each tab (updates)
Global $aiOTTabNum[15]

;an array of each tab's name, used to look-up a tab index based on name
Global $asOTNameIndex[15] = [ _
		"Dispatch", _
		"Warehouse", "Dismantling", "Yard", "Brokered", _
		"Arrived", "Void", "RdySnd", _
		"CPU", "Truck", "LTL", "FedEx/UPS", _
		"Returned", "Delivered", "Restocked" _
		]

;a constant array of the base widths of each tab
Global Const $aiOTTabBaseWidths[15] = [80, 94, 100, 59, 78, 66, 59, 68, 60, 60, 60, 93, 79, 83, 88]
;an array of the currect widths of each tab (deprecated, updated)
Global $aiOTTabCurrentWidths[15] = [80, 94, 100, 59, 78, 66, 59, 68, 60, 60, 60, 93, 79, 83, 88]
;a string denoting the last active tab (for regex search purposes)
Global $sLastActiveTab = "Restocked"

#comments-start

	-- OTGetActiveTab --
	This function gets the name and number of parts in the active tab. It does
	this primarily through getting the text of the window, stripping it of all
	whitespace and carriage returns, and using REGEX to look for where the name
	of the active tab must be (between "Dispatch" and the last active tab name).
	It returns this as a string.

#comments-end
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

#comments-start

	-- OTGetActiveTabIndex --
	This function uses OTGetActiveTab() to find the active tab's name, and then
	checks it against $asOTNameIndex to see what the index number of the active
	tab is. It returns this as an integer.

#comments-end
Func OTGetActiveTabIndex()
	Local $sActiveTab = OTGetActiveTab()
	Local $asActiveInfoSplit = StringSplit($sActiveTab, "()")
	Local $iTabIndex = _ArraySearch($asOTNameIndex, $asActiveInfoSplit[1])
	Return $iTabIndex
EndFunc   ;==>OTGetActiveTabIndex

#comments-start

	-- OTSetActiveTabNum --
	This function uses OTGetActiveTab() to get the number of parts
	in the active tab. It then updates the $aiOTTabNum array accordingly.

#comments-end
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
	EndIf
EndFunc   ;==>OTSetActiveTabNum

#comments-start

	-- OTSwitchToTab --
	This function uses the $aiOTTabPos array and ControlClick() to switch tabs.
	It only attempts the tab switch if the active tab is not the same as the tab
	to switch to. After switching tabs, it updates $sLastActiveTab.

	- $iTargetTabIndex: Integer -
	This parameter is an integer denoting the index of the target tab.

#comments-end
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

#comments-start

	-- OTSetInitialTabPos --
	This function is a function run at the beginning of the test to set the
	initial positions of the tabs based on tab number. After this, the position
	array automatically updates when it needs to, so repeating this function
	is unnecessary. However, this function could be run at any time to accurately
	update the tab positions.

	Note: This function is redundant, and could be removed. It is also deprecated.

#comments-end
Func OTSetInitialTabPos()
	OTSetInitialNums()
	OTUpdatePos()
EndFunc   ;==>OTSetInitialTabPos

#comments-start

	-- OTSetInitialNums --
	This function updates the tab numbers (number of parts in the tab) of every
	tab initially. After running this function once at the beginning, tab numbers
	are updated as they are changed, so repeating this function is unnecessary.
	However, this function can be run at any time to update all tab numbers.
	This function also calls OTUpdatePos() after setting each tab numbers,
	which effectively updates all tab positions as well.

#comments-end
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

#comments-start

	-- OTUpdatePos --
	This function is the heart of the tab tracking system. It goes through
	the $aiOTTabPos array, and for each index calculates the position of the
	next index based on the current tab's position, tab width, and tab number.
	(The first tab position is initialized to 30, which isn't affected by any
	tab numbers.)

#comments-end
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

#comments-start

	-- OTSendPartToTab --
	This function sends the top part (or part selected) in the active tab
	to another tab specified by the parameter. It then updates both tabs'
	tab numbers, and updates the position of all tabs accordingly.

	- $iTargetTabIndex: Integer -
	This variable holds the index of the targetted tab for the part move.

#comments-end
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

#comments-start

	-- ImageTest123 --
	This function completes the first three Imaging tests. It is a separate
	function because it is useful during the Settings Test as well as the
	Imaging Test (and it would be wasteful to write the code twice).

	- $sTestSubDir: String -
	This parameter designates what subfolder to save any captured screenshots
	in. By default, it's set to "ImagingTest", but it is useful to set it to
	"SettingsTest" for use during the Settings Test.

#comments-end
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
			Sleep(100)
			Send($i)
			Sleep(100)
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

	ControlClick($g_wMain, "", "TAdvGlowButton21")
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
ConsoleWrite(TestTerminal() & @CRLF)
ConsoleWrite(TestImaging() & @CRLF)
ConsoleWrite(TestReports() & @CRLF)
;ConsoleWrite(TestTrakker() & @CRLF)
;ConsoleWrite(TestEbay() & @CRLF)
;ConsoleWrite(TestSettings() & @CRLF)

Exit 2

#EndRegion --- RUNNING TEST CODE ---


