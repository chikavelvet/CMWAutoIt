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
#include <DashboardTest.au3>
#include <Date.au3>


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
Func TerminalOpenLogin()
	_OpenApp("term")

	ControlFocus($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]")
	Send("{Enter 4}")
	Sleep(200)
	TermTextWait($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]", "[2J")
	Send($g_sUserPwd & "{Enter}")

	If TermTextWait($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]", "Find and Sell") = -1 Then
		MsgBox(0, "No Find and Sell Found", "Something went wrong.")
		Exit
	EndIf
EndFunc   ;==>TerminalOpenLogin

#Region --- DASHBOARD TEST FUNCTION ---
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
#comments-start
#comments-end
;TO-DO: make this customizable
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
		ControlClick($g_wGadget, "", "TBitBtn2")

		WaitForNewGadget($sCurrentStatus)
		Sleep(250)
		CaptureScreen($g_wMain, "GadgetTest" & ($j + 1), "DashboardTest")
		Sleep(250)
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

Func TestTerminal()


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
;_OpenApp("trak")



#EndRegion --- ORDER TRAKKER TEST FUNCTION ---

#Region --- IMAGING TEST FUNCTION ---

Func TestImaging()
	_OpenApp("image")
	#comments-start
	#Region -- Imaging Test 1 - Look up a stock number -done
	Local $sLoginFileImagingLine = $g_asNonDefaultLogin[2]
	Local $asImagingLineArray = StringSplit($sLoginFileImagingLine, ",", 2)
	Local $iStockNumber = $asImagingLineArray[0]
	Local $sPartCode = $asImagingLineArray[1]
	;ConsoleWrite($iStockNumber & @CRLF)
	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 5)
	ControlSetText($g_wMain, "Imaging", "TAdvEdit4", $iStockNumber)
	ControlClick($g_wMain, "Imaging", "TAdvGlowButton15", "primary")
	WinWait($g_wMain, $iStockNumber)

	CaptureScreen($g_wMain, "StockNumber" & $iStockNumber & "Lookup", "ImagingTest")

	#EndRegion -- Imaging Test 1 - Look up a stock number -done

	#Region -- Imaging Test 2 - Look up a specific part -done

	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 5)
	ControlSetText($g_wMain, "Imaging", "TAdvEdit2", $sPartCode)
	ControlClick($g_wMain, "Imaging", "TAdvGlowButton15", "primary")
	WinWait($g_wMain, $sPartCode)

	CaptureScreen($g_wMain, "SpecificPart" & $sPartCode, "ImagingTest")

	#EndRegion -- Imaging Test 2 - Look up a specific part -done

	#Region -- Imaging Test 3 - Add single image using browse image search window -done

	ControlClick($g_wMain, "Imaging", "TAdvToolPanelTab1", "primary", 1, 5, 150)
	ControlClick($g_wMain, "Imaging", "TDirectoryListBoxEx1", "primary", 2, 1, 1)
	Local $asPicture = StringSplit($g_fiScreenCaptureFolder, "\")
	;_ArrayDisplay($asPicture)
	Opt("SendKeyDelay", 100)
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
			Sleep(100)
			ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Down 2}")
			Send($i)
			ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Enter}")
			Sleep(100)
		EndIf
		$j += 1
	Next
	ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Down 2}")
	Send("ImagingTest")
	ControlSend($g_wMain, "", "TDirectoryListBoxEx1", "{Enter}")
	WinWait($g_wMain, "c: [OS]", 2)

	ControlFocus($g_wMain, "Imaging", "TAdvStringGrid1")
	ControlSend($g_wMain, "Imaging", "TAdvStringGrid1", "{Space}")
	Local $aiImageFromPos = ControlGetPos($g_wMain, "Imaging", "TImageEnMView2")
	Local $aiImageToPos = ControlGetPos($g_wMain, "Imaging", "TImageEnMView1")
	MouseClickDrag("primary", $aiImageFromPos[0] + 175, $aiImageFromPos[1] + 50, $aiImageToPos[0] + 50, $aiImageToPos[1] + 50)
	WinWait("Confirm", "Filename being copied", 5)
	While WinExists("Confirm", "Filename being copied")
		ControlClick("Confirm", "Filename being copied", "Button1", "primary")
		WinWait("Confirm", "Filename being copied", 2.5)
	WEnd
	Sleep(2500)
	CaptureScreen($g_wMain, "AddSinglePartFromBrowse", "ImagingTest")

	#EndRegion -- Imaging Test 3 - Add single image using browse image search window -done

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
	#comments-end
	#Region -- Imaging Test 6 - Verify import images works

	CaptureScreen($g_wMain, Null, "ImagingTest")
	Send("!fi")
	WinWait("Password of the Day")
	ControlSetText("Password of the Day", "", "TEdit1", $g_sPOD)
	ControlClick("Password of the Day", "", "TBitBtn2", "primary")

	WinWait("CMW - Image Import")
	Local $fiCurImageDestPath = ControLgetText("CMW - Image Import", "", "TAdvEdit1")
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
	Send($fiCurImageDestPath & "\" & @YEAR &  "{Enter}")
	WinWait(@YEAR, "Address")
	CaptureScreen(@YEAR, "ImportedLocation", "ImagingTest")

	#EndRegion -- Imaging Test 6 - Verify import images works

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
	$sReplaced = StringLeft($sReplaced,StringLen($sReplaced)-1)
	;Exit
	$sReplaced = "#0" & StringRight($sReplaced, StringLen($sReplaced)-2)
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
	#comments-start
	#Region -- Imaging Test 8 - Look up the images in alphacom and verify link

	TerminalOpenLogin()

	;TO-DO: make this work
	;Send("!es")
	;Local $sTerminalScreen = ClipGet()
	;ConsoleWrite($sTerminalScreen & @CRLF)
	;Local $asTerminalScreenArray = StringSplit($sTerminalScreen, ":", 2)
	;_ArrayDisplay($asTerminalScreenArray)
	;_ArraySearch($asTerminalScreenArray, ""

	;Opt("SendKeyDelay", 500)

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
	Exit
	Return "Imaging Test Complete"
EndFunc   ;==>TestImaging

#EndRegion --- IMAGING TEST FUNCTION ---

#Region --- RUNNING TEST CODE ---

_OpenWS(@AppDataDir & "\AutoIt\CMWTest.csv")
WinActivate($g_wMain)
WinSetState($g_wMain, "", @SW_MAXIMIZE)
;ConsoleWrite(TestDashboard() & @CRLF)
;ConsoleWrite(TestTerminal() & @CRLF)
ConsoleWrite(TestImaging() & @CRLF)

Exit 1

#EndRegion --- RUNNING TEST CODE ---


