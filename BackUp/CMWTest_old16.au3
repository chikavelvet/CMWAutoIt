#include <Header.au3>
#include <DashboardTest.au3>
#include <Date.au3>
#Region --- DASHBOARD TEST FUNCTION ---



Func StringArrayToIntArray(ByRef $asStringArray)
	For $i = 0 To UBound($asStringArray) - 1
		$asStringArray[$i] = Int($asStringArray[$i])
	Next
EndFunc   ;==>StringArrayToIntArray

;TO-DO: make this work for the first seven days of the month
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

	#EndRegion -- Dashboard Test 7 - Show details for WO

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
	_OpenApp("term")

	ControlFocus($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]")
	Send("{Enter 4}")
	Sleep(200)
	TermTextWait($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]", "[2J")
	Send($g_sUserPwd & "{Enter}")

	If TermTextWait($g_wTerminal, "Ready", "[CLASS:AfxFrameOrView80]", "Find and Sell") = -1 Then
		Exit
	EndIf

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

	#Region -- Image Test 1 - Look up a stock number

	#EndRegion -- Image Test 1 - Look up a stock number


	Return "Imaging Test Complete"
EndFunc   ;==>TestImaging

#EndRegion --- IMAGING TEST FUNCTION ---

#Region --- RUNNING TEST CODE ---

_OpenWS()
WinSetState($g_wMain, "", @SW_MAXIMIZE)
;ConsoleWrite(TestDashboard() & @CRLF)
;ConsoleWrite(TestTerminal() & @CRLF)
ConsoleWrite(TestImaging() & @CRLF)

Exit 1

#EndRegion --- RUNNING TEST CODE ---


