#include <Header.au3>
#include <DashboardTest.au3>
#Region --- DASHBOARD TEST FUNCTION ---

Func StringArrayToIntArray(ByRef $asStringArray)
	For $i = 0 To UBound($asStringArray) - 1
		$asStringArray[$i] = Int($asStringArray[$i])
	Next
EndFunc   ;==>StringArrayToIntArray

Func MinusSeven(ByRef $tDateArray)
	If $tDateArray[1] > 7 Then
		$tDateArray[1] = $tDateArray[1] - 7
	EndIf
EndFunc   ;==>MinusSeven

Func TestDashboard()
	_OpenApp("dash")

	ControlFocus($g_wMain, "Dashboard", "TAdvToolBar2")

	#Region -- Dashboard Test 1 - Turn on all gadgets -done
	Sleep(1000)
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	ControlClick($g_wDate, "", "TButton1")
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	Sleep(200)
	CaptureScreen($g_wMain, "GadgetTest1")
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

		;TO-DO make it wait for changes in the status, instead of 30 sec timeout on updating
		WinWait($g_wMain, "Updating", 30)
		While StringInStr(WinGetText($g_wMain), "Updating")
			Sleep(250)
		WEnd
		;Sleep(100)
		CaptureScreen($g_wMain, "GadgetTest" & ($j + 1))
	Next
	#EndRegion -- Dashboard Test 1 - Turn on all gadgets

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 25}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")

	#Region -- Dashboard Test 2 - Set date for 7 days -done

	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	ControlClick($g_wDate, "Date Range", "TButton1", "primary")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sd")
	ControlCommand($g_wDate, "Date Range", "TCheckBox1", "Check", "")
	Local $tDateArrayMDY = StringSplit(ControlGetText($g_wDate, "Date Range", "TDateTimePicker2") & @CRLF, "/", 2)
	MinusSeven($tDateArrayMDY)
	ControlClick($g_wDate, "Date Range", "TDateTimePicker1", "primary")
	ControlSend($g_wDate, "Date Range", "TDateTimePicker1", "{Right 2}" & $tDateArrayMDY[1])
	;Exit
	ControlClick($g_wDate, "", "TBitBtn1", "primary")

	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	Sleep(100)

	CaptureScreen($g_wMain, "Gadget7DaysTest")

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

		WinWait($g_wMain, "Updating", 30)
		While StringInStr(WinGetText($g_wMain), "Updating")
			Sleep(250)
		WEnd
		;Sleep(100)
		CaptureScreen($g_wMain, "Gadget7DaysTest" & ($j + 1))
	Next

	#EndRegion -- Dashboard Test 2 - Set date for 7 days

	;Exit
	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 25}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")




	#Region -- Dashboard Test 4 - Display Only WO for "NameLoggedIn" -done

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down 2}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
	ControlClick($g_wGadget, "", "TBitBtn2")
	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd

	CaptureScreen($g_wMain, "WOs_ALL")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	CaptureScreen($g_wMain, "WOs_(" & $g_sUserName & ")")
	ControlSend($g_wMain, "", "TAdvToolBar2", "!si")
	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd

	#EndRegion -- Dashboard Test 4 - Display Only WO for "NameLoggedIn"

	#Region -- Dashboard Test 5 - Expand and retract gadgets
	#EndRegion -- Dashboard Test 5 - Expand and retract gadgets

	#Region -- Dashboard Test 6 - Change columns to display
	#EndRegion -- Dashboard Test 6 - Change columns to display

	#Region -- Dashboard Test 7 - Show details for WO
	#EndRegion -- Dashboard Test 7 - Show details for WO

	#Region -- Dashboard Test 3 - Verify resizing -done

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	For $i = 0 To 12
		If Random(0, 1, 1) = 1 Then
			ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		EndIf
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next


	ControlClick($g_wGadget, "", "TBitBtn2")
	Sleep(250)
	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	Sleep(100)
	CaptureScreen($g_wMain, "Resize")

	ControlSend($g_wMain, "", "TAdvToolBar2", "!sg")
	For $i = 0 To 12
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Space}")
		ControlSend($g_wGadget, "", "TAdvStringGrid1", "{Down}")
	Next
	ControlClick($g_wGadget, "", "TBitBtn2")

	WinWait($g_wMain, "Updating", 30)
	While StringInStr(WinGetText($g_wMain), "Updating")
		Sleep(250)
	WEnd
	Sleep(100)

	CaptureScreen($g_wMain, "Resize2")

	#EndRegion -- Dashboard Test 3 - Verify resizing




	Exit



	Sleep(20000)
	Sleep(20000)

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

#Region --- RUNNING TEST CODE ---

_OpenWS()
WinSetState($g_wMain, "", @SW_MAXIMIZE)
ConsoleWrite(TestDashboard() & @CRLF)
;ConsoleWrite(TestTerminal() & @CRLF)

;Exit 1

#EndRegion --- RUNNING TEST CODE ---


