#include <Header.au3>

_OpenWS()

#Region --- DASHBOARD TEST ---
_OpenApp("dash")

ControlFocus($main, "Dashboard", "TAdvToolBar2")

;WinActivate($main)

Sleep(5000)

ControlSend($main, "", "TAdvToolBar2", "!sg")
For $i = 0 To 24
	Local $rand = Random(0, 1, 1)
	If $rand <> 1 Then
		ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	EndIf
	ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
Next

ControlClick($gadget, "", "TBitBtn2")

Sleep(20000)

ControlSend($main, "", "TAdvToolBar2", "!sg")
For $i = 0 To 24
	ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
Next

ControlClick($gadget, "", "TBitBtn2")

Sleep(20000)


ConsoleWrite("No Problemo 1" & @CRLF)
ControlClick($main, "", "TAdvGlowButton14")
WinWait($message, "Never ask again", 5)
If WinExists($message) Then
	ControlClick($message, "", "TAdvOfficeRadioButton2")
	ControlClick($message, "", "TAdvGlowButton1")
	Sleep(100)
EndIf
#EndRegion --- DASHBOARD TEST ---

#Region --- TERMINAL TEST ---
_OpenApp("term")

ControlFocus($main, "Ready", "AfxFrameOrView801")
Send("{Enter 4}")
Sleep(200)
Send($userPwd & "{Enter}")
Sleep(5000)

Local $t = 0
Local $timeout = 0
Local $worked = 0
while $timeout <> 1 and $worked <> 1
	Send("{!es}")
	If StringInStr(ClipGet(), "Find and Sell") Then
		;ConsoleWrite(ClipGet() & @CRLF)
		$worked = 1
	Else
		If $t >= 10 Then
			$timeout = 1
		EndIf
		$t += 1
	EndIf
WEnd

If $worked = 1 Then
	ConsoleWrite("No Problemo 2" & @CRLF)
	ControlClick($main, "", "TAdvGlowButton14")
	WinWait($message, "Never ask again", 5)
	If WinExists($message) Then
		ControlClick($message, "", "TAdvOfficeRadioButton2")
		ControlClick($message, "", "TAdvGlowButton1")
	EndIf
EndIf

If $timeout = 1 Then
	MsgBox(0, "Terminal Timed Out", "Couldn't find the Find and Sell Screen; looks like something went wrong.")
EndIf

#EndRegion --- TERMINAL TEST ---

#Region --- ORDER TRAKKER TEST ---
;_OpenApp("trak")



#EndRegion --- ORDER TRAKKER TEST ---

Exit
