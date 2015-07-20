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


ConsoleWrite("No Problemo" & @CRLF)
ControlClick($main, "", "TAdvGlowButton14")
WinWait($message, "Never ask again", 5)
If WinExists($message) Then
	ControlClick($message, "", "TAdvOfficeRadioButton2")
	ControlClick($message, "", "TAdvGlowButton1")
EndIf
#EndRegion --- DASHBOARD TEST ---

#Region --- TERMINAL TEST ---
_OpenApp("term")



#EndRegion

#Region --- ORDER TRAKKER TEST ---
;_OpenApp("trak")



#EndRegion --- ORDER TRAKKER TEST ---

Exit
