#include <Header.au3>

_OpenWS()

#Region --- DASHBOARD TEST ---
_OpenApp("dash")

ControlFocus($main, "Dashboard", "TAdvToolBar2")

;WinActivate($main)

For $i = 0 to 24
	Sleep(2000)
	Send("!sg")
	WinWait($gadget)
 	;move i dow n
	if $i <> 0 Then
		for $j = 1 to $i
			ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
			;Sleep(1000)
		Next
	EndIf

	;press space
	ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	;Sleep(1000)
	;move another down
	ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
	;Sleep(1000)
	;press space
	ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	;Sleep(1000)
	;click accept
	ControlClick($gadget, "", "TBitBtn2")
Next

ConsoleWrite("No Problemo" & @CRLF)
ControlClick($main, "", "TAdvGlowButton14")
ControlClick($message, "", "TAdvOfficeRadioButton2")
ControlClick($message, "", "TAdvGlowButton1")
#EndRegion

#Region --- ORDER TRAKKER TEST ---
_OpenApp("trakker")



#EndRegion

Exit