#include <Header.au3>

_OpenWS()

#Region --- DASHBOARD TEST ---
_OpenApp("dash")

ControlFocus($main, "Dashboard", "TAdvToolBar2")

;WinActivate($main)

#comments-start
For $i = 0 To 24
	Sleep(2000)
	Send("!sg")
	WinWait($gadget)
	;move i down
	If $i <> 0 Then
		For $j = 1 To $i
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
#comments-end

Sleep(5000)

ControlSend($main, "", "TAdvToolBar2", "!sg")
for $i = 0 to 24
	Local $rand = Random(0,1,1)
	If $rand <> 1 Then
		ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	EndIf
	ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
Next

ControlClick($gadget, "", "TBitBtn2")

Sleep(20000)

ControlSend($main, "", "TAdvToolBar2", "!sg")
for $i = 0 to 24
	ControlSend($gadget, "", "TAdvStringGrid1", "{Space}")
	ControlSend($gadget, "", "TAdvStringGrid1", "{Down}")
Next

ControlClick($gadget, "", "TBitBtn2")

Sleep(20000)


ConsoleWrite("No Problemo" & @CRLF)
ControlClick($main, "", "TAdvGlowButton14")
WinWait($message, "Never ask again", 5)
ControlClick($message, "", "TAdvOfficeRadioButton2")
ControlClick($message, "", "TAdvGlowButton1")
#EndRegion --- DASHBOARD TEST ---

#Region --- ORDER TRAKKER TEST ---
_OpenApp("trakker")



#EndRegion --- ORDER TRAKKER TEST ---

Exit
