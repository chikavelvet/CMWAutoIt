#include <Header.au3>

WinActivate($main)

Local $payment = "Cash"

If $payment = "Cash"
	Local $amt = ControlGetText($main, "", "TAdvEdit3")

Else
	If ControlCommand($main, "", "TAdvComboBox1", "FindString", $payment) <> 0 Then
		ControlCommand($main, "", "TAdvComboBox1", "SelectString", $payment)
	Else
		ConsoleWrite("Cash Not Found" & @CRLF)
	EndIf
EndIf