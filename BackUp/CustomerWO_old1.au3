#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---
#include <Header.au3>

_OpenWS("C:\Users\Paul\Desktop\AutoIT Stuff\Login.csv")
;--- Custom Script Code here for different test plans ---

;Local $main = "[CLASS:TfrmMain_CMW]"


ControlSetText($main,"","TAdvEditBtn1","123")
_sendinput($main, "TAdvEditBtn1", "{Enter}")
WinWait($find,"Select Customer")
ControlFocus($find, "", "TAdvStringGrid2")
ControlClick($find, "Select Customer", "TAdvGlowButton3")
WinWait($main,"frm_FindandSell_Main")
ControlSetText($main,"","TAdvEdit2","96,METRO,ENG")
ControlClick($main, "FIN&D", "TAdvGlassButton2")
WinWait($main,"frm_FindAndSell_InterchangeSelection")

ControlTreeView($main, "Smart Vin", "THTMLTreeList1", "Select", "#0|#0")
Send("{SPACE}")
ControlClick($main, "", "TAdvGlassButton1")
WinWait($main,"Start New Work Order")

ControlSetText($main, "Start New Invoice", "TAdvEdit1", "1")
Send("{Enter}")

WinWait($extra,"Create Extra Sale")

ControlSetText($extra, "", "TAdvEdit1", "TEST STOCK")
ControlSetText($extra, "", "TAdvEdit2", "YARD")
ControlSetText($extra, "", "TAdvMoneyEdit2", "10000")
ControlClick($extra,"Create Extra Sale","TButton1")
WinWait($extra, "")
ControlClick($extra, "Cancel", "TButton3")
WinWait($main,"Start New Work Order")


ControlClick($main,"Start New Work Order","TAdvGlowButton21")
_logIt("Creating Work Order Button Pressed.")
;ControlClick($main, "", "TAdvComboBox1")
Local $payment = "Charge"
If ControlCommand($main, "", "TAdvComboBox1", "FindString", $payment) <> 0 Then
	ControlCommand($main, "","TAdvComboBox1", "SelectString", $payment)
Else
	ConsoleWrite("Cash Not Found" & @CRLF)
EndIf
Exit
WinWait($main, "Cash")
ControlClick($main,"Print Work Order","TAdvGlowButton17")
_logIt("Print WorkOrder Button Pressed")
WinWait($main,"frm_FindandSell_Main")
;--------------------------------------------------------
;_endProg()
Exit
#endregion --- Main Execution Section End ---
