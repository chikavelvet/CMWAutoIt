#comments-start

--- NOTES ---
This is a script intended to create customers from a CSV (comma-delimited list) file


#comments-end

#include <Header.au3>

_OpenWS()

Local $fiCustomerFile = FileOpen("C:\Users\Paul\Desktop\AutoIT Stuff\Customer.csv")
Local $asCustomerInfoAllLines = FileReadToArray($fiCustomerFile)

For $i = 0 To UBound($asCustomerInfoAllLines) - 1
	Local $asCustomerInfo = StringSplit($asCustomerInfoAllLines[$i], ",")

	ControlClick("[CLASS:TfrmMain_CMW]", "", "TEditButton1")
	WinWait($g_wCustomer, "")
	ConsoleWrite($asCustomerInfo[0] & @CRLF)
	ControlSetText($g_wCustomer, "", "TAdvEdit1", $asCustomerInfo[1])
	ControlSetText($g_wCustomer, "", "TAdvEdit19", $asCustomerInfo[2])
	ControlSetText($g_wCustomer, "", "TAdvEdit17", $asCustomerInfo[3])
	ControlSetText($g_wCustomer, "", "TAdvEdit16", $asCustomerInfo[4])
	ControlSetText($g_wCustomer, "", "TAdvEdit11", $asCustomerInfo[5])
	ControlSetText($g_wCustomer, "", "TAdvEdit14", $asCustomerInfo[6])
	ControlSetText($g_wCustomer, "", "TAdvEdit13", $asCustomerInfo[7])
	ControlSetText($g_wCustomer, "", "TAdvEdit10", $asCustomerInfo[8])
	ControlSetText($g_wCustomer, "", "TAdvEdit5", $asCustomerInfo[9])
	ControlSetText($g_wCustomer, "", "TAdvEdit15", $asCustomerInfo[10])

	ControlFocus($g_wCustomer, "", "TAdvSmoothTabPager1")
	Send("{Right}")
	ControlClick($g_wCustomer, "", "TAdvGlowButton1")
	Sleep(50)
	ControlFocus($g_wCustomer, "", "TAdvSmoothTabPager1")
	Send("{Right}")
	_EndProg()
	ControlSetText($g_wCustomer, "", "TAdvEdit4", $asCustomerInfo[11])
	ControlSetText($g_wCustomer, "", "TAdvEdit6", $asCustomerInfo[12])
	ControlSetText($g_wCustomer, "", "TCMWEditDate1", $asCustomerInfo[13])
	ControlClick($g_wCustomer, "", "TAdvComboBox6", "primary")
	ControlSetText($g_wCustomer, "", "TAdvComboBox6", $g_sUserName)
	;ControlSetText($g_wCustomer,"","TAdvComboBox5","A")
	ControlSetText($g_wCustomer, "", "TAdvComboBox2", $asCustomerInfo[14])
	ControlClick($g_wCustomer, "", "TAdvSmoothTabPager1")
	Send("{Right}")
	ControlClick($g_wCustomer, "", "TAdvGlowButton6")
	ControlClick($g_wMain, "", "TAdvGlassButton1")
Next

_EndProg()
#comments-start
	; looking up part
	WinWait("[CLASS:TfrmMain_CMW]", "frm_FindandSell_Main")
	;Send("{Enter}")
	ControlSetText("[CLASS:TfrmMain_CMW]", "", "TAdvEdit2", "00,TL,ENG")
	ControlSend($g_wMain, "", "TAdvEdit2, "{Enter}")
	WinWait("[CLASS:TfrmMain_CMW]", "frm_FindAndSell_InterchangeSelection")
	Send("{ENTER}")
	WinWait("[CLASS:TfrmMain_CMW]", "Start New Work Order")
	_sendInput("[CLASS:TfrmMain_CMW]", "TAdvEdit1", "19{Enter}")
	If WinWait("[CLASS:TfrmWiqConflict]", "Make Available", 2) Then
	ControlClick("[CLASS:TfrmWiqConflict]", "Make Available", "TButton4")
	WinWait("[CLASS:Tfrm_Reasons]", "OK")
	Send("{Down}{Down}")
	ControlClick("[CLASS:Tfrm_Reasons]", "OK", "TBitBtn2")
	EndIf
	;CLICK THE EDIT BUTTON
	ControlClick("[CLASS:TfrmMain_CMW]", "Start New Work Order", "TAdvGlowButton16")
	_WinWait("[CLASS:TfrmInventoryUpdate]", "Accept", "Part Edit Screen")
	_sendInput("[CLASS:TfrmInventoryUpdate]", "TAdvEdit8", "STOCK{Enter}")
	;_sendInput("[CLASS:TfrmInventoryUpdate]","Edit1","CAMRY{Enter}")
	_sendInput("[CLASS:TfrmInventoryUpdate]", "TAdvEdit9", "LOCAT{Enter}")
	_sendInput("[CLASS:TfrmInventoryUpdate]", "TJvCaptionPanel2", "549685{Enter}")
	_sendInput("[CLASS:TfrmInventoryUpdate]", "TAdvComboBox1", "U{Enter}")
	;_sendInput("[CLASS:TfrmInventoryUpdate]","TRadioGroup1","{Space}")
	_sendInput("[CLASS:TfrmInventoryUpdate]", "TMemo1", "THIS IS MY DESCRIPTION{Enter}")
	ControlClick("[CLASS:TfrmInventoryUpdate]", "Accept", "TAdvGlowButton3")
	;_WinWait("[CLASS:#32770]","You have not selected an Interchange for this Part","Selected no ic")
	;ControlClick("[CLASS:#32770]","Yes","Button1")
	; done with edit start new wo
	ControlClick("[CLASS:TfrmMain_CMW]", "Start New Work Order", "TAdvGlowButton19")
	_WinWait("[CLASS:TfrmMain_CMW]", "Print Work Order", "Create Work Order Screen")
	ControlFocus("[CLASS:TfrmMain_CMW]", "", "TAdvGlowButton17")
	ControlClick("[CLASS:TfrmMain_CMW]", "Print Work Order", "TAdvGlowButton17")
	_logIt("Print WorkOrder Button Pressed")
	_WinWait("[CLASS:TfrmMain_CMW]", "frm_FindandSell_Main", "Work Order Printed. Find-N-Sell screen")

	;back at fns looking up same part
	Send("{Tab}")
	_sendInput("[CLASS:TfrmMain_CMW]", "TAdvEdit2", "00,TL,ENG{Enter}")
	_WinWait("[CLASS:TfrmMain_CMW]", "frm_FindAndSell_InterchangeSelection", "Interchange Selection Screen")
	Send("{ENTER}")
	_WinWait("[CLASS:TfrmMain_CMW]", "Start New Work Order", "Inventory Search Results")
	_sendInput("[CLASS:TfrmMain_CMW]", "TAdvEdit1", "19{Enter}")
	If WinWait("[CLASS:TfrmWiqConflict]", "Make Available", 2) Then
	ControlClick("[CLASS:TfrmWiqConflict]", "Make Available", "TButton4")
	_WinWait("[CLASS:Tfrm_Reasons]", "OK", "Choose a Reason")
	Send("{Down}")
	ControlClick("[CLASS:Tfrm_Reasons]", "OK", "TBitBtn2")
	EndIf

	_WinWait("[CLASS:TfrmMain_CMW]", "Start New Work Order", "Inventory Search Results")
	ControlClick("[CLASS:TfrmMain_CMW]", "Start New Work Order", "TAdvGlowButton19")
	_WinWait("[CLASS:TfrmMain_CMW]", "Print Work Order", "Create Work Order Screen")
	ControlFocus("[CLASS:TfrmMain_CMW]", "", "TAdvEditBtn2")
	_sendInput("[CLASS:frm_findandsell_WorkOrder]", "TAdvEditBtn2", "21{Enter}")
	_WinWait("[CLASS:TfrmFind_Customer]", "Select Customer", "Customer Selection Screen")
	Send("{Enter}")
	ControlFocus("[CLASS:TfrmMain_CMW]", "", "TAdvGlowButton17")
	ControlClick("[CLASS:TfrmMain_CMW]", "Print Work Order", "TAdvGlowButton17")
	_logIt("Print WorkOrder Button Pressed")
	_WinWait("[CLASS:TfrmMain_CMW]", "frm_FindandSell_Main", "Work Order Printed. Find-N-Sell screen")
	;;--------------------------------------------------------

#comments-end
_EndProg()

