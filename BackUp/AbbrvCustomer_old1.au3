#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---
#include <Header.au3>

_OpenWS("C:\Users\Paul\Desktop\AutoIT Stuff\Login.csv")
;--- Custom Script Code here for different test plans ---
; Creating abbreviated customer screen

Local $customerFile = FileOpen("C:\Users\Paul\Desktop\AutoIT Stuff\Customer.csv")
Local $customerInfoAllLines = FileReadToArray($customerFile)

For $i = 0 to UBound($customerInfoAllLines)-1
	Local $customerInfo = StringSplit($customerInfoAllLines[$i], ",")

	ControlClick("[CLASS:TfrmMain_CMW]","","TEditButton1")
	WinWait($cuswind,"")
	ConsoleWrite($customerInfo[0] & @CRLF)
	ControlSetText($cuswind,"","TAdvEdit1",$customerInfo[1])
	ControlSetText($cuswind,"","TAdvEdit19",$customerInfo[2])
	ControlSetText($cuswind,"","TAdvEdit17",$customerInfo[3])
	ControlSetText($cuswind,"","TAdvEdit16",$customerInfo[4])
	ControlSetText($cuswind,"","TAdvEdit11",$customerInfo[5])
	ControlSetText($cuswind,"","TAdvEdit14",$customerInfo[6])
	ControlSetText($cuswind,"","TAdvEdit13",$customerInfo[7])
	ControlSetText($cuswind,"","TAdvEdit10",$customerInfo[8])
	ControlSetText($cuswind,"","TAdvEdit5",$customerInfo[9])
	ControlSetText($cuswind,"","TAdvEdit15",$customerInfo[10])

	ControlFocus($cuswind,"","TAdvSmoothTabPager1")
	Send("{Right}")
	ControlClick($cuswind,"","TAdvGlowButton1")
	Sleep(50)
	ControlFocus($cuswind,"","TAdvSmoothTabPager1")
	Send("{Right}")
	Exit
	ControlSetText($cuswind,"","TAdvEdit4",$customerInfo[11])
	ControlSetText($cuswind,"","TAdvEdit6",$customerInfo[12])
	ControlSetText($cuswind,"","TCMWEditDate1",$customerInfo[13])
	ControlClick($cuswind, "", "TAdvComboBox6", "primary")
	ControlSetText($cuswind, "", "TAdvComboBox6", $userName)
	;ControlSetText($cuswind,"","TAdvComboBox5","A")
	ControlSetText($cuswind,"","TAdvComboBox2",$customerInfo[14])
	ControlClick($cuswind,"","TAdvSmoothTabPager1")
	Send("{Right}")
	ControlClick($cuswind,"","TAdvGlowButton6")
	_logIt("Creating Abbrv Customer.")

	ControlClick($main, "", "TAdvGlassButton1")
Next

Exit

; looking up part
WinWait("[CLASS:TfrmMain_CMW]","frm_FindandSell_Main")
send("{Enter}")
_sendInput("[CLASS:TfrmMain_CMW]","TAdvEdit2","00,TL,ENG{Enter}")
WinWait("[CLASS:TfrmMain_CMW]","frm_FindAndSell_InterchangeSelection")
Send("{ENTER}")
WinWait("[CLASS:TfrmMain_CMW]","Start New Work Order")
_sendInput("[CLASS:TfrmMain_CMW]","TAdvEdit1","19{Enter}")
if WinWait("[CLASS:TfrmWiqConflict]","Make Available",2) Then
	ControlClick("[CLASS:TfrmWiqConflict]","Make Available","TButton4")
	WinWait("[CLASS:Tfrm_Reasons]","OK")
	send ("{Down}{Down}")
	ControlClick("[CLASS:Tfrm_Reasons]","OK","TBitBtn2")
EndIf
;CLICK THE EDIT BUTTON
ControlClick("[CLASS:TfrmMain_CMW]","Start New Work Order","TAdvGlowButton16")
_WinWait("[CLASS:TfrmInventoryUpdate]","Accept","Part Edit Screen")
_sendInput("[CLASS:TfrmInventoryUpdate]","TAdvEdit8","STOCK{Enter}")
;_sendInput("[CLASS:TfrmInventoryUpdate]","Edit1","CAMRY{Enter}")
_sendInput("[CLASS:TfrmInventoryUpdate]","TAdvEdit9","LOCAT{Enter}")
_sendInput("[CLASS:TfrmInventoryUpdate]","TJvCaptionPanel2","549685{Enter}")
_sendInput("[CLASS:TfrmInventoryUpdate]","TAdvComboBox1","U{Enter}")
;_sendInput("[CLASS:TfrmInventoryUpdate]","TRadioGroup1","{Space}")
_sendInput("[CLASS:TfrmInventoryUpdate]","TMemo1","THIS IS MY DESCRIPTION{Enter}")
ControlClick("[CLASS:TfrmInventoryUpdate]","Accept","TAdvGlowButton3")
;_WinWait("[CLASS:#32770]","You have not selected an Interchange for this Part","Selected no ic")
;ControlClick("[CLASS:#32770]","Yes","Button1")
; done with edit start new wo
ControlClick("[CLASS:TfrmMain_CMW]","Start New Work Order","TAdvGlowButton19")
_WinWait("[CLASS:TfrmMain_CMW]","Print Work Order","Create Work Order Screen")
ControlFocus("[CLASS:TfrmMain_CMW]","","TAdvGlowButton17")
ControlClick("[CLASS:TfrmMain_CMW]","Print Work Order","TAdvGlowButton17")
_logIt("Print WorkOrder Button Pressed")
_WinWait("[CLASS:TfrmMain_CMW]","frm_FindandSell_Main","Work Order Printed. Find-N-Sell screen")

;back at fns looking up same part
send("{Tab}")
_sendInput("[CLASS:TfrmMain_CMW]","TAdvEdit2","00,TL,ENG{Enter}")
_WinWait("[CLASS:TfrmMain_CMW]","frm_FindAndSell_InterchangeSelection","Interchange Selection Screen")
Send("{ENTER}")
_WinWait("[CLASS:TfrmMain_CMW]","Start New Work Order","Inventory Search Results")
_sendInput("[CLASS:TfrmMain_CMW]","TAdvEdit1","19{Enter}")
if WinWait("[CLASS:TfrmWiqConflict]","Make Available",2) Then
	ControlClick("[CLASS:TfrmWiqConflict]","Make Available","TButton4")
	_WinWait("[CLASS:Tfrm_Reasons]","OK","Choose a Reason")
	send ("{Down}")
	ControlClick("[CLASS:Tfrm_Reasons]","OK","TBitBtn2")
EndIf

_WinWait("[CLASS:TfrmMain_CMW]","Start New Work Order","Inventory Search Results")
ControlClick("[CLASS:TfrmMain_CMW]","Start New Work Order","TAdvGlowButton19")
_WinWait("[CLASS:TfrmMain_CMW]","Print Work Order","Create Work Order Screen")
ControlFocus("[CLASS:TfrmMain_CMW]","","TAdvEditBtn2")
_sendInput("[CLASS:frm_findandsell_WorkOrder]","TAdvEditBtn2","21{Enter}")
_WinWait("[CLASS:TfrmFind_Customer]","Select Customer","Customer Selection Screen")
send("{Enter}")
ControlFocus("[CLASS:TfrmMain_CMW]","","TAdvGlowButton17")
ControlClick("[CLASS:TfrmMain_CMW]","Print Work Order","TAdvGlowButton17")
_logIt("Print WorkOrder Button Pressed")
_WinWait("[CLASS:TfrmMain_CMW]","frm_FindandSell_Main","Work Order Printed. Find-N-Sell screen")
;;--------------------------------------------------------
_endProg()

#endregion --- Main Execution Section End ---
