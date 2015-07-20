#comments-start
--- NOTE ---
This is a temporary .au3 script used for testing small scripts intended to be
integrated into larger scripts. Everything in here will come and go, so don't
put anything you really like here, unless you put it somewhere else afterward
#comments-end

#include <Header.au3>

_OpenWS()

_OpenApp("term")

;ControlSend($main, "Ready", "AfxFrameOrView801", "{Enter 4}")
;ControlSend($main, "Ready", "AfxFrameOrView801", $userPwd)
;ControlSend($main, "Ready", "AfxFrameOrView801", "{Enter}")
ControlFocus($main, "Ready", "AfxFrameOrView801")
Send("{Enter 4}")
Sleep(200)
Send($userPwd & "{Enter}")
;Send("{Enter}")


