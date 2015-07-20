#comments-start
--- NOTES ---
This is the au3 to run for simply opening CMW. The actual function is in
Header.au3, so as long as you have that included in a script, you can call
it from anywhere, this is just a run-file.

#comments-end

#include <Header.au3>

#Region --- Main Execution Section ---
_OpenWS()

_endProg()
