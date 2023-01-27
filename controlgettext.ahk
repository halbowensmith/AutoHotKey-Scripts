WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlGetText, OutputVar, Edit3 ; Use the window found above.
; Switch to Excel
WinActivate, ahk_class XLMAIN

;wait until the excel window is active
WinWaitActive, ahk_class XLMAIN

; Paste the value of the variable R1
SendInput, %OutputVar%