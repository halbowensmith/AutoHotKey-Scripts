#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.

R1 := ""
R2 := ""
R3 := ""
R4 := ""
R5 := "" 
avg := ""

moveAndClick(x,y)
{
; Move the mouse to location (x, y)
CoordMode, Mouse, Screen
MouseMove, x, y, 0
; Double click to select text
Click, left, x, y, 2
}

moveClickType(x,y,value)
{
; Move the mouse to location (x, y)
CoordMode, Mouse, Screen
MouseMove, x, y, 0
; Double click to select text
Click, left, x, y, 2
; Type value
Send, %value%
}

setIntegrationRange(from, to)
{
; Move the mouse to location (x, y)
CoordMode, Mouse, Screen
MouseMove, 2043, 66, 0
; Double click to select text
Click, left, 2043, 66, 2
; Fill in from value
Send, %from%

; Move the mouse to location (x, y)
CoordMode, Mouse, Screen
MouseMove, 2060, 87, 0
; Double click to select text
Click, left, 2060, 87, 2
; Fill in from value
Send, %to%
}

copyIntegration(x,y)
{
;select spectra from graph list at right hand side of screen
moveAndClick(x,y)
;select top value
moveAndClick(2056,113)
;ctrl c doesn't copy from labspec integration window Send, ^c
;have to right click and copy
;right click on top box
Click, right, 2056, 113
;click copy button
moveAndClick(2096,169)
;save clipboard value to variable R
;R := Clipboard
Return Clipboard
}


;type bottom range
;moveClickType(2043,66,398)

setIntegrationRange(398,435)

R1 := copyIntegration(3831,197)
R2 := copyIntegration(3831,216)
R3 := copyIntegration(3831,232)
R4 := copyIntegration(3831,248)
R5 := copyIntegration(3831,270)
avg := copyIntegration(3831,178)

; Switch to Excel
WinActivate, ahk_class XLMAIN

;wait until the excel window is active
WinWaitActive, ahk_class XLMAIN

; Paste the value of the variable R1
SendInput, %R1%
SendInput, {Enter}
SendInput, %R2%
SendInput, {Enter}
SendInput, %R3%
SendInput, {Enter}
SendInput, %R4%
SendInput, {Enter}
SendInput, %R5%
SendInput, {Enter}
SendInput, %avg%



