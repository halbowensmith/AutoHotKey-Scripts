#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.

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

;type bottom range
;moveClickType(2043,66,398)

setIntegrationRange(398,435)



