#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.

; Move the mouse to location (2263, 210)
CoordMode, Mouse, Screen
MouseMove, 2043, 66, 0

; Double click to select text
Click, left, 2043, 66, 2

; Copy the selected text using ctrl + c
Send, 398



