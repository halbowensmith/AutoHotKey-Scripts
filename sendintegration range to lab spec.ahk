#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlSetText, Edit1, 398 ; Use the window found above.
ControlSetText, Edit2, 435 ; Use the window found above.
ControlClick, Button3, , , Left, 2