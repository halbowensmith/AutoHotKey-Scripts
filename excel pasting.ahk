#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; Switch to Excel
WinActivate, ahk_class XLMAIN

; Paste the copied text into cell A1

Send, {Ctrl down}{g}{Ctrl Up}{C5}{Enter}