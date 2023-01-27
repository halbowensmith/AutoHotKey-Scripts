#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;WinActivate, ahk_class Afx:400000:8:10005:0:0
;WinWaitActive, ahk_class Afx:400000:8:10005:0:0
;ControlClick, Button9, , , Left, 2

moveToSpectra(spectra)
{
WinActivate, ahk_class Afx:400000:8:10005:0:0
WinWaitActive, ahk_class Afx:400000:8:10005:0:0
Switch spectra
{
Case "R1": ControlClick, Button5, , , Left, 2
Case "R2": ControlClick, Button6, , , Left, 3
Case "R3": ControlClick, Button7, , , Left, 2
Case "R4": ControlClick, Button8, , , Left, 2
Case "R5": ControlClick, Button9, , , Left, 2
Case "Avg": ControlClick, Button4, , , Left, 2
Default: MsgBox, Code isn't working

}
}


moveToSpectra("Avg")