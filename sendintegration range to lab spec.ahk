#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

setIntegrationRange(from, to)
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlSetText, Edit1, %from% ; Use the window found above.
ControlSetText, Edit2, %to% ; Use the window found above.
ControlClick, Button3, , , Left, 2
}


copyIntegration()
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlGetText, topIntegration, Edit3 ; Use the window found above.
Return topIntegration
}

moveToSpectra()
{

}


Gui, New , , Gui Testing
Gui, Add, Text,, Select excel file
Gui, Add, Text,, Select sensor number and blank or BT, then press Process
GUI, Add, Radio,, Blank
Gui, Add, Radio,, BT
Gui, Add, Button, Default gsub1, Process Spectra
Gui, Show

sub1:
Gui, submit, nohide
{
switch Name
{
case "btw":   MsgBox, success ;Send, {backspace 4}by the way
case "otoh":  Send, {backspace 5}on the other hand
case "fl":    Send, {backspace 3}Florida
case "ca":    Send, {backspace 3}California
case "ahk":   Run, https://www.autohotkey.com
}
;return
}