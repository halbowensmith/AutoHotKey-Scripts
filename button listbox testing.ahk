#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Add, Button, Default gsub1, Process Spectra
Gui, Add, DropDownList, vColorChoice, 1PP|2PP|3PP|4PP|5PP|6PP|7PP|8PP|9PP|10PP|11PP|12PP|13PP|14PP|15PP|16PP|17PP|18PP|19PP|20PP
gui, show, center
return

sub1:
{
MsgBox, [ Options, OutputVar, buttton, 10000]
}
return

sub2:
{
Msgbox, second button testing complete
}
return