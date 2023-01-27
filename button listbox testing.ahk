#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Add, Button, Default w80 gsub1, just text
Gui, Add, Button, x59 y345 h135 w99 gsub2, test 2(3545)
Gui, Add, DropDownList, vColorChoice, Black|White|Red|Green|Blue
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