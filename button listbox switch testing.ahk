#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Add, Button, Default gsub1, Process Spectra
Gui, Add, DropDownList, vSensorNumber, 1PP|2PP|3PP|4PP|5PP|6PP|7PP|8PP|9PP|10PP|11PP|12PP|13PP|14PP|15PP|16PP|17PP|18PP|19PP|20PP
gui, show, center
return

sub1:
Gui, Submit, Nohide
{
Switch SensorNumber
{
Case "1PP":MsgBox, 1 print pass selected
Case "2PP": MsgBox, 2 print pass selected
Case "3PP": MsgBox, 2 print pass selected
Case "4PP": MsgBox, 2 print pass selected
Case "5PP": MsgBox, 2 print pass selected
Case "6PP": MsgBox, 2 print pass selected
Case "7PP": MsgBox, 2 print pass selected
Case "8PP": MsgBox, 2 print pass selected
Case "9PP": MsgBox, 2 print pass selected
Case "10PP": MsgBox, 2 print pass selected
Case "11PP": MsgBox, 2 print pass selected
Case "12PP": MsgBox, 2 print pass selected
Case "13PP": MsgBox, 2 print pass selected
Case "14PP": MsgBox, 2 print pass selected
Case "15PP": MsgBox, 2 print pass selected
Case "16PP": MsgBox, 2 print pass selected
Case "17PP": MsgBox, 2 print pass selected
Case "18PP": MsgBox, 2 print pass selected
Case "19PP": MsgBox, 2 print pass selected
Case "20PP": MsgBox, 2 print pass selected


Default:
MsgBox, No Sensor Selected/Code is broken
return
}








}
return
