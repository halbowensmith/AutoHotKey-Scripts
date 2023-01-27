#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


InputBox, Sensor#, Sensor Number, Please enter the sensor number (will correspond to position in excel sheet)
Gui, Add, Checkbox, vMyCheckbox, % "Check Me!"