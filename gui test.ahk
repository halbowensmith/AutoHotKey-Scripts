#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Persistent

Gui, New, -MinimizeBox, % "My private search engine"
Gui, Add, Text, , % "Specify a search string:"
Gui, Add, Edit, vUserInput w300
Gui, Add, Checkbox, vMyCheckbox, % "Check Me!"
Gui, Show, Center
return

OkButtonPress:
	GUI, Submit
	Msgbox, %vUserInput%
return

GuiClose:
	ExitApp
return