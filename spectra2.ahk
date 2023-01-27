ahk_id := "133818"
classnn := "Edit3"
ControlGet, OutputVar, Edit3, %ahk_id%
FileAppend, %OutputVar%, C:\Users\Smithha\Documents\AutoHotkey\topvalue.txt
MsgBox, [ Options, OutputVar, %OutputVar%, 10000]