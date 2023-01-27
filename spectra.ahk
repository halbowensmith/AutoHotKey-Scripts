ahk_id := "14408"
classnn := "Edit3"

ControlGetText, OutputVar,, %ahk_id%, %classnn%
FileAppend, %OutputVar%, C:\Users\Smithha\Documents\AutoHotkey\topvalue.txt
MsgBox, [ Options, OutputVar, %OutputVar%, 10000]