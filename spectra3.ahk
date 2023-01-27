ahk_id := "133818"
classnn := "Edit3"
WinGet, hwnd, ID, %ahk_id%, %classnn%
DllCall("SendMessage", "UInt", hwnd, "UInt", 0xD, "UInt", 0, "Str", OutputVar, "UInt", 32767)
FileAppend, %OutputVar%, C:\Users\Smithha\Documents\AutoHotkey\topvalue.txt
MsgBox, [ Options, OutputVar, %OutputVar%, 10000]