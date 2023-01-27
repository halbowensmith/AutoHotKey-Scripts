#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

~[::
Input, UserInput, V T5 L4 C, {enter}.{esc}{tab}, btw,otoh,fl,ahk,ca
switch ErrorLevel
{
case "Max":
    MsgBox, You entered "%UserInput%", which is the maximum length of text.
    return
case "Timeout":
    MsgBox, You entered "%UserInput%" at which time the input timed out.
    return
case "NewInput":
    return
default:
    if InStr(ErrorLevel, "EndKey:")
    {
        MsgBox, You entered "%UserInput%" and terminated the input with %ErrorLevel%.
        return
    }
}
switch UserInput
{
case "btw":   Send, {backspace 4}by the way
case "otoh":  Send, {backspace 5}on the other hand
case "fl":    Send, {backspace 3}Florida
case "ca":    Send, {backspace 3}California
case "ahk":   Run, https://www.autohotkey.com
}
return