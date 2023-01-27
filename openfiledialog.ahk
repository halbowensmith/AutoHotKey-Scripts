#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;FileSelectFile, SelectedFile, 3, , Open a file, Text Documents (*.txt; *.doc)
FileSelectFile, SelectedFile, 3, , Open a file, Excel Workbook (*.xlsx; *.xls)
if (SelectedFile = "")
    MsgBox, The user didn't select anything.
else
    MsgBox, The user selected the following:`n%SelectedFile%