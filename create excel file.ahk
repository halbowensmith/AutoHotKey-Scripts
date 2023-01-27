#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Add 				        ;add a new workbook w/ standard 3 sheets
xl.range("a1:a1").value := "Test"
xL.ActiveWorkbook.SaveAs("testXLfile",51)               ;51 is an xlsx, 56 is an xls
xl.WorkBooks.Close()                                    ;close file
xl.quit