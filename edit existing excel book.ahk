#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.







FullPath := "C:\Users\Smithha\Desktop\21Dec22 Ag Nanocomposix 1-20PP in BT(2).xlsx"	; please adjust full path to your Workbook...
oWorkBook := ComObjGet(FullPath)		; get reference to WorkBook
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1	; just do it - too long to explain why...

oWorkbook.Worksheets("398-435cm-1").Range("C4").Value := "Test"	; set "Test" in "Sheet1" sheet, "A2" cell
;oSheet := oWorkbook.Worksheets.Add()		; add a new Sheet
;oSheet.Name := "Data1"	; name new sheet to "Data1"
;oWorkbook.Worksheets("Data1").Range("A2").Value := "Test 2"	; set "Test 2" in "Data1" sheet, "A2" cell
oWorkbook.Close(1)	; save changes and close Workbook
oWorkBook := "", oSheet := ""	; release references
ExitApp