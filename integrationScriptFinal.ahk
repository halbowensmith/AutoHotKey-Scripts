#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

R1 := ""
R2 := ""
R3 := ""
R4 := ""
R5 := "" 
Avg := ""

setIntegrationRange(from, to)
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlSetText, Edit1, %from% ; Use the window found above.
ControlSetText, Edit2, %to% ; Use the window found above.
ControlClick, Button3, , , Left, 2
}


copyIntegration()
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlGetText, topIntegration, Edit3 ; Use the window found above.
Return topIntegration
}

moveToSpectra(spectra)
{
WinActivate, ahk_class Afx:400000:8:10005:0:0
WinWaitActive, ahk_class Afx:400000:8:10005:0:0
Switch spectra
{
Case "R1": ControlClick, Button5, , , Left, 2
Case "R2": ControlClick, Button6, , , Left, 2
Case "R3": ControlClick, Button7, , , Left, 2
Case "R4": ControlClick, Button8, , , Left, 2
Case "R5": ControlClick, Button9, , , Left, 2
Case "Avg": ControlClick, Button4, , , Left, 2
Default: MsgBox, Code isn't working

}
}


Gui, New , , Integration Extractor
Gui, Add, Text,, Select excel file
Gui, Add, Edit,vFilePath,
Gui, Add, Button,Default gFindFilePath ,Open
Gui, Add, Text,, Select sensor number and blank or BT then press Process
Gui, Add, DropDownList, vSensorNumber, 1PP|2PP|3PP|4PP|5PP|6PP|7PP|8PP|9PP|10PP|11PP|12PP|13PP|14PP|15PP|16PP|17PP|18PP|19PP|20PP
GUI, Add, Radio,vBlankSensor, Blank
Gui, Add, Radio,vBTSensor, BT
Gui, Add, Button, Default gProcessSpectra, Process Spectra
Gui, Show, Center
Return

FindFilePath:
{
FileSelectFile, SelectedFile, 3, , Open a file, Excel Workbook (*.xlsx; *.xls)
if (SelectedFile = "")
    MsgBox, The user didn't select anything.
else
    GuiControl,, FilePath, %SelectedFile%
}
return

ProcessSpectra:
{
setIntegrationRange(398,435)

moveToSpectra("R1")
R1 := copyIntegration()
moveToSpectra("R2")
R2 := copyIntegration()
moveToSpectra("R3")
R3 := copyIntegration()
moveToSpectra("R4")
R4 := copyIntegration()
moveToSpectra("R5")
R5 := copyIntegration()
moveToSpectra("Avg")
Avg := copyIntegration()


Gui, Submit, Nohide
oWorkBook := ComObjGet(FilePath)		; get reference to WorkBook
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1	; just do it - too long to explain why...

If (BlankSensor = 1)
{
Switch SensorNumber
{
Case "1PP": oWorkbook.Worksheets("398-435cm-1").Range("C4").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
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
} Else if(BTSensor = 1)
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
}
}
else
{
MsgBox, Neither option was selected
}
oWorkbook.Close(1)
} 
return

GuiClose:
	ExitApp
return