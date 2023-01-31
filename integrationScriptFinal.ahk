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
R1Class := ""
R2Class := ""
R3Class := ""
R4Class := ""
R5Class := ""
AvgClass := ""
LabSpecClass := ""




Gui, New , , Integration Extractor
Gui, Add, Text,, Select excel file
Gui, Add, Edit,vFilePath,
Gui, Add, Button,Default gFindFilePath ,Open
Gui, Add, Text,, Select sensor number and blank or BT then press Process
Gui, Add, DropDownList, vSensorNumber, 1PP|2PP|3PP|4PP|5PP|6PP|7PP|8PP|9PP|10PP|11PP|12PP|13PP|14PP|15PP|16PP|17PP|18PP|19PP|20PP
GUI, Add, Radio,vBlankSensor, Blank
Gui, Add, Radio,vBTSensor, BT
Gui, Add, Button, Default gProcessSpectra, Process Spectra
Gui, Add, Button, Default gcalibrate, Calibrate
Gui, Add, Button, Default gtesting, testingp passthrough
Gui, Add, Button, Default gtestPress, test press
Gui, Show, Center
Return

setIntegrationRange(from, to)
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlSetText, Edit1, %from% ; Use the window found above.
ControlSetText, Edit2, %to% ; Use the window found above.
ControlClick, Button3, , , Left, 2
return
}


copyIntegration()
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlGetText, topIntegration, Edit3 ; Use the window found above.
Return topIntegration
}




testPress:
;WinActivate, ahk_class %LabSpecClass% 
;msgbox, %LabSpecClass%
;WinWaitActive, ahk_class %LabSpecClass%
;ControlClick, %R1Class%, , , Left, 2
moveToSpectra("R1")
Return

moveToSpectra(spectra)
{
global R1Class
global R2Class 
global R3Class 
global R4Class 
global R5Class 
global AvgClass 
global LabSpecClass
WinActivate, ahk_class %LabSpecClass% 
WinWaitActive, ahk_class %LabSpecClass%
Switch spectra
{
Case "R1": ControlClick, %R1Class%, , , Left, 2
Case "R2": ControlClick, %R2Class%, , , Left, 2
Case "R3": ControlClick, %R3Class%, , , Left, 2
Case "R4": ControlClick, %R4Class%, , , Left, 2
Case "R5": ControlClick, %R5Class%, , , Left, 2
Case "Avg": ControlClick, %AvgClass%, , , Left, 2
Default: MsgBox, Code isn't working
}
}



testing:
Msgbox, R1 %R1Class% `n R2 %R2Class% `n R3 %R3Class% `n R4 %R4Class% `n R5 %R5Class% `n Avg %AvgClass% `n ahk %LabSpecClass% 
return

FindFilePath:
{
FileSelectFile, SelectedFile, 3, , Open a file, Excel Workbook (*.xlsx; *.xls)
if (SelectedFile = "")
    MsgBox, The user didn't select anything.
else
    GuiControl,, FilePath, %SelectedFile%
}
return


calibrate:
Gui, CalibrateWindow:New , , Calibration Window
Gui, Add, Text, ,Welcome to the calibration process
Gui, Add, Text, , Click the examine button, then the lab spec window to autofill its identifiers
Gui, Add, Text, ,For each Measured Spectra click the button then on the view spectra button on the far right side of lab spec
Gui, Add, Text, ,The filled box should then say Button#
Gui, Add, Text, ,Press save to save the parameters, then close to return to previous window
Gui, Add, Text, ,Do not press the X, that will close the program and you will have to start again

Gui, Add, Button, Default x12 y160  gFindAHKClass, Examine
Gui, Add, Edit, vAHKClassEdit x70 y160
Gui, Add, Button, Default x12 y200 gR1ClassFind, R1 Class
Gui, Add, Edit, vR1ControlEdit x70 y200
Gui, Add, Button, Default x12 y240 gR2ClassFind, R2 Class
Gui, Add, Edit, vR2ControlEdit x70 y240
Gui, Add, Button, Default x12 y280 gR3ClassFind, R3 Class
Gui, Add, Edit, vR3ControlEdit x70 y280
Gui, Add, Button, Default x12 y320 gR4ClassFind, R4 Class
Gui, Add, Edit, vR4ControlEdit x70 y320
Gui, Add, Button, Default x12 y360 gR5ClassFind, R5 Class
Gui, Add, Edit, vR5ControlEdit x70 y360
Gui, Add, Button, Default x12 y400 gAvgClassFind, Avg Class
Gui, Add, Edit, vAvgControlEdit x75 y400
Gui, Add, Button, Default x30 y480 gSaveClasses, Save Parameters
Gui, Add, Button, Default x140 y480 gCloseGUI, Close
gui, show, center
Return

findClassNN:
Hotkey, LButton, leftButtonCode, On
Return

R1ClassFind:
Hotkey, LButton, leftButtonCodeR1, On
Return

R2ClassFind:
Hotkey, LButton, leftButtonCodeR2, On
Return

R3ClassFind:
Hotkey, LButton, leftButtonCodeR3, On
Return

R4ClassFind:
Hotkey, LButton, leftButtonCodeR4, On
Return

R5ClassFind:
Hotkey, LButton, leftButtonCodeR5, On
Return

AvgClassFind:
Hotkey, LButton, leftButtonCodeAvg, On
Return

leftButtonCodeR1:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R1ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeR2:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R2ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeR3:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R3ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeR4:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R4ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeR5:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R5ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeAvg:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, AvgControlEdit, %OutputVarControl%
gui, submit, nohide
Return

leftButtonCodeAHKClass:
MouseGetPos , , , ID
WinGetClass, class, ahk_id %ID%
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, AHKClassEdit, %class%
gui, submit, nohide
Return

leftButtonCode:
MouseGetPos , , , ID
MouseGetPos , OutputVarX, OutputVarY, OutputVarWin, OutputVarControl
Hotkey, LButton, leftButtonCode, Off
GuiControl, CalibrateWindow:, R1ControlEdit, %OutputVarControl%
gui, submit, nohide
Return

SaveClasses:
Gui, submit, nohide
R1Class := R1ControlEdit
R2Class := R2ControlEdit
R3Class := R3ControlEdit
R4Class := R4ControlEdit
R5Class := R5ControlEdit
AvgClass := AvgControlEdit
LabSpecClass := AHKClassEdit
Msgbox, R1 %R1Class% `n R2 %R2Class% `n R3 %R3Class% `n R4 %R4Class% `n R5 %R5Class% `n Avg %AvgClass% `n ahk %LabSpecClass% 
;Gui, Destroy
Return

FindAHKClass:
Hotkey, LButton, leftButtonCodeAHKClass, On
Return

CloseGUI:
Gui, Destroy
Return

ProcessSpectra:
{
setIntegrationRange(398,435)

moveToSpectra("R1")
sleep, 200
R1 := copyIntegration()
moveToSpectra("R2")
sleep, 200
R2 := copyIntegration()
moveToSpectra("R3")
sleep, 200
R3 := copyIntegration()
moveToSpectra("R4")
sleep, 200
R4 := copyIntegration()
moveToSpectra("R5")
sleep, 200
R5 := copyIntegration()
moveToSpectra("Avg")
sleep, 200
Avg := copyIntegration()


Gui, Submit, Nohide
oWorkBook := ComObjGet(FilePath)		; get reference to WorkBook
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1	; just do it - too long to explain why...

If (BlankSensor = 1)
{
Switch SensorNumber
{
Case "1PP": oWorkbook.Worksheets("398-435cm-1").Range("C4").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C5").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C6").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C7").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C8").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C9").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "2PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C15").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C16").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C17").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C18").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C19").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C20").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "3PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C26").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C27").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C28").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C29").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C30").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C31").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "4PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C37").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C38").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C39").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C40").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C41").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C42").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "5PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C48").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C49").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C50").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C51").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C52").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C53").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "6PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C59").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C60").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C61").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C62").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C63").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C64").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "7PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C70").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C71").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C72").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C73").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C74").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C75").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "8PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C81").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C82").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C83").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C84").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C85").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C86").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "9PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C92").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C93").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C94").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C95").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C96").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C97").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "10PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C103").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C104").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C105").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C106").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C107").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C108").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "11PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C114").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C115").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C116").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C117").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C118").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C119").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "12PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C125").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C126").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C127").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C128").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C129").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C130").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "13PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C136").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C137").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C138").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C139").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C140").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C141").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "14PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C147").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C148").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C149").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C150").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C151").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C152").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "15PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C158").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C159").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C160").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C161").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C162").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C163").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "16PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C169").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C170").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C171").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C172").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C173").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C174").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "17PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C180").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C181").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C182").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C183").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C184").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C185").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "18PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C191").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C192").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C193").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C194").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C195").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C196").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "19PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C202").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C203").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C204").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C205").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C206").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C207").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "20PP": 
oWorkbook.Worksheets("398-435cm-1").Range("C213").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C214").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C215").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C216").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C217").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("C218").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell
Default:
MsgBox, No Sensor Selected/Code is broken
return
}
} Else if(BTSensor = 1)
{
    Switch SensorNumber
    {
Case "1PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G4").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G5").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G6").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G7").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G8").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G9").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "2PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G15").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G16").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G17").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G18").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G19").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G20").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "3PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G26").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G27").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G28").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G29").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G30").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G31").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "4PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G37").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G38").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G39").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G40").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G41").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G42").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "5PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G48").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G49").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G50").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G51").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G52").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G53").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "6PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G59").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G60").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G61").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G62").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G63").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G64").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "7PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G70").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G71").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G72").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G73").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G74").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G75").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "8PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G81").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G82").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G83").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G84").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G85").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G86").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "9PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G92").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G93").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G94").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G95").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G96").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G97").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "10PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G103").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G104").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G105").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G106").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G107").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G108").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "11PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G114").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G115").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G116").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G117").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G118").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G119").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "12PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G125").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G126").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G127").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G128").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G129").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G130").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "13PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G136").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G137").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G138").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G139").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G140").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G141").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "14PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G147").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G148").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G149").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G150").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G151").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G152").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "15PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G158").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G159").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G160").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G161").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G162").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G163").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "16PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G169").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G170").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G171").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G172").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G173").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G174").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "17PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G180").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G181").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G182").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G183").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G184").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G185").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "18PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G191").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G192").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G193").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G194").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G195").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G196").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "19PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G202").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G203").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G204").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G205").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G206").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G207").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "20PP": 
oWorkbook.Worksheets("398-435cm-1").Range("G213").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G214").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G215").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G216").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G217").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("398-435cm-1").Range("G218").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell
Default:
MsgBox, No Sensor Selected/Code is broken
}
}
else
{
MsgBox, Neither option was selected
}
oWorkbook.Close(1)


setIntegrationRange(927,1041)

moveToSpectra("R1")
sleep, 200
R1 := copyIntegration()
moveToSpectra("R2")
sleep, 200
R2 := copyIntegration()
moveToSpectra("R3")
sleep, 200
R3 := copyIntegration()
moveToSpectra("R4")
sleep, 200
R4 := copyIntegration()
moveToSpectra("R5")
sleep, 200
R5 := copyIntegration()
moveToSpectra("Avg")
sleep, 200
Avg := copyIntegration()


Gui, Submit, Nohide
oWorkBook := ComObjGet(FilePath)		; get reference to WorkBook
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1	; just do it - too long to explain why...

If (BlankSensor = 1)
{
Switch SensorNumber
{
Case "1PP": oWorkbook.Worksheets("927-1041cm-1").Range("C4").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C5").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C6").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C7").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C8").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C9").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "2PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C15").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C16").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C17").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C18").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C19").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C20").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "3PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C26").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C27").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C28").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C29").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C30").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C31").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "4PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C37").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C38").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C39").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C40").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C41").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C42").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "5PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C48").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C49").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C50").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C51").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C52").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C53").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "6PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C59").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C60").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C61").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C62").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C63").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C64").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "7PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C70").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C71").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C72").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C73").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C74").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C75").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "8PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C81").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C82").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C83").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C84").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C85").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C86").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "9PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C92").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C93").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C94").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C95").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C96").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C97").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "10PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C103").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C104").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C105").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C106").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C107").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C108").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "11PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C114").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C115").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C116").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C117").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C118").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C119").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "12PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C125").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C126").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C127").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C128").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C129").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C130").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "13PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C136").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C137").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C138").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C139").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C140").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C141").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "14PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C147").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C148").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C149").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C150").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C151").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C152").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "15PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C158").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C159").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C160").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C161").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C162").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C163").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "16PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C169").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C170").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C171").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C172").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C173").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C174").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "17PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C180").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C181").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C182").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C183").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C184").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C185").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "18PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C191").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C192").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C193").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C194").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C195").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C196").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "19PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C202").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C203").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C204").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C205").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C206").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C207").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "20PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("C213").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C214").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C215").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C216").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C217").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("C218").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell
Default:
MsgBox, No Sensor Selected/Code is broken
return
}
} Else if(BTSensor = 1)
{
    Switch SensorNumber
    {
Case "1PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G4").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G5").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G6").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G7").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G8").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G9").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "2PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G15").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G16").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G17").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G18").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G19").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G20").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "3PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G26").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G27").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G28").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G29").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G30").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G31").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "4PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G37").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G38").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G39").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G40").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G41").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G42").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "5PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G48").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G49").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G50").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G51").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G52").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G53").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "6PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G59").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G60").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G61").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G62").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G63").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G64").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "7PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G70").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G71").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G72").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G73").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G74").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G75").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "8PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G81").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G82").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G83").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G84").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G85").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G86").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "9PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G92").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G93").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G94").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G95").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G96").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G97").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "10PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G103").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G104").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G105").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G106").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G107").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G108").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "11PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G114").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G115").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G116").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G117").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G118").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G119").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "12PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G125").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G126").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G127").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G128").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G129").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G130").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "13PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G136").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G137").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G138").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G139").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G140").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G141").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "14PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G147").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G148").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G149").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G150").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G151").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G152").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "15PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G158").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G159").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G160").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G161").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G162").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G163").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "16PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G169").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G170").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G171").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G172").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G173").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G174").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "17PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G180").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G181").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G182").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G183").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G184").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G185").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "18PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G191").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G192").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G193").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G194").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G195").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G196").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "19PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G202").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G203").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G204").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G205").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G206").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G207").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell

Case "20PP": 
oWorkbook.Worksheets("927-1041cm-1").Range("G213").Value := R1	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G214").Value := R2	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G215").Value := R3	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G216").Value := R4	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G217").Value := R5	; set "Test" in "Sheet1" sheet, "A2" cell
oWorkbook.Worksheets("927-1041cm-1").Range("G218").Value := Avg	; set "Test" in "Sheet1" sheet, "A2" cell
Default:
MsgBox, No Sensor Selected/Code is broken
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