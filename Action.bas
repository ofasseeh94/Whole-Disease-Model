Attribute VB_Name = "Action"
'General Variables


Dim TRI As Integer
Dim SheetType As String

Option Explicit

Const Mname_DCost As String = "Action_Menu"
'*************************************************************

Sub DeleteActionPopUpMenu()
    ' Delete the popup menu if it already exists.
    On Error Resume Next
    Application.CommandBars(Mname_DCost).Delete
    On Error GoTo 0
End Sub

Sub CreateActionPopUpMenu()
    Call DeleteActionPopUpMenu
    Call Action_PopUpMenu
    On Error Resume Next
    Application.CommandBars(Mname_DCost).ShowPopup
    On Error GoTo 0
End Sub

Sub Action_PopUpMenu()
    Dim MenuItem As CommandBarPopup
    ' Add the popup menu.
    With Application.CommandBars.Add(name:=Mname_DCost, Position:=msoBarPopup, _
         MenuBar:=False, Temporary:=True)
       
'        With .Controls.Add(Type:=msoControlButton)
'            .Caption = "Run Simulation"
'            .OnAction = "Sim"
'            .BeginGroup = True
'        End With
'        With .Controls.Add(Type:=msoControlButton)
'            .Caption = "Delete Simulation Results"
'            .OnAction = "ClearSimPat"
'        End With
       
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Load default"
            .OnAction = "RestoreDefault"
            .BeginGroup = True
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Save current result"
            .OnAction = "save_curr_result"
            .BeginGroup = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Delete saved results"
            .OnAction = "reset_result"
            '.BeginGroup = True
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Save current settings"
            .OnAction = "SaveCurrent"
            .BeginGroup = True
        End With
        

        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Load saved settings"
            .OnAction = "LoadScenario"
            '.BeginGroup = True
        End With
        

                
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Delete saved settings"
            .OnAction = "DeleteScenario"
            '.BeginGroup = True
        End With
        
GoTo Line1
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Save current Results"
            .OnAction = "Show_savecur"
            '.BeginGroup = True
        End With

        
Line1:
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Run probabilistic analysis"
            .OnAction = "PSA"
            .BeginGroup = True
        End With
    
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Run one way deterministic analysis"
            .OnAction = "DSA"
            .BeginGroup = True
        End With
    
    End With
    
End Sub

Sub RestoreDefault()

'Run "UnProtector"
Run "UnProtector"

'turn off screen updating
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim x As Double

'Load inputs
x = 20
ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Default").Range(Cells(10, x - 1), Cells(800, x - 1)).Copy
ThisWorkbook.Sheets("input").Range("H10").PasteSpecial xlPasteAll

'Load settings
ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Default").Range(Cells(10, x), Cells(100, x)).Copy
ThisWorkbook.Sheets("Settings").Range("L10").PasteSpecial xlPasteAll


'Load other sheet
TRI = Sheets("Template").Cells(4, 4).Value + Sheets("Template").Cells(5, 4).Value + 5 + 1
SheetType = Sheets("Template").Cells(TRI, 9).Value

Do Until SheetType <> "Input"
      
      If Sheets("Template").Cells(TRI, 8).Value = "Settings" Then GoTo SkipSheet
      
      With ThisWorkbook.Sheets(Sheets("Template").Cells(TRI, 8).Value)
      .Activate
      .Range(Cells(10, x - 1), Cells(200, x)).Copy
      .Range("E10").PasteSpecial xlPasteAll
      End With
      
SkipSheet:
    TRI = TRI + 1
    SheetType = Sheets("Template").Cells(TRI, 9).Value
Loop


'Run "Protector"
Run "Protector"

'turn on screen updating
Sheets("Results").Activate
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox "Base Case" & " loaded"

End Sub
Sub SaveDefaultOnce()
'This macro saves the current values as defaults and it is intended to be run manually only and not to be called from other macros
'Run "UnProtector"
Run "UnProtector"

'turn off screen updating
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim x As Double


x = 20

Line20:
'save scenario name
ThisWorkbook.Sheets("Default").Cells(9, x).Value = "BaseCase"


'Save inputs in the left column
ThisWorkbook.Sheets("input").Range("H10:H800").Copy

ThisWorkbook.Sheets("Default").Cells(10, x - 1).PasteSpecial xlPasteAll


'Save settings in the next column
ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Settings").Range("L10:L100").Copy
ThisWorkbook.Sheets("Default").Cells(10, x).PasteSpecial xlPasteAll

'Save other sheet
TRI = Sheets("Template").Cells(4, 4).Value + Sheets("Template").Cells(5, 4).Value + 5 + 1
SheetType = Sheets("Template").Cells(TRI, 9).Value

Do Until SheetType <> "Input"

      If Sheets("Template").Cells(TRI, 8).Value = "Settings" Then GoTo SkipSheet
       
      With ThisWorkbook.Sheets(Sheets("Template").Cells(TRI, 8).Value)
      .Activate
      .Range("E10:F200").Copy
      .Cells(10, x - 1).PasteSpecial xlPasteAll
      End With
      
SkipSheet:

    TRI = TRI + 1
    SheetType = Sheets("Template").Cells(TRI, 9).Value
Loop



'turn on screen updating
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


'Run "Protector"
Run "Protector"

MsgBox "Base Case" & " saved"

End Sub

Sub SaveCurrent()


'Run "UnProtector"
Run "UnProtector"

'turn off screen updating
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim CurrentSheet As String
Dim x As Double
Dim ScenarioName As String

CurrentSheet = ThisWorkbook.ActiveSheet.name


'Get Scenario Name
Line0:
ScenarioName = InputBox("Please name the scenario you want to save", "Scenario Name")
If ScenarioName = vbNullString Then Exit Sub

For x = 22 To 200 Step (2)
If ThisWorkbook.Sheets("Default").Cells(9, x).Value = ScenarioName Then
MsgBox "Sorry the name you have choosen has been used before, please use a different name"
GoTo Line0
ElseIf ThisWorkbook.Sheets("Default").Cells(9, x).Value = "" Then GoTo Line20
End If
Next x

Line20:
'save scenario name
ThisWorkbook.Sheets("Default").Cells(9, x).Value = ScenarioName


'Save inputs in the left column
ThisWorkbook.Sheets("input").Range("H10:H600").Copy

ThisWorkbook.Sheets("Default").Cells(10, x - 1).PasteSpecial xlPasteAll


'Save settings in the next column
ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Settings").Range("L10:L58").Copy
ThisWorkbook.Sheets("Default").Cells(10, x).PasteSpecial xlPasteAll

'Save other sheet
TRI = Sheets("Template").Cells(4, 4).Value + Sheets("Template").Cells(5, 4).Value + 5 + 1
SheetType = Sheets("Template").Cells(TRI, 9).Value

Do Until SheetType <> "Input"
      If Sheets("Template").Cells(TRI, 8).Value = "Settings" Then GoTo SkipSheet
      
      With ThisWorkbook.Sheets(Sheets("Template").Cells(TRI, 8).Value)
      .Activate
      .Range("E10:F200").Copy
      .Cells(10, x - 1).PasteSpecial xlPasteAll
      End With
 
SkipSheet:
 
    TRI = TRI + 1
    SheetType = Sheets("Template").Cells(TRI, 9).Value
Loop


'turn on screen updating

Sheets(CurrentSheet).Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


'Run "Protector"
Run "Protector"


MsgBox ScenarioName & " saved"

End Sub
Sub LoadScenario()

'Run "UnProtector"
Run "UnProtector"

'turn off screen updating
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim zz As Double

'Check if there are saved scenarios
If ThisWorkbook.Sheets("Default").Cells(9, 22).Value = "" Then
MsgBox "No saved scenarios to load from. Please save a scenario first."
End
End If


'Get scenario to be loaded
For zz = 22 To 100 Step (2)
If ThisWorkbook.Sheets("Default").Cells(9, zz).Value <> "" Then
ListBox.ListBox1.AddItem ThisWorkbook.Sheets("Default").Cells(9, zz).Value
ElseIf ThisWorkbook.Sheets("Default").Cells(9, zz).Value = "" Then GoTo Line1
End If
Next
Line1:

ListBox.Show

Dim Choice As String
Dim x As Double

Choice = ThisWorkbook.Sheets("Default").Cells(9, 20).Value

If MsgBox("Are you sure you want to load scenario (" & Choice & ")", vbYesNo + vbQuestion, "Empty Sheet") = vbNo Then GoTo Line2


'Load inputs
For x = 22 To 200 Step (2)
If ThisWorkbook.Sheets("Default").Cells(9, x).Value = Choice Then
ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Default").Range(Cells(10, x - 1), Cells(600, x - 1)).Copy
ThisWorkbook.Sheets("input").Range("H10").PasteSpecial xlPasteAll

'Load settings

ThisWorkbook.Sheets("Default").Activate
ThisWorkbook.Sheets("Default").Range(Cells(10, x), Cells(58, x)).Copy
ThisWorkbook.Sheets("Settings").Range("L10").PasteSpecial xlPasteAll


'Load other sheet
TRI = Sheets("Template").Cells(4, 4).Value + Sheets("Template").Cells(5, 4).Value + 5 + 1
SheetType = Sheets("Template").Cells(TRI, 9).Value

Do Until SheetType <> "Input"
    
If Sheets("Template").Cells(TRI, 8).Value = "Settings" Then GoTo SkipSheet
    
      With ThisWorkbook.Sheets(Sheets("Template").Cells(TRI, 8).Value)
      .Activate
      .Range(Cells(10, x - 1), Cells(200, x)).Copy
      .Range("E10").PasteSpecial xlPasteAll
      End With

SkipSheet:

    TRI = TRI + 1
    SheetType = Sheets("Template").Cells(TRI, 9).Value
Loop

 
ElseIf ThisWorkbook.Sheets("Default").Cells(9, x).Value = "" Then GoTo Line2
End If
Next
Line2:


'Run "Protector"
Run "Protector"

'turn on screen updating
Sheets("Results").Activate
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox Choice & " loaded"

End Sub

Sub DeleteScenario()

Run "UnProtector"

'turn off screen updating
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim zz As Double

For zz = 22 To 200 Step (2)
'Check if there are saved scenarios
If ThisWorkbook.Sheets("Default").Cells(9, 22).Value = "" Then
MsgBox "No saved scenarios to delete"
End
End If


'Get scenario to be deleted
If ThisWorkbook.Sheets("Default").Cells(9, zz).Value <> "" Then
ListBox.ListBox1.AddItem ThisWorkbook.Sheets("Default").Cells(9, zz).Value
ElseIf ThisWorkbook.Sheets("Default").Cells(9, zz).Value = "" Then GoTo Line1
End If
Next
Line1:

ListBox.Show

Dim Choice As String
Dim x As Double

Choice = ThisWorkbook.Sheets("Default").Cells(9, 20).Value

'Delete stored data
For x = 22 To 200 Step (2)
If ThisWorkbook.Sheets("Default").Cells(9, x).Value = Choice Then

ThisWorkbook.Sheets("Default").Cells(1, x).EntireColumn.Delete
ThisWorkbook.Sheets("Default").Cells(1, x - 1).EntireColumn.Delete

'Load other sheet
TRI = Sheets("Template").Cells(4, 4).Value + Sheets("Template").Cells(5, 4).Value + 5 + 1
SheetType = Sheets("Template").Cells(TRI, 9).Value

Do Until SheetType <> "Input"
    
If Sheets("Template").Cells(TRI, 8).Value = "Settings" Then GoTo SkipSheet
    
      With ThisWorkbook.Sheets(Sheets("Template").Cells(TRI, 8).Value)
      .Activate
      .Range(Cells(10, x - 1), Cells(200, x)).EntireColumn.Delete
      End With

SkipSheet:

    TRI = TRI + 1
    SheetType = Sheets("Template").Cells(TRI, 9).Value
Loop

ElseIf ThisWorkbook.Sheets("Default").Cells(9, x).Value = "" Then GoTo Line2
End If
Next

Line2:

Sheets("Settings").Activate

'Run "Protector"
Run "Protector"

'turn on screen updating
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox Choice & " deleted"

End Sub


Sub Show_savecur()
Call save_curr_result
End Sub
Sub save_curr_result()
'save the current results


Run "UnProtector", "Results"

Application.ScreenUpdating = False


Sheets("Results").Activate
Dim ID As Integer
Dim strName As String

Dim c_name As String
c_name = ActiveSheet.name
ID = Range("save_ID")
strName = InputBox(Prompt:="Enter scenario name e.g. 'base case 2'", Title:="Enter scenario name", Default:="", Xpos:=300, Ypos:=300)
    
    With Sheets("Results")
    
    .Range(Cells(ID, 3), Cells(ID, 4)).Value = .Range("C9:D9").Value
    .Range(Cells(ID, 5), Cells(ID, 6)).Value = .Range("C10:D10").Value
    .Range(Cells(ID, 7), Cells(ID, 8)).Value = .Range("C11:D11").Value
    .Cells(ID, 9).Value = .Range("C12").Value
    Cells(ID, 2) = strName
    
    End With
    
Sheets(c_name).Select

Application.ScreenUpdating = True
Run "Protector"

End Sub


Sub reset_result()
' delete the saved results
' reset_result Macro
Dim answer As Integer

answer = MsgBox("Are you sure you want to delete ALL saved results?", vbYesNo + vbQuestion, "Delete saved results")
If answer = vbNo Then End

Run "UnProtector", "Results"

Application.ScreenUpdating = False
Sheets("Results").Range("B18:I36").ClearContents

Run "Protector"

Application.ScreenUpdating = True
End Sub
Sub RestoreFormulas()
'Do 'Run "UnProtector"
Run "UnProtector"
ThisWorkbook.Sheets("Default").Range("H10:H800").Copy
ThisWorkbook.Sheets("input").Range("H10").PasteSpecial xlPasteAll

End Sub
