Attribute VB_Name = "BIM_Data"
Option Explicit

Sub AutomaticBIMCosts()

Dim wsSettings As Worksheet, wsResults As Worksheet, wsBIMCosts As Worksheet

Set wsSettings = ThisWorkbook.Sheets("Settings")
Set wsResults = ThisWorkbook.Sheets("Results")
Set wsBIMCosts = ThisWorkbook.Sheets("BIMCosts")

Dim Defaultsettings As Variant
           
      'Save Default settings
      Defaultsettings = wsSettings.Range("L12:L34").Value
      
Dim i As Integer

'Loop over 5 years in 6 months cycles
      For i = 1 To 10
      
        'Change time horizon to the current month divded by 12
          wsSettings.Range("L26").Value = i / 2
        'Run the model
              Call Engine
        'Copy model results and paste it into the range allocated in BIM costs
              wsBIMCosts.Range("D4:M4").Offset(i - 1, 0) = wsResults.Range("F14:O14").Value

        Next i

'Restore default settings
    '  wsSettings.Range("L12:L34").Value = Defaultsettings
'Run the model to restore default results
  ' Call Engine
    
End Sub

Sub ClearMS()

If Range("BIM_Intervention") = 1 Then

    Range("D20:F23").Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15849925
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With Worksheets("BIM Settings").Range("G20:G23")
    .ClearContents
    .Interior.Color = vbWhite
    End With

    Else

    With Worksheets("BIM Settings").Range("D20:F23")
    .ClearContents
    .Interior.Color = vbWhite
    End With
    
    Worksheets("BIM Settings").Range("G20:G23").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15849925
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
End If

End Sub

