Attribute VB_Name = "Protection"


Sub LockCells()
'this submodule is intended to lock cells based on the cell colour
'this macro takes a long time to be run and is not intended to be called from another module but to be run manually only.

Run "UnProtector"

On Error Resume Next
Dim ws As Worksheet
Dim c As Range
Dim Rng1 As Range
back:
Set Rng1 = Application.InputBox("select cell", Type:=8)
If Rng1.Value = "" Then GoTo back

ColourCheck = Rng1.Interior.ColorIndex

    For Each ws In ThisWorkbook.Worksheets
    ws.Unprotect
        For Each c In ws.UsedRange
        
        If c.Interior.ColorIndex = ColourCheck Then
        c.Locked = False
        Else
        c.Locked = True
        End If
        
        Next c
    ws.Protect
    Next ws

Run "Protector"

End Sub
Private Sub Protector()
Attribute Protector.VB_ProcData.VB_Invoke_Func = " \n14"
On Error Resume Next

If Range("DeveloperMode").Value = True Then End

Application.ScreenUpdating = False
Dim WorksheetPass As String
WorksheetPass = "Egis"


Call hide

ThisWorkbook.Protect Password:=WorksheetPass

Application.ScreenUpdating = False

Dim ws As Worksheet
 
    For Each ws In ThisWorkbook.Worksheets
    ws.Protect Password:=WorksheetPass, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws.EnableSelection = xlUnlockedCells
    Next ws

ThisWorkbook.Sheets("Results").EnableSelection = 0
ThisWorkbook.Protect Password:=WorksheetPass

Application.ScreenUpdating = True

End Sub

Private Sub callUnprotector()
Call UnProtector
End Sub
Private Sub UnProtector(Optional ShtName As String)

Application.ScreenUpdating = False
On Error Resume Next
Dim WorksheetPass As String
WorksheetPass = "Egis"
Dim StartSheet As Worksheet
Set StartSheet = ThisWorkbook.ActiveSheet

ThisWorkbook.Unprotect Password:=WorksheetPass

Call unhide

Application.ScreenUpdating = False

Dim ws As Worksheet
 
If ShtName = "" Then
    For Each ws In ThisWorkbook.Worksheets
    'ws.Visible = True
    'ws.Cells.ApplyNames IgnoreRelativeAbsolute:=True, UseRowColumnNames:=True, OmitColumn:=True, OmitRow:=True, Order:=1, AppendLast:=False
    ws.Unprotect Password:=WorksheetPass
    ws.EnableSelection = 0
    
    Next ws
Else
    Sheets(ShtName).Unprotect Password:=WorksheetPass
    Sheets(ShtName).EnableSelection = 0
End If

Run "InputsScrollLimitNothing"

StartSheet.Select

Application.ScreenUpdating = True

End Sub
Sub HideAllSheets()
On Error Resume Next
Dim ws As Worksheet
'Hide all sheets except the first sheet in navigation
    For Each ws In ThisWorkbook.Worksheets
            If ws.name = Range("FirstVSheet").Value Then GoTo Jump
            ws.Visible = xlVeryHidden
Jump:
    Next ws
End Sub

Sub hide()
On Error Resume Next
Dim VSheet As Range

Application.ScreenUpdating = False

'Hide All
Call HideAllSheets

'Unhide Sheets in the navigation menu
For Each VSheet In Range("TabSheetName[Worksheet Short Name]")

Sheets(VSheet.Value).Visible = True

Next VSheet

Call InputsScrollLimit
 
Application.ScreenUpdating = True
End Sub

Sub unhide()
On Error Resume Next
Application.ScreenUpdating = False
On Error Resume Next
Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
            ws.Visible = True
    Next ws
Application.ScreenUpdating = True
End Sub

Sub InputsScrollLimit()
Application.ScreenUpdating = False
Dim VSheet As Range
Dim rng As String

For Each VSheet In Range("TabSheetName[Worksheet Short Name]")
If IsEmpty(VSheet) Then End
If VSheet.Offset(0, 1).Value = "Input" Then rng = "A1:R65535" Else rng = "A1:IV65535"
Sheets(VSheet.Value).ScrollArea = "A1:" & LastNonBlankCell(Sheets(VSheet.Value), rng)

Next VSheet

Application.ScreenUpdating = True
End Sub

Sub InputsScrollLimitNothing()
Application.ScreenUpdating = False
Dim VNSheet As Worksheet

For Each VNSheet In Worksheets
VNSheet.ScrollArea = ""

Next VNSheet

Application.ScreenUpdating = True

End Sub
