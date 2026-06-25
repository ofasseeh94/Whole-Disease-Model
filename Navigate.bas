Attribute VB_Name = "Navigate"
Option Explicit

Const Mname_navigate As String = "Navigate_Menu"


Sub Enter_button()
    Sheets("Overview").Select
End Sub

Sub DeleteNavigatePopUpMenu()
    ' Delete the popup menu if it already exists.
    On Error Resume Next
    Application.CommandBars(Mname_navigate).Delete
    On Error GoTo 0
End Sub

'Sub CreateNavigatePopUpMenu()
    ' Delete any existing popup menu.
 '   Call DeleteNavigatePopUpMenu

    ' Create the popup menu.
 '   Call Navigate_PopUpMenu

    ' Display the popup menu.
 '   On Error Resume Next
 '   Application.CommandBars(Mname_navigate).ShowPopup
 '   On Error GoTo 0
'End Sub


Sub Navigate_PopUpMenu()

Dim BGsheetsNo As Long, INPTsheetsNo As Long, CALCsheetsNo As Long, RSLTsheetsNo As Long, MAXsheetsNo As Long
MAXsheetsNo = Sheets("Template").Range("MAXsheetsNo").Value

Dim i As Long, j As Long, k As Long, m As Long, N As Long
Dim PopupMenu As CommandBar, Menu(100) As CommandBarControl, Submenu(100) As CommandBarControl, HaveMenu As Boolean

BGsheetsNo = Sheets("Template").Range("BGsheetsNo").Value
INPTsheetsNo = Sheets("Template").Range("INPTsheetsNo").Value
CALCsheetsNo = Sheets("Template").Range("CALCsheetsNo").Value
RSLTsheetsNo = Sheets("Template").Range("RSLTsheetsNo").Value

On Error Resume Next

    'Delete the menu if it already exists and recreate it.
    Application.CommandBars("SheetMenu").Delete
    Set PopupMenu = CommandBars.Add("SheetMenu", msoBarPopup, , True)
    
    'Add four main menus.
      For i = 1 To UBound(Sheets("Template").Range("main.menu").Value) - 1
            
            If Sheets("Template").Range("main.menu").Cells(i + 1, 2).Value > 0 Then
                  Set Menu(i) = PopupMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
            End If
            
      Next i
      
    'Name the menus as shown in the menus sheet.
    For i = 1 To 4
        Menu(i).Caption = Sheets("Template").Range("main.menu").Cells(i + 1, 1).Value
    Next i
    
    'For the second menu (BackGround), add the submenus and assign gotosheet macro.
    For j = 1 To BGsheetsNo
        Set Submenu(j) = Menu(1).Controls.Add(Type:=msoControlButton, Temporary:=True)
        Submenu(j).Caption = Sheets("Template").Range("TabSheetName").Cells(j, 1).Value
        Submenu(j).Parameter = Sheets("Template").Range("TabSheetName").Cells(j, 2).Value
        Submenu(j).OnAction = "GoToSheet"
    Next j
    
    'For the second menu (Calculations), add the submenus and assign gotosheet macro.
    For k = 1 To CALCsheetsNo
        Set Submenu(k) = Menu(2).Controls.Add(Type:=msoControlButton, Temporary:=True)
        Submenu(k).Caption = Sheets("Template").Range("TabSheetName").Cells(k + BGsheetsNo, 1).Value
        Submenu(k).Parameter = Sheets("Template").Range("TabSheetName").Cells(k + BGsheetsNo, 2).Value
        Submenu(k).OnAction = "GoToSheet"
    Next k
     
    'For the third menu (Inputs), add the submenus and assign gotosheet macro.
    For m = 1 To INPTsheetsNo
        Set Submenu(m) = Menu(3).Controls.Add(Type:=msoControlButton, Temporary:=True)
        Submenu(m).Caption = Sheets("Template").Range("TabSheetName").Cells(m + BGsheetsNo + CALCsheetsNo, 1).Value
        Submenu(m).Parameter = Sheets("Template").Range("TabSheetName").Cells(m + BGsheetsNo + CALCsheetsNo, 2).Value
        Submenu(m).OnAction = "GoToSheet"
    Next m
    
    'For the fourth menu (Results), add the submenus and assign gotosheet macro.
    For N = 1 To RSLTsheetsNo
        Set Submenu(N) = Menu(4).Controls.Add(Type:=msoControlButton, Temporary:=True)
        Submenu(N).Caption = Sheets("Template").Range("TabSheetName").Cells(N + BGsheetsNo + INPTsheetsNo + CALCsheetsNo, 1).Value
        Submenu(N).Parameter = Sheets("Template").Range("TabSheetName").Cells(N + BGsheetsNo + INPTsheetsNo + CALCsheetsNo, 2).Value
        Submenu(N).OnAction = "GoToSheet"
    Next N

'Show the menu.
CommandBars("SheetMenu").ShowPopup
End Sub
Sub GoToSheet()
    'Go to the sheet passed from the commandbar parameter.
    ActiveWorkbook.Sheets(Application.CommandBars.ActionControl.Parameter).Activate
    Range("A1").Select
    
End Sub

Sub ListVisibleSheets()
    Dim ws As Worksheet
    Dim listSheet As Worksheet
    Dim i As Integer

    ' Add a new worksheet to list the sheet names
    Set listSheet = Sheets("Template")
    i = 1

    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not hidden or very hidden
        If ws.Visible = xlSheetVisible Then
            ' Add the sheet name to the list
            listSheet.Cells(i + 26, 3).Value = ws.name
            i = i + 1
        End If
    Next ws

 End Sub
