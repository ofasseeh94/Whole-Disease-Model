Attribute VB_Name = "Navigation_Menu"
Sub Navigation_Menu()

'Turn of screen updating
Application.ScreenUpdating = False

Application.DisplayAlerts = False

Application.Calculation = xlCalculationManual
'
' Sheets Copy Macro is intended to copy names of sheets from the Overview sheet to the Template sheet to be used in creating dropdown menu in navigation.
'
Sheets("Template").Visible = True

    'Clear Previous data
    Sheets("Template").Select
    Range("F5:H50").ClearContents
    
    Sheets("Overview").Select
    Range("A33:C33").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Template").Select
    Range("F5").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    ' Delete table if present
            On Error Resume Next
            
            Debug.Print Range("TabSheetName").Rows.Count
            
            If Err = 1004 Then

            Else

                'ActiveWorkbook.Names("TornadoCEARef").Delete
                ActiveSheet.ListObjects("TabSheetName").Unlist
            End If
            
            Err.Clear
            On Error GoTo 0
    
    ' Create a table from coppied data
    Sheets("Template").Activate
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$F$5:$H$50"), , xlYes).name = _
        "TabSheetName"
        
        
    Sheets("Template").Select
    Range("TabSheetName[#All]").Select
    ActiveWorkbook.Worksheets("Template").ListObjects("TabSheetName").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Template").ListObjects("TabSheetName").Sort.SortFields. _
        Add Key:=Range("TabSheetName[[#All],[Worksheet Type]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Template").ListObjects("TabSheetName").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    ThisWorkbook.Sheets("Template").Range("C4").Formula = "=COUNTIFS(TabSheetName[Worksheet Type], B4)"
    ThisWorkbook.Sheets("Template").Range("C5").Formula = "=COUNTIFS(TabSheetName[Worksheet Type], B5)"
    ThisWorkbook.Sheets("Template").Range("C6").Formula = "=COUNTIFS(TabSheetName[Worksheet Type], B6)"
    ThisWorkbook.Sheets("Template").Range("C7").Formula = "=COUNTIFS(TabSheetName[Worksheet Type], B7)"
    
    ThisWorkbook.Sheets("Template").Range("C10").Formula = "=COUNTA(TabSheetName[Worksheet Type])"
    
Sheets("Template").Visible = False
    
    
Application.ScreenUpdating = True

Application.DisplayAlerts = True

Application.Calculation = xlCalculationAutomatic

End Sub

