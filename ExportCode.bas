Attribute VB_Name = "ExportCode"
Option Explicit

Sub ExportAllModules()
    Dim vbComp As Object
    Dim FolderPath As String
    Dim ExportPath As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Export Modules"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        FolderPath = .SelectedItems(1) & "\"
    End With

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard module
                ExportPath = FolderPath & vbComp.name & ".bas"
            Case 2 ' Class module
                ExportPath = FolderPath & vbComp.name & ".cls"
            Case 3 ' UserForm
                ExportPath = FolderPath & vbComp.name & ".frm"
            Case Else
                GoTo SkipExport
        End Select

        vbComp.Export ExportPath

SkipExport:
    Next vbComp

    MsgBox "Modules exported successfully to: " & FolderPath, vbInformation
End Sub

