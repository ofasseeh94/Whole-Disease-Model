Attribute VB_Name = "ExportCode"
Option Explicit

Sub ExportAllModules()
    Dim vbComp As VBIDE.VBComponent
    Dim FolderPath As String
    Dim ExportPath As String

    ' Prompt for folder path
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Export Modules"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Exit if no folder selected
        FolderPath = .SelectedItems(1) & "\"
    End With

    ' Loop through all VBComponents
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                ExportPath = FolderPath & vbComp.name & ".bas"
            Case vbext_ct_ClassModule
                ExportPath = FolderPath & vbComp.name & ".cls"
            Case vbext_ct_MSForm
                ExportPath = FolderPath & vbComp.name & ".frm"
            Case Else
                ' Skip other types like Document modules
                GoTo SkipExport
        End Select

        ' Export component
        vbComp.Export ExportPath

SkipExport:
    Next vbComp

    MsgBox "Modules exported successfully to: " & FolderPath, vbInformation
End Sub

