Attribute VB_Name = "modProgress"
Option Explicit ' Always a good idea to use this


Sub ShowProgress(ByVal ActionNumber As Long, _
                ByVal TotalActions As Long, _
                Optional ByVal StatusMessage As String = vbNullString, _
                Optional ByVal CloseWhenDone As Boolean = True, _
                Optional ByVal Title As String = vbNullString)

DoEvents 'to ensure that the code to display the form gets executed

'Display the Proressbar
If isFormOpen("ufProgress") Then
    'If the form is already open, just update the ActionNumbers and Status
    'message
    Call ufProgress.UpdateForm(ActionNumber, TotalActions, StatusMessage)
Else
    'if the form is not already open, Show it
    ufProgress.Show
    'set the title
    If Not Title = vbNullString Then
        ufProgress.Caption = Title
    End If
    'then update the ActionNumber and Status Message
    Call ufProgress.UpdateForm(ActionNumber, TotalActions, StatusMessage)
End If

'If the user chose to close the form automatically when the last action
'is reached, close it
If CloseWhenDone And CBool(ActionNumber >= TotalActions) Then
    Unload ufProgress
End If

End Sub


Function isFormOpen(ByVal FormName As String) As Boolean
'Declare Function level Objects
Dim ufForm As Object
'Set the Function to False
isFormOpen = False
'Loop through all the open forms
For Each ufForm In VBA.UserForms
    'Check the form names
    If ufForm.name = FormName Then
        'if the form is open, set the function value to True
        isFormOpen = True
        'and exit the loop
        Exit For
    End If
Next ufForm
End Function
