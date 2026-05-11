Attribute VB_Name = "Complications_Module"
Option Explicit

'Public Type complication_Prob

'complication ID
'ID As Integer
'Prob6m As Single
'Probability 6 month - 12 month
'Prob6m12m As Single
'Probability 13 month and later
'Prob13m As Single

'End Type

Public Type Complication
'complication name
name As String
'complication ID
ID As Integer
'utility decrement
Utility_Decrement As Double
'Length of complication in weeks
Length As Single
'Probability 6 month
Prob6m As Single
'Probability 6 month - 12 month
Prob6m12m As Single
'Probability 13 month and later
Prob13m As Single
'Complication cost
Cost As Double
'Included in the model or not
Active As Boolean

End Type

Function Load_Complications() As Complication()

'Collection of complications included in the model
Dim Col_Complications() As Complication

'general index variables
Dim i As Byte
Dim j As Byte

'Get range where complications are located
Dim Complications_Matrix As Variant
Complications_Matrix = Range("Complications")

'get number of complications
NComplications = UBound(Complications_Matrix)

'check how many complications are active in the model
Dim Temp_Matrix As Variant
Temp_Matrix = Application.Transpose(Complications_Matrix)
Temp_Matrix = Application.WorksheetFunction.index(Temp_Matrix, 6, 0)
Active_Complications = CByte(ArrayCountIf(Temp_Matrix, True))

ReDim Col_Complications(1 To Active_Complications)

Col_Complications(1).ID = 1

'Load Complications from input parameters (excel)
'set active Complications counter as a start = 1
j = 1
For i = 1 To NComplications

      With Col_Complications(j)
            
            If CBool(Complications_Matrix(i, 6)) = True Then
            
                  .ID = CByte(Complications_Matrix(i, 1))
                  .name = CStr(Complications_Matrix(i, 2))
                  .Utility_Decrement = CDbl(Complications_Matrix(i, 3))
                  .Length = CSng(Complications_Matrix(i, 4))
                  .Cost = CDbl(Complications_Matrix(i, 5))
                  .Active = CBool(Complications_Matrix(i, 6))
                  j = j + 1
                  If j > Active_Complications Then GoTo Last_Complication
                  
            End If
            
      End With

Next i

Last_Complication:

Load_Complications = Col_Complications

End Function

