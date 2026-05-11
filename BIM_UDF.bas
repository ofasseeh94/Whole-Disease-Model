Attribute VB_Name = "BIM_UDF"
Option Explicit
Function BUDGETIMPACT(CohortRange As Range, CostRange As Range) As Double

Dim TotalBudget As Double

If CohortRange.Count = 1 And CostRange.Count = 1 Then

      TotalBudget = CohortRange * CostRange
      
Else

      Dim Cohort As Variant
      Dim Cost As Variant
      
      Cohort = CohortRange
      Cost = CostRange
      
      'if range is horizontal transpose it
      If Not RangeIsVertical(CohortRange) Then Cohort = Application.Transpose(CohortRange)
      If Not RangeIsVertical(CostRange) Then Cost = Application.Transpose(CostRange)
      
      Cohort = ConvUnivariant(Cohort)
      Cost = ConvUnivariant(Cost)
      
      Dim LenCost As Integer, LenCohort As Integer
      
      LenCohort = UBound(Cohort) - LBound(Cohort) + 1
      LenCost = UBound(Cost) - LBound(Cost) + 1
      
      Dim j As Long, i As Long
      
      
            For j = 1 To LenCost
            
                  For i = 1 To LenCohort
                  'Cost = ChangeArrayLength(Cost, LenCohort - i + 1)
                  'LenCost = UBound(Cost) - LBound(Cost)
                  
                  
                  If j >= i Then
                        TotalBudget = TotalBudget + Cohort(i) * Cost(j - i + 1)
                  End If
                  
                  Next i
            
            Next j
      
End If

BUDGETIMPACT = CDbl(TotalBudget)

End Function

Function getDimension(var As Variant) As Integer
On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    i = 0
    Do While True:
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function
Public Function ConvUnivariant(MultiArray As Variant, Optional RowNumber As Integer = 0) As Variant
'This function converts a bidimensional array into a unidimensional array
On Error Resume Next
If Not IsArray(MultiArray) Then Exit Function
If getDimension(MultiArray) = 1 Then Exit Function
If Not IsNull(MultiArray) Or IsEmpty(MultiArray) Then
    If RowNumber = 0 Then
    RowNumber = LBound(MultiArray, 1)
    End If
    Dim j As Integer
    Dim UniArray As Variant
    ReDim UniArray(LBound(MultiArray) To UBound(MultiArray))
    For j = LBound(MultiArray, 1) To UBound(MultiArray)
    UniArray(j) = MultiArray(j, RowNumber)
    Next
    ConvUnivariant = UniArray
End If
    Exit Function
End Function

Function ChangeArrayLength(OldArray As Variant, NewLen As Integer) As Variant
'On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    Dim NewArray As Variant
    ReDim NewArray(NewLen)
    
For i = 1 To NewLen

NewArray(i) = OldArray(i)

Next i

Err:
    ChangeArrayLength = NewArray
End Function


Function RangeIsVertical(rng As Range) As Boolean
    RangeIsVertical = IIf(rng.Columns.Count = 1, 1, 0)
End Function

Sub BIA()

End Sub
