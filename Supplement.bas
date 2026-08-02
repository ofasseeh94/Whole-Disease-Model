Attribute VB_Name = "Supplement"
Public Function LN(ByVal Q As Double)
   LN = Application.WorksheetFunction.LN(Q)
End Function
Function ArrayCountIf(arr As Variant, str As String) As Integer
     
    Dim i As Integer, iCount As Integer
    For i = LBound(arr) To UBound(arr)
     
        If arr(i) = str Then iCount = iCount + 1
     
    Next i
     
    ArrayCountIf = iCount
End Function

Function VBASumif(ByVal arr As Variant, ByVal criteria As Variant, ByVal criteriaColNo As Integer, ByVal sumColNo As Integer) As Double
    For i = LBound(arr) To UBound(arr)
        If arr(i, criteriaColNo) = criteria Then VBASumif = VBASumif + arr(i, sumColNo)
    Next i
End Function

Function Product_Inv_NotZero(Input_Array() As Double) As Double
'this function calculates the product of all values in a univariate array if it is not zero
'equation is = 1 - Product(1-X)

Dim i As Long

Product_Inv_NotZero = 1

      For i = LBound(Input_Array) To UBound(Input_Array)
            
            If Input_Array(i) <> 0 Then Product_Inv_NotZero = (1 - Input_Array(i)) * Product_Inv_NotZero
            
      Next i

      Product_Inv_NotZero = 1 - Product_Inv_NotZero
            
End Function

Function Product_NotZero(Input_Array() As Double) As Double
'this function calculates the product of all values in a univariate array if it is not zero
'equation is = 1 - Product(1-X)

Dim i As Long

Product_NotZero = 1

      For i = LBound(Input_Array) To UBound(Input_Array)
            
            If Input_Array(i) <> 0 Then Product_NotZero = (Input_Array(i)) * Product_NotZero
            
      Next i
            
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

'Covert Multivariate array into a Univariate
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

'Merge two arrays
'Merges two arrays in a series after converting them into univariate arrays. Needs another function to run.
Public Function MergeArrays(FirstArrayB As Variant, SecondArrayB As Variant) As Variant
'This function compare values in 2 arrays and return back an array with only the unique values
If IsNull(FirstArrayB) Or IsNull(SecondArrayB) Or IsEmpty(FirstArrayB) Or IsEmpty(SecondArrayB) Then
    If IsNull(SecondArrayB) Then
    UniqueArrays = EliminateEmpty(FirstArrayB)
    ElseIf IsNull(FirstArrayB) Then
    UniqueArrays = EliminateEmpty(SecondArrayB)
    End If
GoTo Ending
End If

'FirstArray = (ConvUnivariant(CVar(FirstArrayB)))
FirstArray = CVar(FirstArrayB)
SecondArray = CVar(SecondArrayB)

Dim z As Variant
ReDim z(UBound(FirstArray) + UBound(SecondArray) + 1 + UBound(SecondArray) + 1)

Dim j As Integer
For j = LBound(FirstArray) To UBound(FirstArray)
z(j) = FirstArray(j)
Next
For j = UBound(FirstArray) + 1 To UBound(FirstArray) + UBound(SecondArray) + 1
z(j) = SecondArray(j - UBound(FirstArray) - 1 + LBound(SecondArray))
Next
For j = UBound(FirstArray) + UBound(SecondArray) + 1 + 1 To UBound(FirstArray) + UBound(SecondArray) + 1 + UBound(SecondArray) + 1
z(j) = SecondArray(j - UBound(FirstArray) - UBound(SecondArray) - 1 - 1 + LBound(SecondArray))
Next

MergeArrays = z

Ending:
End Function

'Remove duplicates from an array
Public Function EliminateDuplicate(poArr As Variant, Optional KeepOne As Boolean = True) As Variant
Dim poArrNoDup()
Dim UpperLimit As Integer
If Not IsArray(poArr) Then Exit Function
If IsNull(poArr) Then Exit Function
If IsEmpty(poArr) Then Exit Function
dupArrIndex = -1
For i = LBound(poArr) To UBound(poArr)
dupBool = False
If KeepOne = True Then UpperLimit = i Else UpperLimit = UBound(poArr)
        For j = LBound(poArr) To UpperLimit
            If poArr(i) = poArr(j) And Not i = j Then
                dupBool = True
            End If
        Next j
        If dupBool = False Then
            dupArrIndex = dupArrIndex + 1
            ReDim Preserve poArrNoDup(dupArrIndex)
            poArrNoDup(dupArrIndex) = poArr(i)
        End If
Next i
EliminateDuplicate = EliminateEmpty(poArrNoDup)
End Function
'Filter array
Function FilterArray(OriginalArray As Variant, ColCrit1 As Integer, rCrit1 As String, Optional ColCrit2 As Integer, Optional rCrit2 As String, Optional ColCrit3 As Integer, Optional rCrit3 As String, Optional ColCrit4 As Integer, Optional rCrit4 As String) As Variant
'this function is intended to filter a multidimensional array based on upto 4 criteria. It needs the range, criteria and criteria columns

Dim i As Long
Dim j As Long
Dim c As Long
Dim LastColumn As Long
Dim LastRow As Long
Dim FilteredArray As Variant

'Initial dimensions of the array
LastColumn = UBound(OriginalArray, 2)
LastRow = UBound(OriginalArray)
ReDim FilteredArray(1 To LastRow, 1 To (LastColumn))

'Get the length of total questions with the topic
c = 1
j = 1

If ColCrit2 = 0 Then
      
      'In case of 1 criteria only
      For i = UBound(OriginalArray) To 1 Step -1
          If OriginalArray(i, ColCrit1) = rCrit1 Then
                  For c = 1 To LastColumn
                      FilteredArray(j, c) = OriginalArray(i, c)
                  Next c
                   j = j + 1
          End If
      Next i

ElseIf ColCrit3 = 0 Then
      
      'In case of 2 criteria only
      For i = UBound(OriginalArray) To 1 Step -1
          If OriginalArray(i, ColCrit1) = rCrit1 And OriginalArray(i, ColCrit2) = rCrit2 Then
                  For c = 1 To LastColumn
                      FilteredArray(j, c) = OriginalArray(i, c)
                  Next c
                   j = j + 1
          End If
      Next i

ElseIf ColCrit4 = 0 Then
      
      'In case of 3 criteria
      For i = UBound(OriginalArray) To 1 Step -1
          If OriginalArray(i, ColCrit1) = rCrit1 And OriginalArray(i, ColCrit2) = rCrit2 And OriginalArray(i, ColCrit3) = rCrit3 Then
                  For c = 1 To LastColumn
                      FilteredArray(j, c) = OriginalArray(i, c)
                  Next c
                   j = j + 1
          End If
      Next i

Else

      'In case of 4 criteria
      For i = UBound(OriginalArray) To 1 Step -1
          If OriginalArray(i, ColCrit1) = rCrit1 And OriginalArray(i, ColCrit2) = rCrit2 And OriginalArray(i, ColCrit3) = rCrit3 And OriginalArray(i, ColCrit4) = rCrit4 Then
                  For c = 1 To LastColumn
                      FilteredArray(j, c) = OriginalArray(i, c)
                  Next c
                   j = j + 1
          End If
      Next i

End If
FilterArray = ReDimPreserve(FilteredArray, j - 1, LastColumn)

End Function

Private Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long, Optional ArrayName As String = "") As Variant

    Dim i, j As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants. MsgBox UCase(TypeName(MyArray))
'---------------------------------------------------------------
'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
    If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
        MsgBox "This function only works if your array is a Variant Data Type." & vbNewLine & _
               "You have two choice:" & vbNewLine & _
               " 1) Change your array to a Variant and try again." & vbNewLine & _
               " 2) Change the DataType of TempArray to match your array and comment the top block out of the function ReDimPreserve" _
                , vbCritical, "Invalid Array Data Type"
        End
    End If
'---------------------------------------------------------------
    ReDimPreserve = False
    'check if its in array first
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End
    
    'get old lBound/uBound
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array
    If nNewFirstUBound = 0 Then Stop
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For j = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way
            If nOldFirstUBound >= i And nOldLastUBound >= j Then
                TempArray(i, j) = MyArray(i, j)
            End If
        Next
    Next
    'return the array redimmed
    If IsArray(TempArray) Then ReDimPreserve = TempArray
End Function

'Shuffle array
Function ShuffleArray(UnShuffledArray As Variant) As Variant()
' ShuffleArray
' This function returns the values of UnShuffledArray in random order. The original
' UnShuffledArray is not modified.

    Dim N As Long
    Dim temp As Variant
    Dim j As Long
    Dim arr() As Variant
    
    Randomize
    L = UBound(UnShuffledArray) - LBound(UnShuffledArray) + 1
    ReDim arr(LBound(UnShuffledArray) To UBound(UnShuffledArray))
    For N = LBound(UnShuffledArray) To UBound(UnShuffledArray)
        arr(N) = UnShuffledArray(N)
    Next N
    For N = LBound(UnShuffledArray) To UBound(UnShuffledArray)
        j = CLng(((UBound(UnShuffledArray) - N) * Rnd) + N)
        temp = arr(N)
        arr(N) = arr(j)
        arr(j) = temp
    Next N
    ShuffleArray = arr

End Function

'Remove empty values from an array
Public Function EliminateEmpty(EmptyArray As Variant) As Variant
'Remove empty values and reset array to start from 0
If Not IsArray(EmptyArray) Then Exit Function
If IsNull(EmptyArray) Then Exit Function
If IsEmpty(EmptyArray) Then Exit Function
On Error Resume Next
j = -1
ReDim poArrNoEmpty(0 To UBound(EmptyArray))
For i = 0 To UBound(EmptyArray)
    If EmptyArray(i) <> "" Then
        j = j + 1
        poArrNoEmpty(j) = EmptyArray(i)
    End If
Next i
ReDim Preserve poArrNoEmpty(0 To j)
EliminateEmpty = poArrNoEmpty
Exit Function


'Reduces the length of an array by trimming from bottom
Function ChangeArrayLength(OldArray As Variant, NewLen As Integer) As Variant
'On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    Dim NewArray As Variant
'Define the length of the new array
    ReDim NewArray(NewLen)
'Assign values of new array according to old array
For i = 1 To NewLen
NewArray(i) = OldArray(i)
Next i
Err:
    ChangeArrayLength = NewArray
End Function

Function TransposeArray(MyArray As Variant) As Variant
    Dim x As Long, Y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    Dim tempArr As Variant
    
    'Get Upper and Lower Bounds
    maxX = UBound(MyArray, 1)
    minX = LBound(MyArray, 1)
    maxY = UBound(MyArray, 2)
    minY = LBound(MyArray, 2)
    
    'Create New Temp Array
    ReDim tempArr(minY To maxY, minX To maxX)
    
    'Transpose the Array
    For x = minX To maxX
        For Y = minY To maxY
            tempArr(Y, x) = MyArray(x, Y)
        Next Y
    Next x
    
    'Output Array
    TransposeArray = tempArr
    
End Function

Function GenerateCSV(FileName As String)

Dim Path As String

Path = ActiveWorkbook.Path

If FileExist(Path & "\" & FileName & ".csv") = False Then

        With ActiveWorkbook

        .SaveAs FileName:=Path & "\" & FileName, FileFormat:=xlCSV
        '.Close False
        
        End With
        
        GenerateCSV = Path & "\" & FileName & ".csv"
Else

End If

End Function

Function FileExist(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    FileExist = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function
Function GetFilenameFromPath(ByVal strPath As String) As String
Dim FileName As String
' Returns the rightmost characters of a string upto but not including the rightmost '\' and removes the extension
' e.g. 'c:\winnt\win.ini' returns 'win'
FileName = Split(strPath, "\")(UBound(Split(strPath, "\")))
GetFilenameFromPath = Left(FileName, (InStrRev(FileName, ".", -1, vbTextCompare) - 1))
End Function


Function WorkbookOpen(strWorkBookName As String) As Boolean
    'Returns TRUE if the workbook is open
    Dim oXL As Excel.Application
    Dim oBk As Workbook

    On Error Resume Next
    Set oXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        'Excel is NOT open, so the workbook cannot be open
        Err.Clear
        WorkbookOpen = False
    Else
        'Excel is open, check if workbook is open
        Set oBk = oXL.Workbooks(strWorkBookName)
        If oBk Is Nothing Then
            WorkbookOpen = False
        Else
            WorkbookOpen = True
            Set oBk = Nothing
        End If
    End If
    Set oXL = Nothing
End Function
Function onlyDigits(s As Variant) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function
' Shell sort algorithm for sorting a double from largest to smallest.
' Adopted from "Numerical Recipes in C" aka NRC 2nd Edition p330ff.
' Speed is on the range of N^1.25 to N^1.5 (somewhere between bubble and quicksort)
' Refer to the NRC reference for more details on efficiency.
'
Sub ShellSortDescending(ByRef a() As Double)

    ' requires a(1..N)
      Dim N As Integer
      Dim lb As Long, ub As Long, nElements As Long
     
    N = UBound(a)

    ' setup
    lb = LBound(a)
    ub = UBound(a)
    nElements = ub - lb + 1
    Dim i, j, inc As Integer
    Dim v As Double
    inc = 0

    ' determine the starting incriment
    Do
        inc = inc * 3 + 1
        
    Loop While inc <= N

    ' loop over the partial sorts

    Do
        inc = inc / 3

        ' Outer loop of straigh insertion

        For i = inc To N
            v = a(i)
            j = i

            ' Inner loop of straight insertion
            ' switch to a(j - inc) > v for ascending

            Do While a(j - inc) < v
                a(j) = a(j - inc)
                j = j - inc
                If j <= inc Then Exit Do
            Loop
            a(j) = v
        Next i
        
    Loop While inc > 1

End Sub

Function Build_Stroke_Reference_Array() As Variant

    'Creates a temporary in-memory matrix from Patient_Cohort_Matrix
    
    'Patient_Cohort_Matrix columns:
    '   Column 2  = Age
    '   Column 3  = BMI
    '   Column 19 = SBP
    '   Column 20 = DBP
    
    'Temporary array columns:
    '   Column 1 = age group string
    '   Column 2 = SBP
    '   Column 3 = DBP
    '   Column 4 = BMI

    Dim StrokeReferenceArray As Variant
    Dim i As Long
    Dim PatientRow As Long

    ReDim StrokeReferenceArray(1 To NPatients, 1 To 4)

    For i = 1 To NPatients

        PatientRow = i + 1

        If Patient_Cohort_Matrix(PatientRow, 2) < 50 Then
            StrokeReferenceArray(i, 1) = "<50"
        ElseIf Patient_Cohort_Matrix(PatientRow, 2) < 60 Then
            StrokeReferenceArray(i, 1) = "50-59"
        ElseIf Patient_Cohort_Matrix(PatientRow, 2) < 70 Then
            StrokeReferenceArray(i, 1) = "60-69"
        Else
            StrokeReferenceArray(i, 1) = "70+"
        End If

        StrokeReferenceArray(i, 2) = Patient_Cohort_Matrix(PatientRow, 19)
        StrokeReferenceArray(i, 3) = Patient_Cohort_Matrix(PatientRow, 20)
        StrokeReferenceArray(i, 4) = Patient_Cohort_Matrix(PatientRow, 3)

    Next i

    Build_Stroke_Reference_Array = StrokeReferenceArray

End Function

Function Stroke_Reference_Mean(StrokeReferenceArray As Variant, AgeGroup As String, ValueColumn As Integer) As Double

    'Filters the temporary stroke reference array by age group, extracts one
    'variable column, and returns the mean.

    Dim FilteredArray As Variant
    Dim ValueArray As Variant

    If Stroke_Reference_Count(StrokeReferenceArray, AgeGroup) = 0 Then
        Stroke_Reference_Mean = 0
        Exit Function
    End If

    FilteredArray = FilterArray(StrokeReferenceArray, 1, AgeGroup)
    ValueArray = ConvUnivariant(FilteredArray, ValueColumn)

    Stroke_Reference_Mean = Application.WorksheetFunction.Average(ValueArray)

End Function

Function Stroke_Reference_SD(StrokeReferenceArray As Variant, AgeGroup As String, ValueColumn As Integer) As Double

    'Filters the temporary stroke reference array by age group, extracts one
    'variable column, and returns the sample standard deviation.

    Dim FilteredArray As Variant
    Dim ValueArray As Variant

    If Stroke_Reference_Count(StrokeReferenceArray, AgeGroup) < 2 Then
        Stroke_Reference_SD = 0
        Exit Function
    End If

    FilteredArray = FilterArray(StrokeReferenceArray, 1, AgeGroup)
    ValueArray = ConvUnivariant(FilteredArray, ValueColumn)

    Stroke_Reference_SD = Application.WorksheetFunction.StDev_S(ValueArray)

End Function

Function Stroke_Reference_Count(StrokeReferenceArray As Variant, AgeGroup As String) As Long

    'Counts patients in each age group before calling FilterArray.

    Dim i As Long

    For i = LBound(StrokeReferenceArray, 1) To UBound(StrokeReferenceArray, 1)

        If StrokeReferenceArray(i, 1) = AgeGroup Then
            Stroke_Reference_Count = Stroke_Reference_Count + 1
        End If

    Next i

End Function
