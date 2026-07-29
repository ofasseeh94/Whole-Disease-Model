Attribute VB_Name = "BuildStroke_References"
Public Sub Build_Stroke_Risk_Reference_Table()

    'Build age-group-specific reference values for the Hunter stroke risk equation.
    '
    'The Hunter equation standardizes SBP, DBP, and BMI:
    '   zSBP = (patient SBP - age-group mean SBP) / age-group SD SBP
    '   zDBP = (patient DBP - age-group mean DBP) / age-group SD DBP
    '   zBMI = (patient BMI - age-group mean BMI) / age-group SD BMI
    '
    'In this model, the reference means and SDs are calculated once from the baseline
    'Patient list sheet before the simulation run. They are then written as fixed values
    'to StrokeRisk. This avoids recalculating the reference population during the simulation,
    'which would incorrectly dilute treatment effects such as SBP reduction.

    Dim wsPatients As Worksheet
    Dim wsOut As Worksheet
    Dim outStart As Range

    Dim patientLastRow As Long
    Dim ageGroupIndex As Long

    Dim ageGroupLabel As String
    Dim minAge As Double
    Dim maxAge As Double

    'Point variables to the Excel sheets used by this calculation.
    Set wsPatients = ThisWorkbook.Sheets("Patient list")
    Set wsOut = ThisWorkbook.Sheets("StrokeRisk")
    Set outStart = wsOut.Range("B5")

    'Find the last patient row dynamically using the ID column.
    'This is not a fixed row number. It asks Excel to start at the bottom of column A
    'and move upward until it finds the last non-empty patient ID.
    patientLastRow = wsPatients.Cells(wsPatients.Rows.Count, "A").End(xlUp).row

    'Clear only the output area used by this table.
    wsOut.Range("B5:H9").ClearContents

    'Write table headers.
    outStart.Offset(0, 0).Value = "AgeGroup"
    outStart.Offset(0, 1).Value = "Mean_SBP"
    outStart.Offset(0, 2).Value = "SD_SBP"
    outStart.Offset(0, 3).Value = "Mean_DBP"
    outStart.Offset(0, 4).Value = "SD_DBP"
    outStart.Offset(0, 5).Value = "Mean_BMI"
    outStart.Offset(0, 6).Value = "SD_BMI"

    'Loop through the four Hunter age groups.
    For ageGroupIndex = 1 To 4

        Select Case ageGroupIndex

            Case 1
                ageGroupLabel = "<50"
                minAge = 0
                maxAge = 49.999999

            Case 2
                ageGroupLabel = "50-59"
                minAge = 50
                maxAge = 59.999999

            Case 3
                ageGroupLabel = "60-69"
                minAge = 60
                maxAge = 69.999999

            Case 4
                ageGroupLabel = "70+"
                minAge = 70
                maxAge = 200

        End Select

        'Write age group label.
        outStart.Offset(ageGroupIndex, 0).Value = ageGroupLabel

        'SBP is in Patient list column S.
        outStart.Offset(ageGroupIndex, 1).Value = MeanByAgeGroup(wsPatients, "S", minAge, maxAge, patientLastRow)
        outStart.Offset(ageGroupIndex, 2).Value = SDByAgeGroup(wsPatients, "S", minAge, maxAge, patientLastRow)

        'DBP is in Patient list column T.
        outStart.Offset(ageGroupIndex, 3).Value = MeanByAgeGroup(wsPatients, "T", minAge, maxAge, patientLastRow)
        outStart.Offset(ageGroupIndex, 4).Value = SDByAgeGroup(wsPatients, "T", minAge, maxAge, patientLastRow)

        'BMI is in Patient list column C.
        outStart.Offset(ageGroupIndex, 5).Value = MeanByAgeGroup(wsPatients, "C", minAge, maxAge, patientLastRow)
        outStart.Offset(ageGroupIndex, 6).Value = SDByAgeGroup(wsPatients, "C", minAge, maxAge, patientLastRow)

    Next ageGroupIndex

    'Create or refresh the named range used later by the stroke risk equation.
    'This name lets the VBA stroke equation call Range("Stroke_Risk_Reference_Table")
    'without caring where the table sits on the worksheet.
    ThisWorkbook.Names("Stroke_Risk_Reference_Table").Delete

    ThisWorkbook.Names.Add _
        name:="Stroke_Risk_Reference_Table", _
        RefersTo:=wsOut.Range("B5:H9")

End Sub

Private Function MeanByAgeGroup( _
    ByVal wsPatients As Worksheet, _
    ByVal valueColumn As String, _
    ByVal minAge As Double, _
    ByVal maxAge As Double, _
    ByVal patientLastRow As Long) As Double

    'Calculates the mean value for one patient variable within one age group.
    '
    'Age is read from column B.
    'The variable is read from valueColumn, for example:
    '   S = SBP
    '   T = DBP
    '   C = BMI

    Dim rowIndex As Long
    Dim total As Double
    Dim countPatients As Long
    Dim ageValue As Variant
    Dim patientValue As Variant

    For rowIndex = 2 To patientLastRow

        ageValue = wsPatients.Cells(rowIndex, "B").Value
        patientValue = wsPatients.Cells(rowIndex, valueColumn).Value

        If IsNumeric(ageValue) And IsNumeric(patientValue) Then

            If CDbl(ageValue) >= minAge And CDbl(ageValue) <= maxAge Then

                total = total + CDbl(patientValue)
                countPatients = countPatients + 1

            End If

        End If

    Next rowIndex

    If countPatients = 0 Then
        MeanByAgeGroup = 0
    Else
        MeanByAgeGroup = total / countPatients
    End If

End Function

Private Function SDByAgeGroup( _
    ByVal wsPatients As Worksheet, _
    ByVal valueColumn As String, _
    ByVal minAge As Double, _
    ByVal maxAge As Double, _
    ByVal patientLastRow As Long) As Double

    'Calculates the sample standard deviation for one patient variable within one age group.
    '
    'This uses the sample SD formula:
    '   sqrt(sum((x - mean)^2) / (n - 1))
    '
    'We use sample SD because the patient list is treated as a generated sample from
    'the target population, not the full true population.

    Dim rowIndex As Long
    Dim meanValue As Double
    Dim sumSquaredDifference As Double
    Dim countPatients As Long
    Dim ageValue As Variant
    Dim patientValue As Variant

    meanValue = MeanByAgeGroup(wsPatients, valueColumn, minAge, maxAge, patientLastRow)

    For rowIndex = 2 To patientLastRow

        ageValue = wsPatients.Cells(rowIndex, "B").Value
        patientValue = wsPatients.Cells(rowIndex, valueColumn).Value

        If IsNumeric(ageValue) And IsNumeric(patientValue) Then

            If CDbl(ageValue) >= minAge And CDbl(ageValue) <= maxAge Then

                sumSquaredDifference = sumSquaredDifference + (CDbl(patientValue) - meanValue) ^ 2
                countPatients = countPatients + 1

            End If

        End If

    Next rowIndex

    If countPatients < 2 Then
        SDByAgeGroup = 0
    Else
        SDByAgeGroup = Sqr(sumSquaredDifference / (countPatients - 1))
    End If

End Function
