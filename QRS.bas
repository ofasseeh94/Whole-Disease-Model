Attribute VB_Name = "QRS"

Option Explicit

Private Const QRS_MINIMUM_AGE As Double = 20
Private Const QRS_MAXIMUM_AGE As Double = 89

Private Const QRS_HF_WIDE_PROBABILITY As Double = 0.25
Private Const QRS_HF_WIDE_ADDITION As Double = 40
Private Const QRS_HF_NARROW_ADDITION As Double = 3

Private Const QRS_OSA_ADDITION As Double = 4

Public Sub Initialize_QRS()

    'Initializes QRS for every loaded patient after:
    '   1. Patients() has been loaded into VBA memory
    '   2. RandArray has been generated
    
    'QRS is stored in Patient.QRS
    'This is only the initial QRS value after loading .QRS is recalculated later
    'inside Characteristics_Progression as patient characteristics change.

    Dim i As Long

    For i = LBound(Patients) To UBound(Patients)

        Call Update_QRS(Patients(i))

    Next i

End Sub

Public Sub Update_QRS(ByRef Patient As Patient)

    'Recalculates QRS for the active patient using the characteristics currently
    'stored in the Patient data type
    
    'This updates only Patient.QRS
    
    'When used in Characteristics_Progression, this should be called before HF
    'incidence is evaluated because ProbHF uses Patient.QRS

    Patient.QRS = QRS_Estimate(Patient)

End Sub

Public Function QRS_Estimate(ByRef Patient As Patient) As Double

    'Estimates QRS duration in milliseconds using current patient characteristics
    
    'source:Rao ACA, Ng ACC, Sy RW, et al. Electrocardiographic QRS duration is
    'influenced by body mass index and sex. Int J Cardiol Heart Vasc.
    '2021;37:100884.
    
    'HF adjustment:
    'Applied when the patient has HF at the time QRS is calculated
    'The 25% / 75% HF subgroup assignment uses a fixed patient-level random
    'draw from RandArray column 53, so the patient's HF-related QRS subgroup
    'is stable across cycles
    
    'OSA adjustment:
    'Pressman GS, Orban M, Leinveber P, et al. Association between QRS duration
    'and obstructive sleep apnea. J Clin Sleep Med. 2012;8(6):649-654

    Dim AgeCategory As Single
    Dim BMICategory As Single
    Dim BaseQRS As Single
    Dim HFAdjustment As Single
    Dim OSAAdjustment As Single

    With Patient

        If .Age <= 0 Then Exit Function
        If .BMI <= 0 Then Exit Function
        If .ID <= 0 Then Exit Function

        AgeCategory = QRS_Age_Category(.Age)
        BMICategory = QRS_BMI_Category(.BMI)

        BaseQRS = QRS_Base_Value(AgeCategory, BMICategory, .Female)
        HFAdjustment = QRS_HF_Adjustment(Patient)

        If .OSA Then OSAAdjustment = QRS_OSA_ADDITION

    End With

    QRS_Estimate = BaseQRS + HFAdjustment + OSAAdjustment

End Function

Private Function QRS_Age_Category(ByVal Age As Double) As Long

    'Converts age into the age bands used by the QRS reference table
    'Values outside the source range are capped to the closest supported age

    If Age < QRS_MINIMUM_AGE Then Age = QRS_MINIMUM_AGE
    If Age > QRS_MAXIMUM_AGE Then Age = QRS_MAXIMUM_AGE

    If Age < 30 Then
        QRS_Age_Category = 1
    ElseIf Age < 40 Then
        QRS_Age_Category = 2
    ElseIf Age < 50 Then
        QRS_Age_Category = 3
    ElseIf Age < 60 Then
        QRS_Age_Category = 4
    ElseIf Age < 70 Then
        QRS_Age_Category = 5
    ElseIf Age < 80 Then
        QRS_Age_Category = 6
    Else
        QRS_Age_Category = 7
    End If

End Function

Function QRS_BMI_Category(ByVal BMI As Double) As Long

    'Converts BMI into the BMI bands used by the QRS reference table.

    If BMI < 18.5 Then
        QRS_BMI_Category = 1
    ElseIf BMI < 25 Then
        QRS_BMI_Category = 2
    ElseIf BMI < 30 Then
        QRS_BMI_Category = 3
    Else
        QRS_BMI_Category = 4
    End If

End Function

Function QRS_HF_Adjustment(ByRef Patient As Patient) As Double

    'Applies HF-related QRS widening when the patient has HF at the time QRS is calculated.
    
    'RandArray column 53 is used for this fixed patient-level subgroup draw:
    '   25% of HF patients receive +40 ms
    '   75% of HF patients receive +3 ms
    
    'The cycle index is fixed at 1 so the same patient remains in the same
    'HF-QRS subgroup across all cycles.

    Dim HFRandomNumber As Double

    With Patient

        If .HF = False Then Exit Function

        HFRandomNumber = RandArray(.ID, 1, 53)

    End With

    If HFRandomNumber <= QRS_HF_WIDE_PROBABILITY Then
        QRS_HF_Adjustment = QRS_HF_WIDE_ADDITION
    Else
        QRS_HF_Adjustment = QRS_HF_NARROW_ADDITION
    End If

End Function
Private Function QRS_Base_Value( _
    ByVal AgeCategory As Long, _
    ByVal BMICategory As Long, _
    ByVal Female As Boolean) As Double

    'Returns the QRS reference value by sex, age category, and BMI category.
    '
    'Age categories:
    '   1 = 20-29
    '   2 = 30-39
    '   3 = 40-49
    '   4 = 50-59
    '   5 = 60-69
    '   6 = 70-79
    '   7 = 80-89
    '
    'BMI categories:
    '   1 = BMI < 18.5
    '   2 = BMI 18.5 to <25
    '   3 = BMI 25 to <30
    '   4 = BMI >=30

    If Female Then

        Select Case AgeCategory

            Case 1
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 82.8
                    Case 2: QRS_Base_Value = 83.4
                    Case 3: QRS_Base_Value = 83.9
                    Case 4: QRS_Base_Value = 85.6
                End Select

            Case 2
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 80.1
                    Case 2: QRS_Base_Value = 83.1
                    Case 3: QRS_Base_Value = 83.6
                    Case 4: QRS_Base_Value = 85#
                End Select

            Case 3
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 80.5
                    Case 2: QRS_Base_Value = 82.8
                    Case 3: QRS_Base_Value = 83.5
                    Case 4: QRS_Base_Value = 84.6
                End Select

            Case 4
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 80.3
                    Case 2: QRS_Base_Value = 82.6
                    Case 3: QRS_Base_Value = 83.2
                    Case 4: QRS_Base_Value = 84.5
                End Select

            Case 5
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 80.8
                    Case 2: QRS_Base_Value = 82.8
                    Case 3: QRS_Base_Value = 83.5
                    Case 4: QRS_Base_Value = 84.2
                End Select

            Case 6
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 81.2
                    Case 2: QRS_Base_Value = 82.6
                    Case 3: QRS_Base_Value = 82.9
                    Case 4: QRS_Base_Value = 84#
                End Select

            Case 7
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 80.1
                    Case 2: QRS_Base_Value = 82.4
                    Case 3: QRS_Base_Value = 82.7
                    Case 4: QRS_Base_Value = 89.9
                End Select

        End Select

    Else

        Select Case AgeCategory

            Case 1
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 91.3
                    Case 2: QRS_Base_Value = 94.4
                    Case 3: QRS_Base_Value = 94#
                    Case 4: QRS_Base_Value = 94.3
                End Select

            Case 2
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 87.7
                    Case 2: QRS_Base_Value = 92.3
                    Case 3: QRS_Base_Value = 93.2
                    Case 4: QRS_Base_Value = 93.1
                End Select

            Case 3
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 90.5
                    Case 2: QRS_Base_Value = 91.8
                    Case 3: QRS_Base_Value = 92.4
                    Case 4: QRS_Base_Value = 92.9
                End Select

            Case 4
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 86.5
                    Case 2: QRS_Base_Value = 90.2
                    Case 3: QRS_Base_Value = 91.5
                    Case 4: QRS_Base_Value = 91.8
                End Select

            Case 5
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 83#
                    Case 2: QRS_Base_Value = 89.5
                    Case 3: QRS_Base_Value = 90.7
                    Case 4: QRS_Base_Value = 91.4
                End Select

            Case 6
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 86.1
                    Case 2: QRS_Base_Value = 89.4
                    Case 3: QRS_Base_Value = 90.5
                    Case 4: QRS_Base_Value = 90.9
                End Select

            Case 7
                Select Case BMICategory
                    Case 1: QRS_Base_Value = 88.5
                    Case 2: QRS_Base_Value = 89.5
                    Case 3: QRS_Base_Value = 89.6
                    Case 4: QRS_Base_Value = 89.9
                End Select

        End Select

    End If

End Function

