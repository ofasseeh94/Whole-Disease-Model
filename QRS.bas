Attribute VB_Name = "QRS"

Option Explicit

Private Const QRS_MINIMUM_AGE As Double = 20
Private Const QRS_MAXIMUM_AGE As Double = 89

Private Const QRS_HF_WIDE_PROBABILITY As Double = 0.25
Private Const QRS_HF_WIDE_ADDITION As Double = 40
Private Const QRS_HF_NARROW_ADDITION As Double = 3

Private Const QRS_OSA_ADDITION As Double = 4

'Use a dedicated random-number column for baseline QRS HF subgroup assignment.
'Confirm this column is not already used elsewhere in the model.
Private Const QRS_RANDOM_COLUMN As Long = 51

Public Sub Load_Baseline_QRS()

    'Calculates baseline QRS once for every patient after:
    '   1. Patients() has been loaded
    '   2. RandArray has been generated
    
    'QRS is stored as a patient-level baseline characteristic.
    'It is not updated during model cycles.

    Dim i As Long

    For i = LBound(Patients) To UBound(Patients)

        With Patients(i)

            .QRS = QRS_Estimate(.Age, .BMI, .Female, .HF, .OSA, .ID)

        End With

    Next i

End Sub

Public Function QRS_Estimate( _
    ByVal Age As Double, _
    ByVal BMI As Double, _
    ByVal Female As Boolean, _
    ByVal BaselineHF As Boolean, _
    ByVal OSA As Boolean, _
    ByVal PatientID As Long) As Double

    'Estimates baseline QRS duration in milliseconds.
    
    'Baseline QRS source:
    'Rao ACA, Ng ACC, Sy RW, et al. Electrocardiographic QRS duration is
    'influenced by body mass index and sex. Int J Cardiol Heart Vasc.
    '2021;37:100884.
    
    'HF adjustment:
    'Applied only if the patient already has HF at baseline.
    'This is not updated later during cycle progression.
    
    'OSA adjustment:
    'Pressman GS, Orban M, Leinveber P, et al. Association between QRS duration
    'and obstructive sleep apnea. J Clin Sleep Med. 2012;8(6):649-654.

    Dim AgeCategory As Long
    Dim BMICategory As Long
    Dim BaseQRS As Double
    Dim HFAdjustment As Double
    Dim OSAAdjustment As Double

    If Age <= 0 Then Exit Function
    If BMI <= 0 Then Exit Function

    AgeCategory = QRS_Age_Category(Age)
    BMICategory = QRS_BMI_Category(BMI)

    BaseQRS = QRS_Base_Value(AgeCategory, BMICategory, Female)
    HFAdjustment = QRS_Baseline_HF_Adjustment(BaselineHF, PatientID)

    If OSA Then OSAAdjustment = QRS_OSA_ADDITION

    QRS_Estimate = BaseQRS + HFAdjustment + OSAAdjustment

End Function

Function QRS_Age_Category(ByVal Age As Double) As Long

    If Age < QRS_MINIMUM_AGE Then Age = QRS_MINIMUM_AGE
    If Age > QRS_MAXIMUM_AGE Then Age = QRS_MAXIMUM_AGE

    If Age < 30# Then
        QRS_Age_Category = 1
    ElseIf Age < 40# Then
        QRS_Age_Category = 2
    ElseIf Age < 50# Then
        QRS_Age_Category = 3
    ElseIf Age < 60# Then
        QRS_Age_Category = 4
    ElseIf Age < 70# Then
        QRS_Age_Category = 5
    ElseIf Age < 80# Then
        QRS_Age_Category = 6
    Else
        QRS_Age_Category = 7
    End If

End Function

Private Function QRS_BMI_Category(ByVal BMI As Double) As Long

    If BMI < 18.5 Then
        QRS_BMI_Category = 1
    ElseIf BMI < 25# Then
        QRS_BMI_Category = 2
    ElseIf BMI < 30# Then
        QRS_BMI_Category = 3
    Else
        QRS_BMI_Category = 4
    End If

End Function

Private Function QRS_Baseline_HF_Adjustment( _
    ByVal BaselineHF As Boolean, _
    ByVal PatientID As Long) As Double

    'Applies HF-related QRS widening only for patients who already have HF
    'at baseline.
    '
    'This uses RandArray because the model assigns a reproducible subgroup:
    '   25% of baseline HF patients receive +40 ms
    '   75% of baseline HF patients receive +3 ms

    Dim HFRandomNumber As Double

    If BaselineHF = False Then Exit Function

    HFRandomNumber = RandArray(PatientID, 1, 53)

    If HFRandomNumber <= QRS_HF_WIDE_PROBABILITY Then
        QRS_Baseline_HF_Adjustment = QRS_HF_WIDE_ADDITION
    Else
        QRS_Baseline_HF_Adjustment = QRS_HF_NARROW_ADDITION
    End If

End Function

Private Function QRS_Base_Value( _
    ByVal AgeCategory As Long, _
    ByVal BMICategory As Long, _
    ByVal Female As Boolean) As Double

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

