Attribute VB_Name = "QRS"
Option Explicit

'===============================================================================
' Function: QRS_Estimate
'
' Purpose:
'   Estimates QRS duration in milliseconds using:
'       1. Age
'       2. BMI
'       3. Sex
'       4. Heart failure status
'       5. Obstructive sleep apnoea status
'
' Baseline QRS source:
'   Rao ACA, Ng ACC, Sy RW, et al.
'   "Electrocardiographic QRS duration is influenced by body mass index
'   and sex."
'   International Journal of Cardiology Heart & Vasculature.
'   2021;37:100884.
'
' Heart failure adjustment:
'   Approximately 25% of hospitalized heart failure patients were reported
'   to have QRS duration >=120 ms.
'
'   In this model:
'       25% of HF patients receive +40 ms
'       75% of HF patients receive +3 ms
'
'   The 25% probability is literature-based.
'   The +40 ms and +3 ms values are modelling assumptions.
'
' OSA adjustment:
'   Pressman GS, Orban M, Leinveber P, et al.
'   "Association between QRS duration and obstructive sleep apnea."
'   Journal of Clinical Sleep Medicine. 2012;8(6):649-654.
'
'   Mean QRS values reported:
'       No OSA:                85 ms
'       Mild-to-moderate OSA: 89 ms
'
'   Therefore, binary OSA receives an addition of:
'       89 - 85 = 4 ms
'
' Inputs:
'   Age:
'       Age in years.
'       Ages below 20 are treated as age 20.
'       Ages above 89 are treated as age 89.
'
'   BMI:
'       Body mass index in kg/m^2.
'
'   Female:
'       True  = female
'       False = male
'
'   HF:
'       True  = heart failure present
'       False = heart failure absent
'
'   OSA:
'       True  = obstructive sleep apnoea present
'       False = obstructive sleep apnoea absent
'
'   ID:
'       Patient identifier passed to RandArray.
'
'       RandArray(ID, 1, 1) is assumed to return a reproducible random number
'       between 0 and 1.
'
' Return:
'   Estimated QRS duration in milliseconds.
'
' Example:
'   =QRS_Estimate(65,32,TRUE,TRUE,TRUE,1001)
'
' Important limitation:
'   This is a population-level estimate and does not replace direct ECG
'   measurement. Bundle branch block, pacing, pre-excitation, electrolyte
'   disturbances and other conduction abnormalities are not explicitly modelled.
'===============================================================================

Public Function QRS_Estimate( _
    ByVal Age As Single, _
    ByVal BMI As Single, _
    ByVal Female As Boolean, _
    ByVal HF As Boolean, _
    ByVal OSA As Boolean, _
    ByVal ID As Variant) As Variant

    '===========================================================================
    ' Local fixed model parameters
    '
    ' These constants are declared inside the function so VBA can always find
    ' them, regardless of the module structure.
    '===========================================================================

    Const QRS_MINIMUM_AGE As Single = 20
    Const QRS_MAXIMUM_AGE As Single = 89

    Const QRS_HF_WIDE_PROBABILITY As Double = 0.25
    Const QRS_HF_WIDE_ADDITION As Double = 40
    Const QRS_HF_NARROW_ADDITION As Double = 3

    Const QRS_OSA_ADDITION As Double = 4

    '===========================================================================
    ' Calculation variables
    '===========================================================================

    Dim EffectiveAge As Single
    Dim AgeCategory As Long
    Dim BMICategory As Long
    Dim BaseQRS As Double
    Dim HFAdjustment As Double
    Dim OSAAdjustment As Double
    Dim HFRandomNumber As Double

    On Error GoTo ErrorHandler

    '---------------------------------------------------------------------------
    ' Validate BMI
    '---------------------------------------------------------------------------

    If BMI <= 0 Then
        QRS_Estimate = CVErr(xlErrNA)
        Exit Function
    End If

    '---------------------------------------------------------------------------
    ' Normalize age
    '
    ' Ages below 20 are treated as age 20.
    ' Ages above 89 are treated as age 89.
    '---------------------------------------------------------------------------

    EffectiveAge = Age

    If EffectiveAge < QRS_MINIMUM_AGE Then

        EffectiveAge = QRS_MINIMUM_AGE

    ElseIf EffectiveAge > QRS_MAXIMUM_AGE Then

        EffectiveAge = QRS_MAXIMUM_AGE

    End If

    '---------------------------------------------------------------------------
    ' Assign age category
    '
    '   1 = 20-29
    '   2 = 30-39
    '   3 = 40-49
    '   4 = 50-59
    '   5 = 60-69
    '   6 = 70-79
    '   7 = 80-89
    '---------------------------------------------------------------------------

    Select Case EffectiveAge

        Case Is < 30
            AgeCategory = 1

        Case Is < 40
            AgeCategory = 2

        Case Is < 50
            AgeCategory = 3

        Case Is < 60
            AgeCategory = 4

        Case Is < 70
            AgeCategory = 5

        Case Is < 80
            AgeCategory = 6

        Case Else
            AgeCategory = 7

    End Select

    '---------------------------------------------------------------------------
    ' Assign BMI category
    '
    '   1 = BMI <18.5
    '   2 = BMI 18.5 to <25
    '   3 = BMI 25 to <30
    '   4 = BMI >=30
    '---------------------------------------------------------------------------

    Select Case BMI

        Case Is < 18.5
            BMICategory = 1

        Case Is < 25
            BMICategory = 2

        Case Is < 30
            BMICategory = 3

        Case Else
            BMICategory = 4

    End Select

    '---------------------------------------------------------------------------
    ' Retrieve baseline QRS duration according to sex, age category and BMI
    '---------------------------------------------------------------------------

    If Female Then

        '=======================================================================
        ' FEMALE BASELINE QRS DURATION, ms
        '=======================================================================

        Select Case AgeCategory

            Case 1      'Age 20-29

                Select Case BMICategory
                    Case 1: BaseQRS = 82.8
                    Case 2: BaseQRS = 83.4
                    Case 3: BaseQRS = 83.9
                    Case 4: BaseQRS = 85.6
                End Select

            Case 2      'Age 30-39

                Select Case BMICategory
                    Case 1: BaseQRS = 80.1
                    Case 2: BaseQRS = 83.1
                    Case 3: BaseQRS = 83.6
                    Case 4: BaseQRS = 85#
                End Select

            Case 3      'Age 40-49

                Select Case BMICategory
                    Case 1: BaseQRS = 80.5
                    Case 2: BaseQRS = 82.8
                    Case 3: BaseQRS = 83.5
                    Case 4: BaseQRS = 84.6
                End Select

            Case 4      'Age 50-59

                Select Case BMICategory
                    Case 1: BaseQRS = 80.3
                    Case 2: BaseQRS = 82.6
                    Case 3: BaseQRS = 83.2
                    Case 4: BaseQRS = 84.5
                End Select

            Case 5      'Age 60-69

                Select Case BMICategory
                    Case 1: BaseQRS = 80.8
                    Case 2: BaseQRS = 82.8
                    Case 3: BaseQRS = 83.5
                    Case 4: BaseQRS = 84.2
                End Select

            Case 6      'Age 70-79

                Select Case BMICategory
                    Case 1: BaseQRS = 81.2
                    Case 2: BaseQRS = 82.6
                    Case 3: BaseQRS = 82.9
                    Case 4: BaseQRS = 84#
                End Select

            Case 7      'Age 80-89

                Select Case BMICategory
                    Case 1: BaseQRS = 80.1
                    Case 2: BaseQRS = 82.4
                    Case 3: BaseQRS = 82.7
                    Case 4: BaseQRS = 89.9
                End Select

        End Select

    Else

        '=======================================================================
        ' MALE BASELINE QRS DURATION, ms
        '=======================================================================

        Select Case AgeCategory

            Case 1      'Age 20-29

                Select Case BMICategory
                    Case 1: BaseQRS = 91.3
                    Case 2: BaseQRS = 94.4
                    Case 3: BaseQRS = 94#
                    Case 4: BaseQRS = 94.3
                End Select

            Case 2      'Age 30-39

                Select Case BMICategory
                    Case 1: BaseQRS = 87.7
                    Case 2: BaseQRS = 92.3
                    Case 3: BaseQRS = 93.2
                    Case 4: BaseQRS = 93.1
                End Select

            Case 3      'Age 40-49

                Select Case BMICategory
                    Case 1: BaseQRS = 90.5
                    Case 2: BaseQRS = 91.8
                    Case 3: BaseQRS = 92.4
                    Case 4: BaseQRS = 92.9
                End Select

            Case 4      'Age 50-59

                Select Case BMICategory
                    Case 1: BaseQRS = 86.5
                    Case 2: BaseQRS = 90.2
                    Case 3: BaseQRS = 91.5
                    Case 4: BaseQRS = 91.8
                End Select

            Case 5      'Age 60-69

                Select Case BMICategory
                    Case 1: BaseQRS = 83#
                    Case 2: BaseQRS = 89.5
                    Case 3: BaseQRS = 90.7
                    Case 4: BaseQRS = 91.4
                End Select

            Case 6      'Age 70-79

                Select Case BMICategory
                    Case 1: BaseQRS = 86.1
                    Case 2: BaseQRS = 89.4
                    Case 3: BaseQRS = 90.5
                    Case 4: BaseQRS = 90.9
                End Select

            Case 7      'Age 80-89

                Select Case BMICategory
                    Case 1: BaseQRS = 88.5
                    Case 2: BaseQRS = 89.5
                    Case 3: BaseQRS = 89.6
                    Case 4: BaseQRS = 89.9
                End Select

        End Select

    End If

    '---------------------------------------------------------------------------
    ' Apply heart failure adjustment
    '
    ' RandArray is called only once so the patient remains in the same HF
    ' subgroup during the current function call.
    '---------------------------------------------------------------------------

    HFAdjustment = 0#

    If HF Then

        HFRandomNumber = CDbl(RandArray(ID, 1, 1))

        'Ensure RandArray returned a valid probability.
        If HFRandomNumber < 0# Or HFRandomNumber > 1# Then
            QRS_Estimate = CVErr(xlErrNum)
            Exit Function
        End If

        If HFRandomNumber <= QRS_HF_WIDE_PROBABILITY Then

            '25% of HF patients
            HFAdjustment = QRS_HF_WIDE_ADDITION

        Else

            'Remaining 75% of HF patients
            HFAdjustment = QRS_HF_NARROW_ADDITION

        End If

    End If

    '---------------------------------------------------------------------------
    ' Apply obstructive sleep apnoea adjustment
    '---------------------------------------------------------------------------

    OSAAdjustment = 0#

    If OSA Then
        OSAAdjustment = QRS_OSA_ADDITION
    End If

    '---------------------------------------------------------------------------
    ' Return final estimated QRS duration
    '---------------------------------------------------------------------------

    QRS_Estimate = BaseQRS + HFAdjustment + OSAAdjustment

    Exit Function

ErrorHandler:

    'Return Excel #N/A if an unexpected error occurs.
    QRS_Estimate = CVErr(xlErrNA)

End Function


