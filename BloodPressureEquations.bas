Attribute VB_Name = "BloodPressureEquations"
Option Explicit

'====================================================================================
' BloodPressureEquations
'
' Purpose:
'   Provides a BMI/age/sex based systolic blood pressure (SBP) equation to replace
'   the previous SBP calculation based on pulse pressure.
'
' Sources used in the derivation provided by the user:
'   1) Boledovicova 2013 SBP equations by sex and age subgroup
'      https://medscimonit.com/abstract/full/idArt/883984
'   2) Gomez-Ambrosi 2012 body fat percentage estimation equation
'      https://doi.org/10.2337/dc11-1334
'
' Model convention:
'   Female = True for women, False for men.
'   Age is in years.
'   BMI is kg/m^2.
'====================================================================================

Public Function BodyFatPercentage_GomezAmbrosi(ByVal Age As Single, _
                                               ByVal Female As Boolean, _
                                               ByVal BMI As Single) As Single
    ' BF% = -44.988 + 0.503*Age + 10.689*Female + 3.172*BMI - 0.026*BMI^2
    '       + 0.181*BMI*Female - 0.020*BMI*Age - 0.005*BMI^2*Female
    '       + 0.00021*BMI^2*Age

    Dim Female01 As Single
    Female01 = IIf(Female, 1, 0)

    BodyFatPercentage_GomezAmbrosi = -44.988 _
        + (0.503 * Age) _
        + (10.689 * Female01) _
        + (3.172 * BMI) _
        - (0.026 * BMI ^ 2) _
        + (0.181 * BMI * Female01) _
        - (0.02 * BMI * Age) _
        - (0.005 * BMI ^ 2 * Female01) _
        + (0.00021 * BMI ^ 2 * Age)
End Function

Public Function Estimate_SBP_Boledovicova(ByVal Age As Single, _
                                           ByVal Female As Boolean, _
                                           ByVal BMI As Single) As Single
    ' Final simplified SBP equation after merging the SBP/body-fat equation with
    ' the body-fat estimation equation from the uploaded workbook.
    '
    ' Age subgroup rule:
    '   Age <= 40: younger subgroup equation
    '   Age > 40 : older subgroup equation

    If Age <= 0 Or BMI <= 0 Then
        Estimate_SBP_Boledovicova = 0
        Exit Function
    End If

    If Age > 40 Then
        If Female Then
            ' Women >40 years
            Estimate_SBP_Boledovicova = 78.498433 _
                + (0.716499 * Age) _
                + (1.116549 * BMI) _
                - (0.010323 * BMI ^ 2) _
                - (0.00666 * BMI * Age) _
                + (0.00006993 * BMI ^ 2 * Age)
        Else
            ' Men >40 years
            Estimate_SBP_Boledovicova = 58.411424 _
                + (0.626856 * Age) _
                + (3.019744 * BMI) _
                - (0.024752 * BMI ^ 2) _
                - (0.01904 * BMI * Age) _
                + (0.00019992 * BMI ^ 2 * Age)
        End If
    Else
        If Female Then
            ' Women <=40 years
            Estimate_SBP_Boledovicova = 85.247317 _
                + (0.561751 * Age) _
                + (1.398201 * BMI) _
                - (0.012927 * BMI ^ 2) _
                - (0.00834 * BMI * Age) _
                + (0.00008757 * BMI ^ 2 * Age)
        Else
            ' Men <=40 years
            Estimate_SBP_Boledovicova = 59.001928 _
                + (0.794982 * Age) _
                + (3.152968 * BMI) _
                - (0.025844 * BMI ^ 2) _
                - (0.01988 * BMI * Age) _
                + (0.00020874 * BMI ^ 2 * Age)
        End If
    End If
End Function

Public Sub Update_SBP_From_BMI_Age(ByRef Patient As patient)
    ' Updates current patient SBP using the new BMI/age/sex equation.
    ' Recommended placement: after Age and BMI are updated in Characteristics_Progression,
    ' replacing the old SBP update that was based on pulse pressure.

    With Patient
        If .Age <= 0 Or .BMI <= 0 Then Exit Sub
        .SBP = Estimate_SBP_Boledovicova(.Age, .Female, .BMI)
    End With
End Sub
