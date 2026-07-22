Attribute VB_Name = "BP_Update"

Option Explicit

'========================================================================================
' Module name:
'   BP_Update
'
' Purpose:
'   This module updates systolic blood pressure (SBP) cycle-by-cycle using the new
'   SBP equation that depends on age, sex, and BMI.
'
' Why this module exists:
'   The model is an individual patient simulation. Patients already have their own
'   baseline SBP. Therefore, we should NOT overwrite patient.SBP directly with the
'   predicted SBP from the equation.
'
'   Instead, we estimate the expected SBP before and after the current cycle's
'   age/BMI update, calculate the difference, and apply only this delta to the
'   patient's actual SBP.
'
'   New SBP = Old actual SBP + (Expected new SBP - Expected old SBP)
'
' Sources / documentation:
'
'   Source 1:
'   Boledovicova 2013
'   URL: https://medscimonit.com/abstract/full/idArt/883984
'   Role in this module:
'       Source listed in the uploaded workbook for the SBP equation using age,
'       body-fat percentage, sex, and age subgroup.
'
'   Source 2:
'   Gomez-Ambrosi et al. 2012
'   Diabetes Care. 2012;35(2):383-388.
'   DOI: https://doi.org/10.2337/dc11-1334
'   Role in this module:
'       Source listed in the uploaded workbook for estimating body-fat percentage
'       from age, sex, and BMI.
'
'
'========================================================================================


Public Function BP_Estimate_SBP_From_Age_BMI(ByVal Age As Double, _
                                             ByVal Female As Boolean, _
                                             ByVal BMI As Double) As Double

    '====================================================================================
    ' Function:
    '   BP_Estimate_SBP_From_Age_BMI
    '
    ' Purpose:
    '   Estimates expected SBP using the simplified final SBP equation from the uploaded
    '   workbook.
    '
    ' Inputs:
    '   Age    = patient age in years
    '   Female = True if patient is female, False if patient is male
    '   BMI    = body mass index in kg/m^2
    '
    ' Output:
    '   Expected SBP in mmHg
    '
    ' Equation selection logic:
    '   Female and Age <= 40  -> women <= 40 years equation
    '   Male and Age <= 40    -> men <= 40 years equation
    '   Female and Age > 40   -> women > 40 years equation
    '   Male and Age > 40     -> men > 40 years equation
    '
    '====================================================================================

    'Basic input checks.
    'If inputs are invalid, return 0. The calling subroutine will decide whether to update.
    If Age <= 0 Then
        BP_Estimate_SBP_From_Age_BMI = 0
        Exit Function
    End If

    If BMI <= 0 Then
        BP_Estimate_SBP_From_Age_BMI = 0
        Exit Function
    End If


    If Age <= 40 Then

        If Female = True Then

            'Women <= 40 years
            'Source: uploaded workbook "SBP equation.xlsx", final simplified equation.
            BP_Estimate_SBP_From_Age_BMI = _
                85.247317 _
                + (0.561751 * Age) _
                + (1.398201 * BMI) _
                - (0.012927 * BMI ^ 2) _
                - (0.00834 * BMI * Age) _
                + (0.00008757 * BMI ^ 2 * Age)

        Else

            'Men <= 40 years
            'Source: uploaded workbook "SBP equation.xlsx", final simplified equation.
            BP_Estimate_SBP_From_Age_BMI = _
                59.001928 _
                + (0.794982 * Age) _
                + (3.152968 * BMI) _
                - (0.025844 * BMI ^ 2) _
                - (0.01988 * BMI * Age) _
                + (0.00020874 * BMI ^ 2 * Age)

        End If

    Else

        If Female = True Then

            'Women > 40 years
            'Source: uploaded workbook "SBP equation.xlsx", final simplified equation.
            BP_Estimate_SBP_From_Age_BMI = _
                78.498433 _
                + (0.716499 * Age) _
                + (1.116549 * BMI) _
                - (0.010323 * BMI ^ 2) _
                - (0.00666 * BMI * Age) _
                + (0.00006993 * BMI ^ 2 * Age)

        Else

            'Men > 40 years
            'Source: uploaded workbook "SBP equation.xlsx", final simplified equation.
            BP_Estimate_SBP_From_Age_BMI = _
                58.411424 _
                + (0.626856 * Age) _
                + (3.019744 * BMI) _
                - (0.024752 * BMI ^ 2) _
                - (0.01904 * BMI * Age) _
                + (0.00019992 * BMI ^ 2 * Age)

        End If

    End If

End Function


Public Sub BP_Update_SBP_By_Delta(ByRef Patient As Patient, _
                                  ByVal OldAge As Double, _
                                  ByVal OldBMI As Double, _
                                  ByVal OldSBP As Double)

    '====================================================================================
    ' Subroutine:
    '   BP_Update_SBP_By_Delta
    '
    ' Purpose:
    '   Updates patient.SBP using the change predicted by the SBP equation between
    '   the start and end of the current model cycle.
    '
    ' Why we use delta instead of direct replacement:
    '
    '   Direct replacement would be:
    '       Patient.SBP = predicted SBP
    '
    '   That is not recommended in resistant hypertension because it would erase the
    '   patient's actual baseline SBP and force them toward the average SBP predicted
    '   by age/sex/BMI only.
    '
    '   Instead we use:
    '       DeltaSBP = ExpectedNewSBP - ExpectedOldSBP
    '       Patient.SBP = OldSBP + DeltaSBP
    '
    '   This preserves:
    '       - baseline resistant hypertension burden
    '       - between-patient heterogeneity
    '       - prior treatment/intervention effects
    '       - cycle-by-cycle continuity
    '
    ' Inputs:
    '   Patient = patient object currently being simulated
    '   OldAge  = age at the start of the cycle, before age progression
    '   OldBMI  = BMI at the start of the cycle, before BMI progression
    '   OldSBP  = actual SBP at the start of the cycle
    '
    ' Output:
    '   Updates Patient.SBP directly.
    '
    '====================================================================================

    Dim ExpectedOldSBP As Double
    Dim ExpectedNewSBP As Double
    Dim DeltaSBP As Double

    With Patient

        'Safety checks.
        'If any required value is invalid, do not update SBP.
        If OldAge <= 0 Then Exit Sub
        If OldBMI <= 0 Then Exit Sub
        If OldSBP <= 0 Then Exit Sub

        If .Age <= 0 Then Exit Sub
        If .BMI <= 0 Then Exit Sub

        'Calculate expected SBP at the start of the cycle,
        'using age and BMI before progression.
        ExpectedOldSBP = BP_Estimate_SBP_From_Age_BMI(OldAge, .Female, OldBMI)

        'Calculate expected SBP at the end of the cycle,
        'using updated age and updated BMI.
        ExpectedNewSBP = BP_Estimate_SBP_From_Age_BMI(.Age, .Female, .BMI)

        'If either expected SBP could not be calculated, do not update.
        If ExpectedOldSBP <= 0 Then Exit Sub
        If ExpectedNewSBP <= 0 Then Exit Sub

        'Calculate the expected SBP change caused by age/BMI progression.
        DeltaSBP = ExpectedNewSBP - ExpectedOldSBP

        'Apply only the expected change to the patient's actual SBP.
        .SBP = OldSBP + DeltaSBP

        'Clinical plausibility bounds.
        'These are not from the source equations; they are model safety checks.
        If .SBP < 80 Then .SBP = 80
        If .SBP > 260 Then .SBP = 260

        'Logical consistency check:
        'SBP should usually remain above DBP.
        If .DBP > 0 Then
            If .SBP <= .DBP Then .SBP = .DBP + 10
        End If

    End With

End Sub

