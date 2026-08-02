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


Public Function BP_SBP_Absolute_Change(ByRef Patient As Patient, _
                                       ByVal OldAge As Double, _
                                       ByVal OldBMI As Double) As Double

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

        'Calculate expected SBP at the start of the cycle,using age and BMI before progression.
        ExpectedOldSBP = BP_Estimate_SBP_From_Age_BMI(OldAge, .Female, OldBMI)

        'Calculate expected SBP at the end of the cycle,using updated age and updated BMI.
        ExpectedNewSBP = BP_Estimate_SBP_From_Age_BMI(.Age, .Female, .BMI)

        'Calculate the expected SBP change.
        DeltaSBP = ExpectedNewSBP - ExpectedOldSBP
        BP_SBP_Absolute_Change = DeltaSBP

    End With

End Function

Function DBP_Framingham(ByVal Age As Double, ByVal BaselineBMI As Double) As Double

    'Cheng et al. Framingham DBP age component.
    
    'Source:
    'Cheng S, Xanthakis V, Sullivan LM, Vasan RS.
    'Blood Pressure Tracking Over the Adult Life Course:
    'Patterns and Correlates in the Framingham Heart Study.
    'Hypertension. 2012/2013.
    
    'DBP coefficients:
    'Age per 10 years = 4.159
    'Age-squared = -0.838
    'Age x BMI = -0.422
    
    'Age is centered at 49 years and expressed per 10 years.
    'BMI is expressed per 5 kg/m2.

    Dim Age10 As Double
    Dim BMI5 As Double

    If Age <= 0 Then Exit Function
    If BaselineBMI <= 0 Then Exit Function

    Age10 = (Age - 49) / 10
    BMI5 = BaselineBMI / 5

    DBP_Framingham = _
          (4.159 * Age10) _
        - (0.838 * Age10 ^ 2) _
        - (0.422 * Age10 * BMI5)

End Function

Public Function DBP_Framingham_Absolute_Change(Patient As Patient) As Double

    With Patient

        DBP_Framingham_Absolute_Change = _
            DBP_Framingham(.Age, .BMI_Baseline) _
            - DBP_Framingham(OldAge, .BMI_Baseline)

    End With

End Function

