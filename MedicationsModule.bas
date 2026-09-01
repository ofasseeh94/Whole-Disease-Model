Attribute VB_Name = "MedicationsModule"
Option Explicit

Type Diabetes_Medication

    'Identification and Classification
    ID As Byte                          'Unique identifier for the medication
    INN As String                       'International Nonproprietary Name
    Insulin As Boolean                  'Is the medication insulin?
    WeightLoss_Drug As Boolean          'Is it a GLP-1/SGLT2 or similar with weight-loss effects?

    'HbA1C Effect (Monotherapy)
    HbA1C_Reduction_Mean As Single      'Mean HbA1C reduction as first therapy
    HbA1C_Reduction_Mean_Adj As Single  'Mean HbA1C reduction as first therapy adjusted to adherence
    HbA1C_Reduction_SE As Single        'Standard error of HbA1C reduction

    'HbA1C Effect (Add-on)
    HbA1C_Reduction_Addon_Mean As Single        'Mean HbA1C reduction as add-on
    HbA1C_Reduction_Addon_Mean_Adj As Single    'Mean HbA1C reduction as add-on adjusted to adherence
    HbA1C_Reduction_Addon_SE As Single          'SE of HbA1C reduction as add-on

    'Other Clinical Effects
    Weight_Impact As Single             'Weight impact (kg or %)
    Minor_Hypoglycemia As Single        'Probability of minor hypoglycemia
    Major_Hypoglycemia As Single        'Probability of major hypoglycemia
    Discontinuation_Rate As Single      'Probability of discontinuation (e.g. due to AE)

    'Economic Inputs
    Cost_Daily As Currency              'Daily cost of the drug
    source As String                    'Source of evidence or study

    'Use Characteristics
    Durability As Single                'Duration of effect (months)
    Compliance As Single                'Compliance rate (0 to 1)
    Adherence As Single                 'Adherence rate (0 to 1)
    Success_Rate As Single              'Proportion achieving clinical success
    Num_Medications As Byte             'Number of active ingredients or combo meds
    GLP As Boolean                      'if the treatment mix contains a GLP or not

End Type

Sub Load_Diabetes_Meds()

    Dim nMeds As Long
    Dim Col_Medications() As Diabetes_Medication
    Dim i As Long

    'Load medication table from the workbook
      Diabetes_Medications_Matrix = Range("DM_Medications_Table")

    'Determine number of medications (exclude header row)
    nMeds = UBound(Diabetes_Medications_Matrix)
    ReDim Col_Medications(1 To nMeds)

    'Populate user defined type array from the matrix
    For i = 1 To nMeds
        With Col_Medications(i)
            .ID = CByte(Diabetes_Medications_Matrix(i, 1))
            .INN = CStr(Diabetes_Medications_Matrix(i, 2))
            .Insulin = CBool(Diabetes_Medications_Matrix(i, 3))
            .WeightLoss_Drug = CBool(Diabetes_Medications_Matrix(i, 1))
            .HbA1C_Reduction_Mean = CSng(Diabetes_Medications_Matrix(i, 4))
            .HbA1C_Reduction_SE = CSng(Diabetes_Medications_Matrix(i, 5))
            .HbA1C_Reduction_Addon_Mean = CSng(Diabetes_Medications_Matrix(i, 6))
            .HbA1C_Reduction_Addon_SE = CSng(Diabetes_Medications_Matrix(i, 7))
            .Weight_Impact = CSng(Diabetes_Medications_Matrix(i, 8))
            .Minor_Hypoglycemia = CSng(Diabetes_Medications_Matrix(i, 9))
            .Major_Hypoglycemia = CSng(Diabetes_Medications_Matrix(i, 10))
            .Cost_Daily = CCur(Diabetes_Medications_Matrix(i, 11))
            .source = CStr(Diabetes_Medications_Matrix(i, 12))
            .Durability = CSng(Diabetes_Medications_Matrix(i, 13))
            .Compliance = CSng(Diabetes_Medications_Matrix(i, 14))
            .Adherence = CSng(Diabetes_Medications_Matrix(i, 15))
            .Success_Rate = CSng(Diabetes_Medications_Matrix(i, 16))
            .Num_Medications = CByte(Diabetes_Medications_Matrix(i, 17))
            .GLP = CBool(Diabetes_Medications_Matrix(i, 18))
            'currently discontinuation rate is not in the model
            '.Discontinuation_Rate = CSng(Diabetes_Medications_Matrix(i, 1))
            
            'adjust the mean HbA1c reduction using the adherence
            .HbA1C_Reduction_Mean_Adj = AdjustHbA1cForAdherence(Col_Medications(i), False)
            .HbA1C_Reduction_Addon_Mean_Adj = AdjustHbA1cForAdherence(Col_Medications(i), True)
            
        End With
    Next i

    Diabetes_Medications = Col_Medications

End Sub

Private Function AdjustHbA1cForAdherence(ByRef DM_Med As Diabetes_Medication, Addon As Boolean) As Single

    ' Purpose:
    ' Adjust the expected HbA1c reduction of a diabetes medication or
    ' medication combination according to the patient's expected adherence.
    '
    ' Assumptions:
    ' 1. DM_Med.HbA1C_Reduction_Mean represents the HbA1c effect at 100% adherence.
    ' 2. DM_Med.Adherence is stored as a proportion from 0 to 1.
    '    Example: 80% adherence = 0.8.
    ' 3. HbA1c reductions are stored as negative values.
    '
    ' Evidence used:
    ' - Non-insulin treatments:
    '   Gordon et al. 2018. Separate regression equations were reported for
    '   monotherapy, dual therapy, and triple therapy, with MPR as the actual
    '   adherence level.
    '
    ' - Insulin-containing treatments:
    '   Nguyen et al. 2024 systematic review and meta-analysis.
    '   The MPR-specific subgroup found approximately 0.03 percentage-point
    '   greater HbA1c reduction for every 1 percentage-point increase in adherence.
    '
    ' For non-insulin treatments, the Gordon regression is first used to predict
    ' the HbA1c effect at the patient's actual adherence. This is divided by the
    ' predicted Gordon effect at 100% adherence to calculate a relative efficacy
    ' multiplier. That multiplier is then applied to the drug-specific
    ' HbA1C_Reduction_Mean.
    '
    ' For insulin-containing regimens, the absolute HbA1c loss associated with
    ' reduced adherence is added back to the negative full-adherence HbA1c effect.

    Dim Multiplier As Single

    ' At or above 100% adherence, retain the full medication effect.
    If DM_Med.Adherence >= 1 Then
            If Addon = True Then
            
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Addon_Mean
            
            Else
                  
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Mean
            
            End If
            
        Exit Function
    End If

    ' With no adherence, assume no HbA1c-lowering treatment effect.
    If DM_Med.Adherence <= 0 Then
        AdjustHbA1cForAdherence = 0
        Exit Function
    End If


    If DM_Med.Insulin Then

        ' INSULIN-CONTAINING REGIMENS
        '
        ' Nguyen et al. 2024 MPR subgroup:
        ' Each 1 percentage-point increase in adherence was associated with
        ' approximately 0.03 percentage-point greater HbA1c reduction.
        '
        ' Since adherence is stored from 0 to 1:
        '   (1 - Adherence) * 100
        ' converts the adherence deficit from a proportion into percentage points.
        '
        ' Example:
        ' Adherence = 0.80
        ' Adherence deficit = 20 percentage points
        ' HbA1c effect lost = 20 * 0.03 = 0.60 percentage points.
        '
        ' Because HbA1C_Reduction_Mean is negative, the lost treatment effect
        ' is added back, making the resulting reduction less negative.
        
            If Addon = True Then
            
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Addon_Mean + ((1 - DM_Med.Adherence) * 100 * 0.03)
            
            Else
                  
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Mean + ((1 - DM_Med.Adherence) * 100 * 0.03)
            
            End If

        ' Nonadherence should be allowed to reduce the treatment effect to zero,
        ' but should not convert the pharmacological HbA1c-lowering effect into
        ' a positive HbA1c increase.
        If AdjustHbA1cForAdherence > 0 Then _
            AdjustHbA1cForAdherence = 0

        Exit Function


    Else

        ' NON-INSULIN REGIMENS
        '
        ' Gordon et al. 2018 reported separate regression equations predicting
        ' 1-year HbA1c change as a function of actual medication possession ratio
        ' for monotherapy, dual therapy, and triple therapy.
        '
        ' The Gordon-predicted HbA1c effect at the patient's adherence is divided
        ' by the Gordon-predicted effect at 100% adherence.
        '
        ' This produces a relative adherence multiplier:
        '
        '   Multiplier =
        '       Predicted HbA1c effect at actual adherence
        '       ------------------------------------------------
        '       Predicted HbA1c effect at 100% adherence
        '
        ' The relative multiplier is then applied to the specific medication's
        ' HbA1C_Reduction_Mean. This preserves differences in intrinsic efficacy
        ' between individual diabetes drugs and combinations.

        Select Case DM_Med.Num_Medications

            Case 1

                ' MONOTHERAPY
                '
                ' Gordon regression:
                '   HbA1c change = -0.002 - (0.892 * MPR)
                '
                ' At 100% adherence:
                '   -0.002 - (0.892 * 1) = -0.894
                '
                ' Therefore:
                '   Multiplier =
                '   (-0.002 - 0.892 * actual adherence) / -0.894

                Multiplier = (-0.002 - 0.892 * DM_Med.Adherence) / -0.894


            Case 2

                ' DUAL THERAPY
                '
                ' Gordon regression:
                '   HbA1c change = -0.133 - (0.859 * MPR)
                '
                ' At 100% adherence:
                '   -0.133 - (0.859 * 1) = -0.992
                '
                ' Therefore:
                '   Multiplier =
                '   (-0.133 - 0.859 * actual adherence) / -0.992

                Multiplier = (-0.133 - 0.859 * DM_Med.Adherence) / -0.992


            Case Else

                ' TRIPLE OR MORE THERAPY
                '
                ' Gordon reported the triple-therapy regression:
                '   HbA1c change = 0.840 - (2.058 * MPR)
                '
                ' At 100% adherence:
                '   0.840 - (2.058 * 1) = -1.218
                '
                ' The same equation is used for regimens with 3 or more
                ' medications because Gordon did not provide a separate
                ' regression for regimens containing more than 3 medications.
                '
                ' Therefore:
                '   Multiplier =
                '   (0.840 - 2.058 * actual adherence) / -1.218

                Multiplier = (0.84 - 2.058 * DM_Med.Adherence) / -1.218

        End Select


        ' Restrict the relative efficacy multiplier to the logical range 0 to 1.
        '
        ' Multiplier < 0 would reverse the treatment's HbA1c-lowering effect,
        ' which is not intended in this function.
        '
        ' Multiplier > 1 would imply an effect greater than the assumed
        ' 100%-adherence efficacy.
        If Multiplier < 0 Then Multiplier = 0
        If Multiplier > 1 Then Multiplier = 1


        ' Apply the adherence-derived relative efficacy multiplier to the
        ' medication-specific full-adherence HbA1c reduction.
            If Addon = True Then
            
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Addon_Mean * Multiplier
            
            Else
                  
                  AdjustHbA1cForAdherence = DM_Med.HbA1C_Reduction_Mean * Multiplier
            
            End If

    End If

End Function


