Attribute VB_Name = "MedicationsModule"
Option Explicit

Type Diabetes_Medication

    'Identification and Classification
    ID As Byte                          'Unique identifier for the medication
    INN As String                       'International Nonproprietary Name
    Insulin As Boolean                  'Is the medication insulin?
    WeightLoss_Drug As Boolean          'Is it a GLP-1/SGLT2 or similar with weight-loss effects?

    'HbA1C Effect (Monotherapy)
    HbA1C_Reduction_Mean As Single      'Mean HbA1C reduction as monotherapy
    HbA1C_Reduction_SE As Single        'Standard error of HbA1C reduction

    'HbA1C Effect (Add-on)
    HbA1C_Reduction_Addon_Mean As Single 'Mean HbA1C reduction as add-on
    HbA1C_Reduction_Addon_SE As Single   'SE of HbA1C reduction as add-on

    'Other Clinical Effects
    Weight_Impact As Single             'Weight impact (kg or %)
    Minor_Hypoglycemia As Single        'Probability of minor hypoglycemia
    Major_Hypoglycemia As Single        'Probability of major hypoglycemia
    Discontinuation_Rate As Single      'Probability of discontinuation (e.g. due to AE)

    'Economic Inputs
    Cost_Daily As Currency              'Daily cost of the drug
    Source As String                    'Source of evidence or study

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
            .Source = CStr(Diabetes_Medications_Matrix(i, 12))
            .Durability = CSng(Diabetes_Medications_Matrix(i, 13))
            .Compliance = CSng(Diabetes_Medications_Matrix(i, 14))
            .Adherence = CSng(Diabetes_Medications_Matrix(i, 15))
            .Success_Rate = CSng(Diabetes_Medications_Matrix(i, 16))
            .Num_Medications = CByte(Diabetes_Medications_Matrix(i, 17))
            .GLP = CBool(Diabetes_Medications_Matrix(i, 18))
            'currently discontinuation rate is not in the model
            '.Discontinuation_Rate = CSng(Diabetes_Medications_Matrix(i, 1))
            
        End With
    Next i

    Diabetes_Medications = Col_Medications

End Sub


