Attribute VB_Name = "Load_Inputs_Module"
Option Explicit
'Model Time Horizon
Public Timehorizon As Double

'Random Numbers array
Public RandArray() As Single

'General multipurpose matrix to work as a temporary transition probability matrix
Public TPMatrix As Variant

'General Array Variable to be used temporary in the model
Public Temp_Array As Variant

'Discount rate
Public Disc_Costs As Double
Public Disc_QALYs As Double

'Set a global variable as an array which collects probabilities of death from all sub models
Public Mortality_Arr() As Double
Public General_Mortality_Matrix() As Variant
'__________________________________________________________________________________________________________________________
'Interventions
Public Interventions() As intervention
'Total number of Interventions(both included and excluded from the model)
Public NInterventions As Byte
'number of active (selected) interventions in the model
Public Active_Interventions As Byte
'Current Intervention
Public ActiveIntervention As intervention
'Collection of interventions included in the model
Public Col_Interventions() As intervention
'__________________________________________________________________________________________________________________________
'Complications
'Total number of complications (both included and excluded from the model)
Public NComplications As Byte
'number of active (selected) complications in the model
Public Active_Complications As Byte
'Collection of complications included in the model
'Public Col_Complications() As Complication
'__________________________________________________________________________________________________________________________
'Patients
Public Patients() As Patient
'Set data source
Public Patient_Cohort_Matrix As Variant
'Number of patients
Public NPatients As Long
'Collection of patients
Public Col_Patients() As Patient
'Mean Age
Public Avg_Age As Double
'Utilities
Public Disutility_Arr() As Double
Public Utility_Instant As Double
Public Disutility_Method As String
Public Utility_Temp As Double
'Array in which all patient data will be pasted to
Public PastePatientArray() As Variant

'Physiological parameters baseline
Public Baseline_HDL
Public Baseline_TG
Public Baseline_TC
Public Baseline_LDL
'__________________________________________________________________________________________________________________________

'Load publicly reference arrays for CKD risk scoring
Public CKD_Age_Score_Table As Variant
Public CKD_Sex_Score_Table As Variant
Public CKD_BMI_Score_Table As Variant
Public CKD_Diabetes_Score_Table As Variant
Public CKD_SBP_Score_Table As Variant
Public CKD_RiskScore_Table As Variant

Public Sub Load_CKD_Risk_Score_Tables()

    '   Load the simplified CKD risk-score tables once
    '
    ' Source:
    '   Saranburut K, Vathesatogkit P, Thongmung N, et al.
    '   Risk scores to predict decreased glomerular filtration rate at 10 years in an Asian general population BMC Nephrology. 2017;18:240
    '   Table 3, Model 1 (BMI)
    '
    ' Why Model 1 BMI:
    '   This version does not require baseline eGFR
    '   It uses variables already available in the patient object:
    '       Age, Female, BMI, DM, SBP
    '
    ' Table structure:
    '   Each row is:
    '       lower boundary, upper boundary, score or probability
    '
    ' Notes:
    '   The final CKD_RiskScore_Table stores the published 10-year risk as decimals
    '   Example: 0.18 = 18% 10-year risk
    '====================================================================================

    CKD_Age_Score_Table = Array( _
        Array(0, 44.999, 0), _
        Array(45, 54.999, 2), _
        Array(55, 200, 4))

    CKD_Sex_Score_Table = Array( _
        Array(0, 0, 0), _
        Array(1, 1, 2))

    CKD_BMI_Score_Table = Array( _
        Array(0, 24.999, 0), _
        Array(25, 200, 1))

    CKD_Diabetes_Score_Table = Array( _
        Array(0, 0, 0), _
        Array(1, 1, 2))

    CKD_SBP_Score_Table = Array( _
        Array(0, 119.999, -2), _
        Array(120, 129.999, 0), _
        Array(130, 139.999, 1), _
        Array(140, 149.999, 2), _
        Array(150, 159.999, 2), _
        Array(160, 300, 3))

    CKD_RiskScore_Table = Array( _
        Array(-999, -2, 0.01), _
        Array(-1, -1, 0.02), _
        Array(0, 0, 0.03), _
        Array(1, 1, 0.04), _
        Array(2, 2, 0.04), _
        Array(3, 3, 0.06), _
        Array(4, 4, 0.07), _
        Array(5, 5, 0.09), _
        Array(6, 6, 0.11), _
        Array(7, 7, 0.14), _
        Array(8, 8, 0.18), _
        Array(9, 9, 0.23), _
        Array(10, 10, 0.3), _
        Array(11, 11, 0.34), _
        Array(12, 999, 0.5))

End Sub
'_________________________________________________________________________________________________________________________

'Diabetetstreatment algorithm
Public BaseLine_HbA1C As Single
Public treatment_effect_HbA1C As Single
Public Baseline_Before_Medication_HbA1C As Single
Public Time_Start_DM_Medication As Single
Public Diabetes_Treatment_Algorithm_Matrix As Variant
Public Diabetes_Medications_Matrix As Variant
Public Diabetes_Medications() As Diabetes_Medication

'Model Cycle length
'set in years
Public Cycle_Length As Double
Public Current_Year As Integer

Public Sub LoadInputs()

'set cycle length in years
'the value is provided in month in the model settings
Cycle_Length = Range("Cycle_Length").Value / 12

'Set the average age of patients
Avg_Age = Range("Mean_Age").Value

Call Load_General_Mortality
Call Load_Diabetes_Treatment_Algorithm
Call Load_Diabetes_Medications
Call Load_Diabetes_Meds

Current_Year = Year(Now)
Call Load_Discount_Rates
Timehorizon = Range("TimeHorizon")
Disutility_Method = Range("Disutility_Method")

'Load patient cohort
PreparePatientCohort
Call Load_Patient_Characteristics

'Load CKD risk score tables
Call Load_CKD_Risk_Score_Tables

'load random numbers
RandArray = GenerateRandomArray(NPatients, Timehorizon / Cycle_Length, 52, 42)

'Preparation: load interventions info
Interventions = Load_Interventions

'Load data required for submodels
Call Load_Foot_Ulcer
Call Load_NASH
Call Load_Nephro
Call Load_CHD
Call Load_CKD
Call Load_DKA
Call Load_DLP
Call Load_DM
Call Load_HTN
Call Load_MA
Call Load_MI
Call Load_Neuro
Call Load_OA
Call Load_OSA
Call Load_PVD
Call Load_Retino
Call Load_Stroke
Call Load_HypoGly

End Sub

Sub Load_General_Mortality()

'Get range where general mortality is located
General_Mortality_Matrix = Range("General_Mortality")

End Sub

Sub Load_Discount_Rates()

Disc_Costs = Range("Disc_Costs").Value
Disc_QALYs = Range("Disc_QALYs").Value

End Sub

Sub Load_Diabetes_Treatment_Algorithm()

'Set data source
'load diabetes treatment algorithm as a matrix
Diabetes_Treatment_Algorithm_Matrix = Range("DM_Treatment_Algorithm_Table")

End Sub

Sub Load_Diabetes_Medications()

'Set data source
'load diabetes treatment algorithm as a matrix
Diabetes_Medications_Matrix = Range("DM_Medications_Table")

End Sub

Sub Load_Foot_Ulcer()
'Load Costs and Disutilities
Foot_Ulcer_DisU_Costs = Range("Foot_Ulcer_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
Foot_Ulcer_Cycle = Range("Foot_Ulcer_Cycle_Length_Range").Value
Foot_Ulcer_Repeat = Round(Cycle_Length / Foot_Ulcer_Cycle, 0)

'Load TPM
TPMatrix = Range("Foot_Ulcer_TPM_Table")
'load death data from the TPM separetly
Foot_Ulcer_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
Foot_Ulcer_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub

Sub Load_Nephro()
'Load Costs and Disutilities
Nephro_DisU_Costs = Range("Nephro_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
Nephro_Cycle = Range("Nephro_Cycle_Length_Range").Value
Nephro_Repeat = Round(Cycle_Length / Nephro_Cycle, 0)

'Load TPM
TPMatrix = Range("Nephro_TPM_Table")
'load death data from the TPM separetly
Nephro_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
Nephro_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub

Sub Load_NASH()
'Load Costs and Disutilities
NASH_DisU_Costs = Range("NASH_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
NASH_Cycle = Range("NASH_Cycle_Length_Range").Value
NASH_Repeat = Round(Cycle_Length / NASH_Cycle, 0)

'Load TPM
TPMatrix = Range("NASH_TPM_Table")
'load death data from the TPM separetly
NASH_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
NASH_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_Stroke()
'Load Costs and Disutilities
Stroke_DisU_Costs = Range("Stroke_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
Stroke_Cycle = Range("Stroke_Cycle_Length_Range").Value
Stroke_Repeat = Round(Cycle_Length / Stroke_Cycle, 0)

'Load TPM
TPMatrix = Range("Stroke_TPM_Table")
'load death data from the TPM separetly
Stroke_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
Stroke_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_MI()
'Load Costs and Disutilities
MI_DisU_Costs = Range("MI_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
MI_Cycle = Range("MI_Cycle_Length_Range").Value
MI_Repeat = Round(Cycle_Length / MI_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("MI_TPM_Table")
'load death data from the TPM separetly
MI_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
MI_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_HTN()
'Load Costs and Disutilities
HTN_DisU_Costs = Range("HTN_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
HTN_Cycle = Range("HTN_Cycle_Length_Range").Value
HTN_Repeat = Round(Cycle_Length / HTN_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("HTN_TPM_Table")
'load death data from the TPM separetly
HTN_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
HTN_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_CHD()
'Load Costs and Disutilities
CHD_DisU_Costs = Range("CHD_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
CHD_Cycle = Range("CHD_Cycle_Length_Range").Value
CHD_Repeat = Round(Cycle_Length / CHD_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("CHD_TPM_Table")
'load death data from the TPM separetly
CHD_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
CHD_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_CKD()
'Load Costs and Disutilities
CKD_DisU_Costs = Range("CKD_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
CKD_Cycle = Range("CKD_Cycle_Length_Range").Value
CKD_Repeat = Round(Cycle_Length / CKD_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("CKD_TPM_Table")
'load death data from the TPM separetly
CKD_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
CKD_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_DM()
'Load Costs and Disutilities
DM_DisU_Costs = Range("DM_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
DM_Cycle = Range("DM_Cycle_Length_Range").Value
DM_Repeat = Round(Cycle_Length / DM_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("DM_TPM_Table")
'load death data from the TPM separetly
DM_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
DM_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_DLP()
'Load Costs and Disutilities
DLP_DisU_Costs = Range("DLP_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
DLP_Cycle = Range("DLP_Cycle_Length_Range").Value
DLP_Repeat = Round(Cycle_Length / DLP_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("DLP_TPM_Table")
'load death data from the TPM separetly
DLP_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
DLP_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_HypoGly()
'Load Costs and Disutilities
HypoGly_DisU_Costs = Range("HypoGly_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
HypoGly_Cycle = Range("HypoGly_Cycle_Length_Range").Value
HypoGly_Repeat = Round(Cycle_Length / HypoGly_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("HypoGly_TPM_Table")
'load death data from the TPM separetly
HypoGly_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
HypoGly_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_DKA()
'Load Costs and Disutilities
DKA_DisU_Costs = Range("DKA_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
DKA_Cycle = Range("DKA_Cycle_Length_Range").Value
DKA_Repeat = Round(Cycle_Length / DKA_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("DKA_TPM_Table")
'load death data from the TPM separetly
DKA_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
DKA_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_MA()
'Load Costs and Disutilities
MA_DisU_Costs = Range("MA_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
MA_Cycle = Range("MA_Cycle_Length_Range").Value
MA_Repeat = Round(Cycle_Length / MA_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("MA_TPM_Table")
'load death data from the TPM separetly
MA_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
MA_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_Neuro()
'Load Costs and Disutilities
Neuro_DisU_Costs = Range("Neuro_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
Neuro_Cycle = Range("Neuro_Cycle_Length_Range").Value
Neuro_Repeat = Round(Cycle_Length / Neuro_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("Neuro_TPM_Table")
'load death data from the TPM separetly
Neuro_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
Neuro_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_OA()
'Load Costs and Disutilities
OA_DisU_Costs = Range("OA_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
OA_Cycle = Range("OA_Cycle_Length_Range").Value
OA_Repeat = Round(Cycle_Length / OA_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("OA_TPM_Table")
'load death data from the TPM separetly
OA_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
OA_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_OSA()
'Load Costs and Disutilities
OSA_DisU_Costs = Range("OSA_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
OSA_Cycle = Range("OSA_Cycle_Length_Range").Value
OSA_Repeat = Round(Cycle_Length / OSA_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("OSA_TPM_Table")
'load death data from the TPM separetly
OSA_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
OSA_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_PVD()
'Load Costs and Disutilities
PVD_DisU_Costs = Range("PVD_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
PVD_Cycle = Range("PVD_Cycle_Length_Range").Value
PVD_Repeat = Round(Cycle_Length / PVD_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("PVD_TPM_Table")
'load death data from the TPM separetly
PVD_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
PVD_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Sub Load_Retino()
'Load Costs and Disutilities
Retino_DisU_Costs = Range("Retino_DisU_Costs_Table")

'How many times the model should repeate within the main model cycle length
Retino_Cycle = Range("Retino_Cycle_Length_Range").Value
Retino_Repeat = Round(Cycle_Length / Retino_Cycle, 0)

'Load TPM
'The user must provide the transition probability matrix full with headers for columns and rows in text format. the names of the health states must be the same as for the rest of the model
TPMatrix = Range("Retino_TPM_Table")
'load death data from the TPM separetly
Retino_Death_TPM = DeathTPM(TPMatrix)
'Load and adjust the TPM excluding death. this also aggregates the Transition probabilities
Retino_TPM = Adj_TPM(NoDeath_TPM(TPMatrix))

End Sub
Function Adj_TPM(TPMatrixTemp As Variant)
'Adjust Transition Probability Matrix

'define an index variable
Dim c As Integer
Dim r As Integer
Dim i As Integer
Dim TempVal As Double
Dim DESoptions As Variant

'if input is provided as a range rather than an array convert range into a proper array
If TypeName(TPMatrixTemp) = "Range" Then TPMatrixTemp = TPMatrixTemp.Formula

'Get the number of health states in the ttransition probability matrix
Dim TPMatrixSize As Byte
TPMatrixSize = UBound(TPMatrixTemp) - 1
'ReDim DESoptions(1 To TPMatrixSize)

'convert all values within the matrix to numerical values
      
For r = 2 To TPMatrixSize + 1
      For c = TPMatrixSize + 1 To 2 Step -1
            TPMatrixTemp(r, c) = CDbl(TPMatrixTemp(r, c))
      Next c
Next r


'check the orientation of the transition probability matrix
Dim OrientationSum As Double
      For c = 2 To TPMatrixSize + 1
            OrientationSum = OrientationSum + TPMatrixTemp(2, c)
      Next c
'transpose transition probability matrix if not horizontal
'If OrientationSum <> 1 Then TPMatrixTemp = Application.Transpose(TPMatrixTemp)

'aggregate the transition probability matrix to be able to use it with individual simulation models

      For r = 2 To TPMatrixSize + 1
            For c = TPMatrixSize + 1 To 2 Step -1
                  If c = 2 Then
                        TPMatrixTemp(r, c) = 0
                  Else
                        For i = c - 1 To 2 Step -1
                              TempVal = TempVal + TPMatrixTemp(r, i)
                        Next i
                        
                        TPMatrixTemp(r, c) = TempVal
                        TempVal = 0
                        
                  End If
            Next c
      Next r

Adj_TPM = TPMatrixTemp

End Function



Public Function NoDeath_TPM(TPMatrixTemp As Variant) As Variant
    Dim numRows As Long, numCols As Long
    Dim i As Long, j As Long
    Dim rowSum As Double
    Dim processedMatrix As Variant

'if input is provided as a range rather than an array convert range into a proper array
If TypeName(TPMatrixTemp) = "Range" Then TPMatrixTemp = TPMatrixTemp.Formula


    numRows = UBound(TPMatrixTemp, 1)
    numCols = UBound(TPMatrixTemp, 2)

'convert all values within the matrix to numerical values
For i = 2 To numRows
    For j = 2 To numCols
        TPMatrixTemp(i, j) = CDbl(TPMatrixTemp(i, j))
    Next j
Next i



    ' Check if the last column in the first row contains "Death"
    If TPMatrixTemp(1, numCols) = "Death" Then
        ' Resize the matrix to remove the last column
        ReDim processedMatrix(1 To numRows - 1, 1 To numCols - 1)
        
        For i = 2 To numRows - 1
            rowSum = 0
            ' Copy values to the new matrix and calculate row sum
            For j = 2 To numCols - 1
                processedMatrix(i, j) = TPMatrixTemp(i, j)
                rowSum = rowSum + TPMatrixTemp(i, j)
            Next j
            
            ' Reweight each row so that the sum equals one
            If rowSum <> 0 Then
                For j = 2 To numCols - 1
                    processedMatrix(i, j) = processedMatrix(i, j) / rowSum
                Next j
            End If
        Next i
        
        'Fill headers
        For i = 1 To numRows - 1
            For j = 1 To numCols - 1
                  If i = 1 Or j = 1 Then processedMatrix(i, j) = TPMatrixTemp(i, j)
            Next j
        Next i
        NoDeath_TPM = processedMatrix
        
        Else
        
            NoDeath_TPM = TPMatrixTemp
            
        End If

End Function

Public Function DeathTPM(inputMatrix As Variant) As Variant
    Dim numRows As Long, numCols As Long
    Dim i As Long
    Dim outputMatrix As Variant

    ' If input is provided as a range rather than an array, convert range into an array
    If TypeName(inputMatrix) = "Range" Then
        
        inputMatrix = inputMatrix.Value
    
    End If

    numRows = UBound(inputMatrix, 1)
    numCols = UBound(inputMatrix, 2)

    ' Resize the output matrix to have two columns and the same number of rows as the input
    ReDim outputMatrix(1 To numRows, 1 To 2)

If inputMatrix(1, numCols) = "Death" Then
    
    ' Copy the first and last columns from the input matrix to the output matrix
    For i = 1 To numRows
        outputMatrix(i, 1) = inputMatrix(i, 1)       ' Copy the first column
        outputMatrix(i, 2) = inputMatrix(i, numCols) ' Copy the last column
    Next i
    
Else
    
    For i = 1 To numRows
        outputMatrix(i, 1) = inputMatrix(i, 1)       ' Copy the first column
        outputMatrix(i, 2) = 0 ' Copy the last column
    Next i
    
End If

    DeathTPM = outputMatrix
    
End Function
