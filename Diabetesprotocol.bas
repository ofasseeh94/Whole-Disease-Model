Attribute VB_Name = "Diabetesprotocol"
Option Explicit
Sub Diabetes_Treatment_Sequence_Update(Patient As Patient)
'check if patient is treatment naive
      If IsEmpty(Patient.Diabetes_treatment_ID) Or Patient.Diabetes_treatment_ID = 0 Then
            
            'if patient is treatment naive go through the first line algorithm
            Call Diabetes_Treatment_FirstLine(Patient)
            Call Diabetes_Medication_Update(Patient)
            'on drug administration update the HbA1C
            Call Get_Medication_Row(Patient)
            Call HbA1C_Update(Patient, 0)
            
      Else
            
            'if patient is already on treatment check if the patient is not controlled or need to change medication for other reasons
            If Diabetes_Controlled(Patient) = False Then
                  
                  Call Diabetes_Treatment_Followup(Patient)
                  Call Diabetes_Medication_Update(Patient)
                  
            End If
            
      End If

End Sub
Sub HbA1C_Update(Patient As Patient, Treatment_Line As Byte)
Dim i As Integer

With Patient
                  
            If Treatment_Line = 0 Then

            'effect of first line treatment
                  Baseline_Before_Medication_HbA1C = .HbA1C
                  Time_Start_DM_Medication = .time_elapsed
                  treatment_effect_HbA1C = Application.WorksheetFunction.Min(0, Application.WorksheetFunction.Norm_Inv(RandArray(.ID, .time_elapsed * 2, 22), .Diabetes_Drug.HbA1C_Reduction_Mean, .Diabetes_Drug.HbA1C_Reduction_SE))

            Else

            'effect of adding a treatment
                  Baseline_Before_Medication_HbA1C = .HbA1C
                  Time_Start_DM_Medication = .time_elapsed
                  treatment_effect_HbA1C = Application.WorksheetFunction.Min(0, Application.WorksheetFunction.Norm_Inv(RandArray(.ID, .time_elapsed * 2, 23), .Diabetes_Drug.HbA1C_Reduction_Addon_Mean, .Diabetes_Drug.HbA1C_Reduction_Addon_SE))

            End If
            
            'if the patient is currently on treatment for obesity with a GLP don't double count the weight loss effect
            
            Call Update_BMI_DM_Medication(Patient)
            
'            If ActiveIntervention.GLP And .Diabetes_Drug.GLP And ActiveIntervention.Maint_Duration >= .time_elapsed Then
'                *****************Suspended until we add a dedicated column in diabetes medications that identify treatment**********
                  '*******************combinations with GLP as the last addition to the regimen*****************
'            Else
'
'                  Call Update_BMI_DM_Medication(patient)
'
'            End If

            If .Diabetes_Drug.Insulin = True Then .Insulin = True

End With


End Sub
Sub Update_BMI_DM_Medication(Patient As Patient)

With Patient

      Dim PatientMass As Single ' patient mass in kg
      Dim PatientHeight As Single 'patient height in m
      '1- Calculate mass based on BMI
      If .Female = True Then PatientHeight = 1.6088 Else PatientHeight = 1.745
      PatientMass = .BMI * PatientHeight ^ 2
      
      '2- Subtract mass reduction from current mass
      PatientMass = PatientMass + .Diabetes_Drug.Weight_Impact
      '3- Recalculate BMI based on new mass
      .BMI = PatientMass / PatientHeight ^ 2

End With

End Sub

Sub Diabetes_Medication_Update(Patient As Patient)
      
Restart:
      
'match TS code against Diabetes_Treatment_Algorithm_Matrix and update Diabetes_treatment_ID
      Dim i As Integer
      For i = 1 To UBound(Diabetes_Treatment_Algorithm_Matrix)
      
            'get the corresponding medication ID to the treatment sequence provided
            If Diabetes_Treatment_Algorithm_Matrix(i, 1) = Patient.Diabetes_Treatment_Sequence Then
            
                  'if the match results in a code is a jump to another treatment sequence and not a medication ID loop again and get the drug ID next to the TS code
                  If Diabetes_Treatment_Algorithm_Matrix(i, 2) <> onlyDigits(Diabetes_Treatment_Algorithm_Matrix(i, 2)) Then
                          
                          'Update the treatment squence to jump to the new treatment sequence
                          Patient.Diabetes_Treatment_Sequence = Diabetes_Treatment_Algorithm_Matrix(i, 2)
                          GoTo Restart
                          
                  End If
                
                  'in case the current drug is the last resort stay on the same drug.
                  If Diabetes_Treatment_Algorithm_Matrix(i, 2) = "" Or Diabetes_Treatment_Algorithm_Matrix(i, 2) = 0 Then
                  
                        'go back to the previous treatment sequence code
                        Patient.Diabetes_Treatment_Sequence = "TS" & onlyDigits(Patient.Diabetes_Treatment_Sequence) - 1
                        GoTo Restart
                  End If
                                 
                  'update diabetes treatment/medication ID
                  
                  'if the treament changed
                  If Patient.Diabetes_treatment_ID <> Diabetes_Treatment_Algorithm_Matrix(i, 2) Then
                        'update the treatment ID
                        Patient.Diabetes_treatment_ID = Diabetes_Treatment_Algorithm_Matrix(i, 2)
                        Call Get_Medication_Row(Patient)
                        
                        'on drug change update the HbA1C
                        Dim Treatment_Line As Byte
                        Treatment_Line = onlyDigits(Patient.Diabetes_Treatment_Sequence) Mod 10
                        Call HbA1C_Update(Patient, Treatment_Line)

                  End If
                  
                  GoTo JumpOut
                  
            End If
      Next i

JumpOut:

End Sub

Function Get_Diabetes_Medication(Diabetes_Treatment_Sequence As String) As Variant
Dim outcome(1 To 2) As Variant
Restart:
      
'match TS code against Diabetes_Treatment_Algorithm_Matrix and update Diabetes_treatment_ID
      Dim i As Integer
      For i = 1 To UBound(Diabetes_Treatment_Algorithm_Matrix)
      
            'get the corresponding medication ID to the treatment sequence provided
            If Diabetes_Treatment_Algorithm_Matrix(i, 1) = Diabetes_Treatment_Sequence Then
            
                  'if the match results in a code is a jump to another treatment sequence and not a medication ID loop again and get the drug ID next to the TS code
                  If Diabetes_Treatment_Algorithm_Matrix(i, 2) <> onlyDigits(Diabetes_Treatment_Algorithm_Matrix(i, 2)) Then
                          
                          'Update the treatment squence to jump to the new treatment sequence
                          Diabetes_Treatment_Sequence = Diabetes_Treatment_Algorithm_Matrix(i, 2)
                          GoTo Restart
                          
                  End If
                
                  'in case the current drug is the last resort stay on the same drug.
                  If Diabetes_Treatment_Algorithm_Matrix(i, 2) = "" Or Diabetes_Treatment_Algorithm_Matrix(i, 2) = 0 Then
                  
                        'go back to the previous treatment sequence code
                        Diabetes_Treatment_Sequence = "TS" & onlyDigits(Diabetes_Treatment_Sequence) - 1
                        GoTo Restart
                        
                  End If
                  
                  outcome(1) = Diabetes_Treatment_Algorithm_Matrix(i, 2)
                  outcome(2) = Diabetes_Treatment_Sequence
                  
                  Get_Diabetes_Medication = outcome
                  GoTo JumpOut
                  
            End If
      Next i

JumpOut:

End Function
Function SeverityCheck(Patient As Patient) As Byte
'This function classified the severity of diabetes in patients so they can be assigned to the proper treatment for their severity later
      
'The definition is Blood glucose levels > 300mg/dL or HbA1C > 10%  or unexpected weight loss or urine +++ for ketones based on American diabetes association

'check if patient has an extreemly high HbA1C and will NOT be controlled.

      If Patient.HbA1C >= 10 And Predict_Catabolic_Control(Patient) = False Then
            
            If Patient.FBS > 300 Or (Patient.BMI_Baseline - Patient.BMI) > 10 Then
                                  
                  SeverityCheck = 3
            
            Else
                  
                  SeverityCheck = 2
                  
            End If
            
      Else
      
            SeverityCheck = 1
            
      End If

End Function

Function Predict_Catabolic_Control(Patient As Patient) As Boolean
' this function is intended to evaluate if catabolic patient are potential to be controlled on intensive treatment or beyond control.

Dim Catabolic_ttt_Effect As Double
Dim Catabolic_ttt_effect_Mean As Double, Catabolic_ttt_effect_SE As Double

'The effect of the catabolic control is hardcoded here
Catabolic_ttt_effect_Mean = -2.88
Catabolic_ttt_effect_SE = Abs(Catabolic_ttt_effect_Mean * 0.1)

'Get the impact of intensive medication for controlling catabolic patients
'time elapsed is assumed to be 1 cycle which is half a year (assuming the phyisican will maximum wait for half a year for controlling the patient
Catabolic_ttt_Effect = Application.WorksheetFunction.Min(0, Application.WorksheetFunction.Norm_Inv(RandArray(Patient.ID, 0.5 * 2, 52), Catabolic_ttt_effect_Mean, Catabolic_ttt_effect_SE))

'Check if patient could be controlled after intensive treatment or not.
'if the criteria for controlled patient is changed in the Diabetes_Controlled function it has to be similar here
If Patient.HbA1C + Catabolic_ttt_Effect < 7 Then

      Predict_Catabolic_Control = True

Else

      Predict_Catabolic_Control = False

End If

'This section is only used for calibration and should only be used when a new population is entered in the model
'The target used is 85% roughly of the patients to be controlled

'With Sheets("Catabolic Control Test")
'
'.Cells(patient.ID + 1, 1) = patient.HbA1C
'.Cells(patient.ID + 1, 2) = Catabolic_ttt_Effect
'.Cells(patient.ID + 1, 3) = Predict_Catabolic_Control
'
'End With

End Function

Sub Diabetes_Treatment_FirstLine(Patient As Patient)

'*****************
'!!!!!!!if this submodule is updated with a new treatment algorithm the "Function Predict_Current_TTT"
'*****************

'also will need to be updated accordingly to match the same algorithm!!!!!!!

With Patient

      If .DM_type1 = True Then
      'Diabetes Type 1 patients
            'if metformin tolerant
            If .Metformin_intolerance = False Then
                              
                  .Diabetes_Treatment_Sequence = "TS100"
                  
            Else
            'if metformin intolerant
                  .Diabetes_Treatment_Sequence = "TS110"
            
            End If
      
      Else
      'Diabetes Type 2 patients
      
            Select Case SeverityCheck(Patient)
            Case 3
            'HBA1C >=## with (severe wt loss and FBS >300) Severe catabolic manifestations
                  
                  .Diabetes_Treatment_Sequence = "TS10"
        
            Case 2
            'IF ONLY >=##
                  .Diabetes_Treatment_Sequence = "TS10"
                  
            Case 1
            
                  If ASCVD_check(Patient) = True Or .HF Or .CKD Then
                  
                        If .Metformin_intolerance = False Then
                              
                              .Diabetes_Treatment_Sequence = "TS20"
                              
                        Else
                        
                              .Diabetes_Treatment_Sequence = "TS30"
                        
                        End If
                  
                  Else
                  
                        If .NASH = True Then
                        
                              If .Metformin_intolerance = False Then
                              
                                    .Diabetes_Treatment_Sequence = "TS80"
                              
                              Else
                        
                                    .Diabetes_Treatment_Sequence = "TS90"
                        
                              End If
                        
                        Else
                        
                              If .BMI < 30 Then
                                    
                                    If .Metformin_intolerance = False Then
                                    
                                          .Diabetes_Treatment_Sequence = "TS60"
                                    
                                    Else
                              
                                          .Diabetes_Treatment_Sequence = "TS70"
                              
                                    End If
                                    
                              Else
                              
                                    If .Metformin_intolerance = False Then
                                    
                                          .Diabetes_Treatment_Sequence = "TS40"
                                    
                                    Else
                              
                                          .Diabetes_Treatment_Sequence = "TS50"
                              
                                    End If
                              
                              End If
                              
                        End If
                        
                  End If
                  
            End Select
      
      End If

End With

End Sub

Function Diabetes_Controlled(Patient As Patient) As Boolean
'check if patient is controlled
'only factor used here is HbA1C
'lateron other factors can be added such as FBS or a combineation of factors etc.

With Patient

'      If .Age > 65 Or .LVH = True Or .Stroke = True Or .MI = True Or .CKD = True Then
'
'            If patient.HbA1C < 7.5 Then Diabetes_Controlled = True
'
'      Else
'
            If Patient.HbA1C < 7 Then Diabetes_Controlled = True
'
'      End If

End With

End Function

Sub Diabetes_Treatment_Followup(Patient As Patient)
            'increase the TS code by 1
            Patient.Diabetes_Treatment_Sequence = "TS" & CInt(onlyDigits(Patient.Diabetes_Treatment_Sequence)) + 1

End Sub

Sub Get_Medication_Row(Patient As Patient)
Dim i As Integer

With Patient

      For i = 1 To UBound(Diabetes_Medications)
      
            If Diabetes_Medications(i).ID = .Diabetes_treatment_ID Then
                  
                  .Diabetes_Drug = Diabetes_Medications(i)
                                                
                  If .Diabetes_Drug.Insulin = True Then .Insulin = True
            
            GoTo JumpOut
            End If
            
      Next i

End With

JumpOut:

End Sub

Function Get_Med_Row(MedicationID As Single) As Single
Dim i As Integer

      For i = 1 To UBound(Diabetes_Medications)
      
            If Diabetes_Medications(i).ID = MedicationID Then
                  
                  Get_Med_Row = i
            
            GoTo JumpOut
            End If
            
      Next i

JumpOut:

End Function

Function Predict_Current_TTT(Patient As Patient) As Variant

Dim CumulativeToT As Double
Dim outcome(1 To 4) As Variant

'a.    Develop algorithm to calculate the current drug and the starting time
'i.    Simple based on the existing algorithm for starting treatment sequence then add the durability durations until we reach the current time (duration of diabetes)
'ii.Input:
'1.    Same inputs required for the treatment TS algorithm (consider that calling the TS module in this phase doesn't trigger changes in the model)
'2.    Durability of each drug
'iii.Output:
'1.    Current Treatment sequence TS??
'2.    Duration until now in this TS


Dim Current_Treatment_Sequence As String
Dim DM_duration As Double
Dim CurrentMedication As Variant

With Patient

DM_duration = .Age - .DM_Diagnosis_Age

'check if patient is treatment naive
      If IsEmpty(Patient.Diabetes_Treatment_Sequence) Or onlyDigits(Patient.Diabetes_Treatment_Sequence) = 0 Then

            'if patient current treatment sequence is not elaborated we will predict it based on the duration of diabetes and durability

      If .DM_type1 = True Then
            If .Metformin_intolerance = False Then

                  Current_Treatment_Sequence = "TS100"

            Else

                  Current_Treatment_Sequence = "TS110"

            End If

      Else
      'Diabetes Type 2 patients

            Select Case SeverityCheck(Patient)
            Case 3
            'HBA1C >=## with (severe wt loss and FBS >300) Severe catabolic manifestations

                  Current_Treatment_Sequence = "TS10"

            Case 2
            'IF ONLY >=##
                  Current_Treatment_Sequence = "TS10"

            Case 1

                  If ASCVD_check(Patient) = True Or .HF Or .CKD Then

                        If .Metformin_intolerance = False Then

                              Current_Treatment_Sequence = "TS20"

                        Else

                              Current_Treatment_Sequence = "TS30"

                        End If

                  Else

                        If .NASH = True Then

                              If .Metformin_intolerance = False Then

                                    Current_Treatment_Sequence = "TS80"

                              Else

                                    Current_Treatment_Sequence = "TS90"

                              End If

                        Else

                              If .BMI < 30 Then

                                    If .Metformin_intolerance = False Then

                                          Current_Treatment_Sequence = "TS60"

                                    Else

                                          Current_Treatment_Sequence = "TS70"

                                    End If

                              Else

                                    If .Metformin_intolerance = False Then

                                          Current_Treatment_Sequence = "TS40"

                                    Else

                                          Current_Treatment_Sequence = "TS50"

                                    End If

                              End If

                        End If

                  End If

            End Select

      End If

            CurrentMedication = Get_Diabetes_Medication(Current_Treatment_Sequence)
            CumulativeToT = Diabetes_Medications(Get_Med_Row(CSng(CurrentMedication(1)))).Durability

            If CumulativeToT > DM_duration Then

                  outcome(1) = Patient.ID
                  outcome(2) = CurrentMedication(2)
                  outcome(3) = CurrentMedication(1)
                  outcome(4) = DM_duration
            Else

JumpBack:

                  'increase the TS code by 1
                  Current_Treatment_Sequence = "TS" & CInt(onlyDigits(Current_Treatment_Sequence)) + 1

                  CurrentMedication = Get_Diabetes_Medication(Current_Treatment_Sequence)
                  CumulativeToT = CumulativeToT + Diabetes_Medications(Get_Med_Row(CSng(CurrentMedication(1)))).Durability

                  If CumulativeToT > DM_duration Then

                        outcome(1) = Patient.ID
                        outcome(2) = CurrentMedication(2)
                        outcome(3) = CurrentMedication(1)
                        outcome(4) = DM_duration - (CumulativeToT - Diabetes_Medications(Get_Med_Row(CSng(CurrentMedication(1)))).Durability)

                  Else

                        GoTo JumpBack

                  End If

            End If


      Else

            'if patient is already on treatment check if the patient is not controlled or need to change medication for other reasons
            'if user provided current treatment in patient data
            'assume half duration of durability
            CurrentMedication = Get_Diabetes_Medication(Patient.Diabetes_Treatment_Sequence)

            outcome(1) = Patient.ID
            outcome(2) = CurrentMedication(2)
            outcome(3) = CurrentMedication(1)
            outcome(4) = (Diabetes_Medications(Get_Med_Row(CSng(CurrentMedication(1)))).Durability) / 2


      End If

End With

Predict_Current_TTT = outcome

End Function
