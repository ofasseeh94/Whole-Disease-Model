Attribute VB_Name = "Universum"
Option Explicit

Public Sub Engine()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim EmptyPatient As Patient

'clear old results
ThisWorkbook.Sheets("Patient Outcomes").Range("A3:CV1000003").ClearContents

Dim i As Long, j As Byte, x As Long
Dim startTime As Single

ThisWorkbook.Activate

'Load input data from excel to VBA and PC memory
Call LoadInputs

'number of cycles to run
Dim Cycles As Long
Cycles = Timehorizon / Cycle_Length

'Number of itterations
Dim NInterventions As Byte
NInterventions = UBound(Interventions)
Dim Niterations As Long
Niterations = NPatients * NInterventions
Dim Patient_Diabetes_History As Variant

'redim patient paste array
ReDim PastePatientArray(1 To Niterations, 1 To 100)

'Set the progress bar and the timer settings
startTime = Timer
Call modProgress.ShowProgress(0, Niterations, "Progress: " & 0 & " of " & NPatients & " in the intervention number " & 1 & " of " & NInterventions, False)

'Loop through different interventions
For j = 1 To UBound(Interventions)

      'set the global variable active intervention to be used everywhere
      ActiveIntervention = Interventions(j)
      
      'Reset patient data
      For i = LBound(Patients) To UBound(Patients)
          Patients(i) = EmptyPatient     ' resets all fields to defaults
      Next
      Call Load_Patient_Characteristics
      Call Load_Baseline_QRS
      
      'Loop through patients
      For i = 1 To UBound(Patients)
      
            ReDim Mortality_Arr(0)
      
            With Patients(i)
            
                  'populate baseline characteristics
                  BaseLine_HbA1C = .HbA1C
                  Baseline_HDL = .HDL
                  Baseline_TG = .TG
                  Baseline_TC = .TC
                  Baseline_LDL = .LDL
                  
                  'When undergoing evaluation for obesity intervention diabetic patients are checked for diabetes so diabetes will be detected
                  'It is assumed that patients undergoing any obesity intervention are to be checked for diabetes and it will be eventually detected.
                  If .DM = True And .DM_Treated = True Then
                        
                        Patient_Diabetes_History = Predict_Current_TTT(Patients(i))
                        .Diabetes_Treatment_Sequence = Patient_Diabetes_History(2)
                        .Diabetes_treatment_ID = Patient_Diabetes_History(3)
                        .Diabetes_Drug = Diabetes_Medications(Get_Med_Row(CSng(.Diabetes_treatment_ID)))
                        Time_Start_DM_Medication = -1 * Patient_Diabetes_History(4)
                        
                        'if the patient on a subsquent treatment sequence we apply the addon treatment effect otherwise
                        'if the patient is on the intitial treatment sequence we apply the full treatment effect
                        If onlyDigits(.Diabetes_Treatment_Sequence) Mod 10 = 0 Then
                        
                              treatment_effect_HbA1C = Application.WorksheetFunction.Min(0, Application.WorksheetFunction.Norm_Inv(RandArray(.ID, .time_elapsed * 2, 22), .Diabetes_Drug.HbA1C_Reduction_Mean, .Diabetes_Drug.HbA1C_Reduction_SE))
                        
                        Else
                        
                              treatment_effect_HbA1C = Application.WorksheetFunction.Min(0, Application.WorksheetFunction.Norm_Inv(RandArray(.ID, .time_elapsed * 2, 22), .Diabetes_Drug.HbA1C_Reduction_Addon_Mean, .Diabetes_Drug.HbA1C_Reduction_Addon_SE))
                        
                        End If
                                          
                        'predict the baseline before taking the last diabetes medication
                        'this is being done to populate the HbA1c equation to calculate the future HbA1c given (duration, baseline, reduction)
                        Baseline_Before_Medication_HbA1C = Hba1cReverse(.HbA1C, CSng(Patient_Diabetes_History(4)), treatment_effect_HbA1C, 0.980982, 9.3)
                        
                        'Administrative task to fill first DM age data because this is not part of the user inputs in the patient list
                        If .DM_Diagnosis_Age <> 0 Then .Age_First_DM = .DM_Diagnosis_Age Else .Age_First_DM = .Age
                        
                  End If
                  
                  
                  'Record Intervention in patient data
                  .Intervention_ID = ActiveIntervention.ID
                  
                  'If patient is dead jump out of the loop
                  If .Dead = True Then GoTo NextPatient
                  
                  'Capture risk of death from intervention
                  Mortality_Arr(0) = ActiveIntervention.mortality
                  
                  'One time intervention effect
                  'Costs
                  'accumulate costs
                  .Agg_Cost = .Agg_Cost + ActiveIntervention.CostDiscounted
                  .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(ActiveIntervention.CostDiscounted, Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
                  
                  If ActiveIntervention.surgical = True Then
                                          
                        If .DM = True And .DM_type1 = False Then
                              
                              'Assess preoperative probability of relapse of Diabetes
                              DM_Relapse_Prob = Prob_DM_Relapse(Patients(i), ActiveIntervention)
                                                      
                              'Assess remission of Diabetes
                              Call DM_Remission_Evaluate(Patients(i), ActiveIntervention)
                              
                        End If
                  End If
                        'Assess remission of other comorbidities
                        Call OSA_Remission_Evaluate(Patients(i), ActiveIntervention)
                        Call HTN_Remission_Evaluate(Patients(i), ActiveIntervention)
                        Call OA_Remission_Evaluate(Patients(i), ActiveIntervention)
                        Call DLP_Remission_Evaluate(Patients(i), ActiveIntervention)
                        Call CKD_Remission_Evaluate(Patients(i), ActiveIntervention)
                        Call NASH_Remission_Evaluate(Patients(i), ActiveIntervention)
                        
                 
                  
                  'Intervention impact on physiological parameters. Intervention will affect only abnormal values.
                  If .HbA1C > 5.7 Then .HbA1C = .HbA1C * ActiveIntervention.HbA1C_Change
                  
                  'TG is already managed in the characteristics module and update based on the BMI value
                  'no need to adjust for it here otherwise it will be double counting for the effect of the surgery
                  'If .TG > 150 Then .TG = .TG * ActiveIntervention.TG_Change
                  
                  If .TC > 200 Then .TC = .TC * ActiveIntervention.TC_Change
                  
                  If .Female = True Then
                        If .HDL < 45 Then .HDL = .HDL * ActiveIntervention.HDL_Change
                  Else
                        If .HDL < 35 Then .HDL = .HDL * ActiveIntervention.HDL_Change
                  End If
                  
                  If .LDL > 140 Then .LDL = .LDL * ActiveIntervention.LDL_Change
                  If .DBP > 90 Then .DBP = .DBP * ActiveIntervention.DBP_Change
                  If .SBP > 140 Then .SBP = .SBP * ActiveIntervention.SBP_Change
                  If .ALT > 120 Then .ALT = .ALT * ActiveIntervention.ALT_Change
                  If .AST > 120 Then .AST = .AST * ActiveIntervention.AST_Change
                  If .FBS > 120 Then .FBS = .FBS * ActiveIntervention.FBS_Change
                  If .WC > 100 Then .WC = .WC * ActiveIntervention.WC_Change
                  If .Uric_Acid > 7 Then .Uric_Acid = .Uric_Acid * ActiveIntervention.UA_Change
                  If .HCT > 50 Then .HCT = .HCT * ActiveIntervention.HCT_Change
                  If .GGT > 50 Then .GGT = .GGT * ActiveIntervention.GGT_Change
                  If .Scr > 1.2 Then .Scr = .Scr * ActiveIntervention.Scr_Change
                  
              
                  'loop through cycles
                  For x = 1 To Cycles
                        
                       Sheets("SBP_Tracker").Cells(.time_elapsed * 2 + 2, .ID + 3).Value = .SBP
                       Sheets("DBP_Tracker").Cells(.time_elapsed * 2 + 2, .ID + 3).Value = .DBP
                                                                                                                                                                     
                        If .Dead = True Then GoTo NextPatient
                        
                        ReDim Preserve Mortality_Arr(0)
                        ReDim Disutility_Arr(0)
                        Disutility_Arr(0) = 1
                             
                        'Ongoing intervention effect
                        'Evaluate complications from intervention
                        Call Evaluate_Complications(Patients(i))
                        
                        'Maintenance cost if it is still in the maintenence duration
                        If ActiveIntervention.Maint_Duration > .time_elapsed Then
                        
                              .Agg_Cost = .Agg_Cost + ActiveIntervention.Maint_Cost
                              .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(ActiveIntervention.Maint_Cost, Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
                        
                        End If
                        'Assess relapse
                        Call DM_Relapse_Evaluate(Patients(i), ActiveIntervention)
                        
                        'Loop through comorbidities incidence and update patient physiological parameters
                        Call Characteristics_Progression(Patients(i))
                        
                  Next x
            
            End With
            
NextPatient:
            
            
            'Progress Bar
            If i Mod 97 = 0 Then
            
                  Call modProgress.ShowProgress(i + (j - 1) * NPatients, Niterations, "Progress: " & i & " of " & NPatients & " in the intervention number " & j & " of " & NInterventions & ": " & Format((i + (j - 1) * NPatients) / Niterations, "0%") & ", estimated time left = " & Round((Timer - startTime) / (i + (j - 1) * NPatients) * (Niterations - (i + (j - 1) * NPatients)) / 60, 0) & " minutes", False)
            
            End If
            
      Next i
      
      'PASTE RESULTS
      Call StorePatientArray(Patients, j, NPatients, Niterations)

Next j

'Paste all patient data
Sheets("Patient Outcomes").Range("A3:CV" & Niterations + 2) = PastePatientArray

Range("Last_Run").Value = Now

Debug.Print Timer - startTime

'Unload the progress form
Unload ufProgress

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
