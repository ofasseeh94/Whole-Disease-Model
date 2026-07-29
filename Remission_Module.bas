Attribute VB_Name = "Remission_Module"
Option Explicit

Function Prob_DM_Rem(patient As patient, intervention As intervention)

With patient

Dim i As Integer
Dim NumDiabMeds As Integer

      For i = 1 To UBound(Diabetes_Medications_Matrix)
      
            If Diabetes_Medications(i).ID = .Diabetes_treatment_ID Then
            
            NumDiabMeds = Diabetes_Medications_Matrix(i, 17)
            
            GoTo JumpOut
            End If
            
      Next i

JumpOut:

Dim DiabMedsCof As Single

If NumDiabMeds >= 2 Then
      DiabMedsCof = 2.599
Else
      DiabMedsCof = 2.039
End If

'Source: Plaeke, P., Beunis, A., Ruppert, M. et al. Review, Performance Comparison, and Validation of Models Predicting Type 2 Diabetes Remission After Bariatric Surgery in a Western European Population. OBES SURG 31, 1549–1560 (2021). https://doi.org/10.1007/s11695-020-05157-0
Prob_DM_Rem = 5.707 - 0.153 * (.Age - .DM_Diagnosis_Age) - 0.276 * .HbA1C - 1.434 * Abs(.Insulin) - DiabMedsCof

Prob_DM_Rem = Exp(Prob_DM_Rem)
Prob_DM_Rem = Prob_DM_Rem / (1 + Prob_DM_Rem)

End With

End Function

Function Prob_DM_Relapse(patient As patient, intervention As intervention) As Double
'Source: Aminian A, Vidal J, Salminen P, Still CD, Nor Hanipah Z, Sharma G, Tu C, Wood GC, Ibarzabal A, Jimenez A, Brethauer SA. Late relapse of diabetes after bariatric surgery: not rare, but not a failure. Diabetes care. 2020 Mar 1;43(3):534-40.
With patient

Dim i As Integer
Dim NumDiabMeds As Integer

      For i = 1 To UBound(Diabetes_Medications_Matrix)
      
            If Diabetes_Medications(i).ID = .Diabetes_treatment_ID Then
            
            NumDiabMeds = Diabetes_Medications_Matrix(i, 17)
            
            GoTo JumpOut
            End If
            
      Next i

JumpOut:

Dim SurgeryCof As Byte

If intervention.name = "laparoscopic sleeve gastrectomy" Then SurgeryCof = 1

      Prob_DM_Relapse = -2.4917 + (0.0789 * (.Age - .DM_Diagnosis_Age)) + (0.714 * NumDiabMeds) + (0.7929 * SurgeryCof)

End With

Prob_DM_Relapse = Exp(Prob_DM_Relapse) / (1 + Exp(Prob_DM_Relapse))
Prob_DM_Relapse = 1 - (1 - Prob_DM_Relapse) ^ (1 / 16)

End Function


Sub DM_Remission_Evaluate(patient As patient, intervention As intervention)

Dim EmptyDrug As Diabetes_Medication

With patient
                        
      If Prob_DM_Rem(patient, intervention) > RandArray(.ID, .time_elapsed / Cycle_Length, 42) Then
      
            .DM_Remission = True
            .DM = False
            .Diabetes_treatment_ID = 0
            .Diabetes_Treatment_Sequence = 0
            .Diabetes_Drug = EmptyDrug
            .Diabetes_Drug.ID = 0
            .DM_Treated = False
            .Insulin = False
            .DM_recognized = False
                  
      End If

End With

End Sub

Sub DM_Relapse_Evaluate(patient As patient, intervention As intervention)

With patient

      If .DM = False And .DM_Remission = True And .time_elapsed < 8 Then
            
            If DM_Relapse_Prob > RandArray(.ID, .time_elapsed / Cycle_Length, 43) Then
            
                  .DM_Remission = False
                  .DM = True
                  'call diabetes medication update
                  
            End If
            
      End If

End With

End Sub

Sub OSA_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .OSA = True Then
      
            If intervention.OSA_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 46) Then
            
                  .OSA = False
            
            End If
      
      End If

End With

End Sub

Sub HTN_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .Hypertension = True Then
      
            If intervention.HTN_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 47) Then
            
                  .Hypertension = False
            
            End If
      
      End If

End With

End Sub

Sub OA_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .OA = True Then
      
            If intervention.OA_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 48) Then
            
                  .OA = False
            
            End If
      
      End If

End With

End Sub

Sub DLP_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .DLP = True Then
      
            If intervention.DLP_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 49) Then
            
                  .DLP = False
            
            End If
      
      End If

End With

End Sub
Sub CKD_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .CKD = True Then
      
            If intervention.CKD_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 50) Then
            
                  .CKD = False
            
            End If
      
      End If

End With

End Sub
Sub NASH_Remission_Evaluate(patient As patient, intervention As intervention)

With patient

      If .NASH = True Then
      
            If intervention.NASH_Remission > RandArray(.ID, .time_elapsed / Cycle_Length, 51) Then
            
                  .NASH = False
            
            End If
      
      End If

End With

End Sub
