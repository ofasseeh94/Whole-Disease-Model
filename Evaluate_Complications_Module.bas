Attribute VB_Name = "Evaluate_Complications_Module"
Sub Evaluate_Complications(Patient As Patient)

For i = 1 To UBound(ActiveIntervention.Int_Complications)
      
Dim Complication_Rate As Single

'Check which complication rate should be used based on time elapsed
'then check based on a random number of the complication will occur or not
      If Patient.time_elapsed = Cycle_Length Then
            Complication_Rate = ActiveIntervention.Int_Complications(i).Prob6m
      ElseIf Patient.time_elapsed = Cycle_Length * 2 Then
            Complication_Rate = ActiveIntervention.Int_Complications(i).Prob6m12m
      ElseIf Patient.time_elapsed = Cycle_Length * 3 Then
            Complication_Rate = ActiveIntervention.Int_Complications(i).Prob13m
      End If
      
      'Record complication in patient current status
      
      ReDim Preserve Patient.Complication_Status(1 To UBound(ActiveIntervention.Int_Complications))
      
      If Patient.Complication_Status(i).History <> True And Complication_Rate > RandArray(Patient.ID, Patient.time_elapsed / Cycle_Length, 1) Then
      
            With Patient.Complication_Status(i)
            
                  .History = True
                  .ID = ActiveIntervention.Int_Complications(i).ID
                  .name = ActiveIntervention.Int_Complications(i).name
                  .Utility_Decrement = ActiveIntervention.Int_Complications(i).Utility_Decrement
                  .Cost = ActiveIntervention.Int_Complications(i).Cost
                  .Length = ActiveIntervention.Int_Complications(i).Length
                  .Affected = True
                  .FirstOnset = Patient.time_elapsed
                  
            End With
                  
      End If

Next i

End Sub
