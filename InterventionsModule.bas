Attribute VB_Name = "InterventionsModule"
Option Explicit

Type intervention

'General information about the intervention
      ID As Byte
      name As String
      Drug As Boolean
      surgical As Boolean
      Int_Complications() As Complication
      Active As Boolean
      BMI_Change() As Single
      mortality As Single
      Utility_Decrement As Double
      Utility_Decrement_Duration As Double 'in month
      Cost As Double
      CostDiscounted As Double
      Maint_Cost As Double
      Maint_Duration As Double
      GLP As Boolean
'Physiological parameters
      HbA1C_Change As Single
      TG_Change As Single
      TC_Change As Single
      HDL_Change As Single
      LDL_Change As Single
      DBP_Change As Single
      SBP_Change As Single
      ALT_Change As Single
      AST_Change As Single
      FBS_Change As Single
      WC_Change As Single
      UA_Change As Single
      HCT_Change As Single
      GGT_Change As Single
      Scr_Change As Single
'Comorbidities remission
      OSA_Remission As Single
      HTN_Remission As Single
      OA_Remission As Single
      DLP_Remission As Single
      CKD_Remission As Single
      NASH_Remission As Single
      
End Type


Function Load_Interventions() As intervention()

'Collection of complications included in the model
Dim Col_Complications() As Complication

Col_Complications = Load_Complications

'general index variables
Dim i As Byte
Dim j As Byte
Dim z As Byte

'Get range where interventions are located
Dim Interventions_Matrix As Variant
Interventions_Matrix = Range("Interventions")

'Get number of interventions
NInterventions = UBound(Interventions_Matrix) - 1

'check how many interventions are active in the model
Dim Temp_Matrix As Variant
Temp_Matrix = Application.Transpose(Interventions_Matrix)
Temp_Matrix = Application.WorksheetFunction.index(Temp_Matrix, 3, 0)
Active_Interventions = ArrayCountIf(Temp_Matrix, True)

ReDim Col_Interventions(1 To Active_Interventions)

'Get range where complications probabilities are located
Dim Complications_Probs_Matrix As Variant
Complications_Probs_Matrix = Range("Complications_Probs")

'Get range where BMI changes are located
Dim BMI_Change_Matrix As Variant
BMI_Change_Matrix = Range("BMI_Change")


'Load Interventions from input parameters (excel)
'set active interventions counter as a start = 1
j = 1

For i = 1 To NInterventions
      
      With Col_Interventions(j)
            
            If CBool(Interventions_Matrix(i, 3)) = True Then
                  'Load information about the intervention
                  .ID = CByte(Interventions_Matrix(i, 1))
                  .name = CStr(Interventions_Matrix(i, 2))
                  .Active = CBool(Interventions_Matrix(i, 3))
                  .Drug = CBool(Interventions_Matrix(i, 4))
                  .surgical = CBool(Interventions_Matrix(i, 5))
                  .mortality = CSng(Interventions_Matrix(i, 6))
                  .Utility_Decrement = CDbl(Interventions_Matrix(i, 7))
                  .Utility_Decrement_Duration = CInt(Interventions_Matrix(i, 8))
                  .Cost = CDbl(Interventions_Matrix(i, 9))
                  .Maint_Cost = CDbl(Interventions_Matrix(i, 10))
                  .Maint_Duration = CDbl(Interventions_Matrix(i, 11))
                  
                  .HbA1C_Change = CSng(Interventions_Matrix(i, 12))
                  .TG_Change = CSng(Interventions_Matrix(i, 13))
                  .TC_Change = CSng(Interventions_Matrix(i, 14))
                  .HDL_Change = CSng(Interventions_Matrix(i, 15))
                  .LDL_Change = CSng(Interventions_Matrix(i, 16))
                  .DBP_Change = CSng(Interventions_Matrix(i, 17))
                  
                  .SBP_Change = CSng(Interventions_Matrix(i, 18))
                  .ALT_Change = CSng(Interventions_Matrix(i, 19))
                  .AST_Change = CSng(Interventions_Matrix(i, 20))
                  .FBS_Change = CSng(Interventions_Matrix(i, 21))
                  .WC_Change = CSng(Interventions_Matrix(i, 22))
                  .UA_Change = CSng(Interventions_Matrix(i, 23))
                  .HCT_Change = CSng(Interventions_Matrix(i, 24))
                  .GGT_Change = CSng(Interventions_Matrix(i, 25))
                  .Scr_Change = CSng(Interventions_Matrix(i, 26))
                  
                  .OSA_Remission = CSng(Interventions_Matrix(i, 27))
                  .HTN_Remission = CSng(Interventions_Matrix(i, 28))
                  .OA_Remission = CSng(Interventions_Matrix(i, 29))
                  .DLP_Remission = CSng(Interventions_Matrix(i, 30))
                  .CKD_Remission = CSng(Interventions_Matrix(i, 31))
                  .NASH_Remission = CSng(Interventions_Matrix(i, 32))
                  .GLP = CBool(Interventions_Matrix(i, 33))
                  .CostDiscounted = CDbl(Interventions_Matrix(i, 34))
                                                
                        'Load data about complications related to each intervention
                        ReDim Col_Interventions(j).Int_Complications(1 To Active_Complications) As Complication

                        For z = 1 To Active_Complications
                              'Load general data about complications from the previous module Load_Complications
                              .Int_Complications(z).Active = Col_Complications(z).Active
                              .Int_Complications(z).name = Col_Complications(z).name
                              .Int_Complications(z).Cost = Col_Complications(z).Cost
                              .Int_Complications(z).ID = Col_Complications(z).ID
                              .Int_Complications(z).Utility_Decrement = Col_Complications(z).Utility_Decrement
                              .Int_Complications(z).Length = Col_Complications(z).Length
                              'load data about probability of complication incidence from the table Complications_Probs
                              Temp_Matrix = FilterArray(Complications_Probs_Matrix, 1, CStr(.ID), 2, CStr(.Int_Complications(z).ID))
                              .Int_Complications(z).Prob6m = Temp_Matrix(1, 3)
                              .Int_Complications(z).Prob6m12m = Temp_Matrix(1, 4)
                              .Int_Complications(z).Prob13m = Temp_Matrix(1, 5)
                        Next z
                        
                        'Load data about BMI change over time for each intervention
                        Temp_Matrix = FilterArray(BMI_Change_Matrix, 1, CStr(.ID))
                        ReDim .BMI_Change(1 To UBound(Temp_Matrix))
                        For z = 1 To UBound(Temp_Matrix)
                              Temp_Matrix = FilterArray(BMI_Change_Matrix, 1, CStr(.ID), 2, CStr(z))
                              .BMI_Change(z) = CSng(Temp_Matrix(1, 3))
                        Next z
                        
                  j = j + 1
                  If j > Active_Interventions Then GoTo Last_Intervention
            End If
            
      End With
Next i

Last_Intervention:

Load_Interventions = Col_Interventions

End Function


