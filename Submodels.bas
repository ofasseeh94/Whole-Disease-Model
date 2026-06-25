Attribute VB_Name = "Submodels"
Public Diabetes As Boolean
Public DiabetesTreatment As Single


'Foot Ulcer variables
Public Foot_Ulcer_TPM As Variant
Public Foot_Ulcer_Death_TPM As Variant
Public Foot_Ulcer_PPID As Long
Public Foot_Ulcer_CPID As Long
Public Foot_Ulcer_CHS As Byte
Public Const Foot_Ulcer_Starting_HS As Byte = 2
Public Foot_Ulcer_Cycle As Double
Public Foot_Ulcer_Repeat As Integer
Public Foot_Ulcer_DisU_Costs As Variant

'Nephropathy variables
Public Nephro_TPM As Variant
Public Nephro_Death_TPM As Variant
Public Nephro_PPID As Long
Public Nephro_CPID As Long
Public Nephro_CHS As Byte
Public Const Nephro_Starting_HS As Byte = 1
Public Nephro_Cycle As Double
Public Nephro_Repeat As Integer
Public Nephro_DisU_Costs As Variant

'NASH variables
Public NASH_TPM As Variant
Public NASH_Death_TPM As Variant
Public NASH_PPID As Long
Public NASH_CPID As Long
Public NASH_CHS As Byte
Public Const NASH_Starting_HS As Byte = 1
Public NASH_Cycle As Double
Public NASH_Repeat As Integer
Public NASH_DisU_Costs As Variant

'CHD variables
Public CHD_TPM As Variant
Public CHD_Death_TPM As Variant
Public CHD_PPID As Long
Public CHD_CPID As Long
Public CHD_CHS As Byte
Public Const CHD_Starting_HS As Byte = 1
Public CHD_Cycle As Double
Public CHD_Repeat As Integer
Public CHD_DisU_Costs As Variant

'CKD variables
Public CKD_TPM As Variant
Public CKD_Death_TPM As Variant
Public CKD_PPID As Long
Public CKD_CPID As Long
Public CKD_CHS As Byte
Public Const CKD_Starting_HS As Byte = 1
Public CKD_Cycle As Double
Public CKD_Repeat As Integer
Public CKD_DisU_Costs As Variant

'DKA variables
Public DKA_TPM As Variant
Public DKA_Death_TPM As Variant
Public DKA_PPID As Long
Public DKA_CPID As Long
Public DKA_CHS As Byte
Public Const DKA_Starting_HS As Byte = 1
Public DKA_Cycle As Double
Public DKA_Repeat As Integer
Public DKA_DisU_Costs As Variant

'DLP variables
Public DLP_TPM As Variant
Public DLP_Death_TPM As Variant
Public DLP_PPID As Long
Public DLP_CPID As Long
Public DLP_CHS As Byte
Public Const DLP_Starting_HS As Byte = 1
Public DLP_Cycle As Double
Public DLP_Repeat As Integer
Public DLP_DisU_Costs As Variant

'DM variables
Public DM_TPM As Variant
Public DM_Death_TPM As Variant
Public DM_PPID As Long
Public DM_CPID As Long
Public DM_CHS As Byte
Public Const DM_Starting_HS As Byte = 1
Public DM_Cycle As Double
Public DM_Repeat As Integer
Public DM_DisU_Costs As Variant
Public DM_Relapse_Prob As Double

'HTN variables
Public HTN_TPM As Variant
Public HTN_Death_TPM As Variant
Public HTN_PPID As Long
Public HTN_CPID As Long
Public HTN_CHS As Byte
Public Const HTN_Starting_HS As Byte = 1
Public HTN_Cycle As Double
Public HTN_Repeat As Integer
Public HTN_DisU_Costs As Variant

'MA variables
Public MA_TPM As Variant
Public MA_Death_TPM As Variant
Public MA_PPID As Long
Public MA_CPID As Long
Public MA_CHS As Byte
Public Const MA_Starting_HS As Byte = 1
Public MA_Cycle As Double
Public MA_Repeat As Integer
Public MA_DisU_Costs As Variant

'MI variables
Public MI_TPM As Variant
Public MI_Death_TPM As Variant
Public MI_PPID As Long
Public MI_CPID As Long
Public MI_CHS As Byte
Public Const MI_Starting_HS As Byte = 1
Public MI_Cycle As Double
Public MI_Repeat As Integer
Public MI_DisU_Costs As Variant

'Neuro variables
Public Neuro_TPM As Variant
Public Neuro_Death_TPM As Variant
Public Neuro_PPID As Long
Public Neuro_CPID As Long
Public Neuro_CHS As Byte
Public Const Neuro_Starting_HS As Byte = 1
Public Neuro_Cycle As Double
Public Neuro_Repeat As Integer
Public Neuro_DisU_Costs As Variant

'OA variables
Public OA_TPM As Variant
Public OA_Death_TPM As Variant
Public OA_PPID As Long
Public OA_CPID As Long
Public OA_CHS As Byte
Public Const OA_Starting_HS As Byte = 1
Public OA_Cycle As Double
Public OA_Repeat As Integer
Public OA_DisU_Costs As Variant

'OSA variables
Public OSA_TPM As Variant
Public OSA_Death_TPM As Variant
Public OSA_PPID As Long
Public OSA_CPID As Long
Public OSA_CHS As Byte
Public Const OSA_Starting_HS As Byte = 1
Public OSA_Cycle As Double
Public OSA_Repeat As Integer
Public OSA_DisU_Costs As Variant

'PVD variables
Public PVD_TPM As Variant
Public PVD_Death_TPM As Variant
Public PVD_PPID As Long
Public PVD_CPID As Long
Public PVD_CHS As Byte
Public Const PVD_Starting_HS As Byte = 1
Public PVD_Cycle As Double
Public PVD_Repeat As Integer
Public PVD_DisU_Costs As Variant

'Retino variables
Public Retino_TPM As Variant
Public Retino_Death_TPM As Variant
Public Retino_PPID As Long
Public Retino_CPID As Long
Public Retino_CHS As Byte
Public Const Retino_Starting_HS As Byte = 1
Public Retino_Cycle As Double
Public Retino_Repeat As Integer
Public Retino_DisU_Costs As Variant

'Stroke variables
Public Stroke_TPM As Variant
Public Stroke_Death_TPM As Variant
Public Stroke_PPID As Long
Public Stroke_CPID As Long
Public Stroke_CHS As Byte
Public Const Stroke_Starting_HS As Byte = 1
Public Stroke_Cycle As Double
Public Stroke_Repeat As Integer
Public Stroke_DisU_Costs As Variant

'HypoGly variables
Public HypoGly_TPM As Variant
Public HypoGly_Death_TPM As Variant
Public HypoGly_PPID As Long
Public HypoGly_CPID As Long
Public HypoGly_CHS As Byte
Public Const HypoGly_Starting_HS As Byte = 1
Public HypoGly_Cycle As Double
Public HypoGly_Repeat As Integer
Public HypoGly_DisU_Costs As Variant


Sub Foot_Ulcer(patient As patient)

'Capture patient ID
Foot_Ulcer_CPID = patient.ID

'loop through the model

Utility_Temp = 0

For i = 1 To Foot_Ulcer_Repeat

      With patient
            'if new patient flush previous health state
            If Foot_Ulcer_PPID <> Foot_Ulcer_CPID Then
                  
                  Foot_Ulcer_CHS = Foot_Ulcer_Starting_HS
                  Foot_Ulcer_PPID = Foot_Ulcer_CPID
                  
            Else
            
                  Foot_Ulcer_CHS = NextHS(Foot_Ulcer_TPM, Foot_Ulcer_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 24))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + Foot_Ulcer_DisU_Costs(Foot_Ulcer_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(Foot_Ulcer_DisU_Costs(Foot_Ulcer_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + Foot_Ulcer_Cycle / Cycle_Length * Foot_Ulcer_DisU_Costs(Foot_Ulcer_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - Foot_Ulcer_Death_TPM(Foot_Ulcer_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub Nephro(patient As patient)

'Capture patient ID
Nephro_CPID = patient.ID

'loop through the model
Utility_Temp = 0

For i = 1 To Nephro_Repeat

      'if new patient flush previous health state
      If Nephro_PPID <> Nephro_CPID Then
            
            Nephro_CHS = Nephro_Starting_HS
            Nephro_PPID = Nephro_CPID
                  
      Else
      
            Nephro_CHS = NextHS(Nephro_TPM, Nephro_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 25))
      
      End If
      
      With patient
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + Nephro_DisU_Costs(Nephro_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(Nephro_DisU_Costs(Nephro_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + Nephro_Cycle / Cycle_Length * Nephro_DisU_Costs(Nephro_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - Nephro_Death_TPM(Nephro_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub NASH(patient As patient)

'Capture patient ID
NASH_CPID = patient.ID

'loop through the model

Utility_Temp = 0

For i = 1 To NASH_Repeat

      
      With patient
            
            'if new patient flush previous health state
            If NASH_PPID <> NASH_CPID Then
                     
                  NASH_PPID = NASH_CPID
                  NASH_CHS = NASH_Starting_HS
                        
            Else
            
                  NASH_CHS = NextHS(NASH_TPM, NASH_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 26))
            
            End If
                  
            If NASH_CHS = 5 Then
                        
                  If .Age_First_CC = 0 Then .Age_First_CC = .Age
      
            End If
            
            If NASH_CHS = 8 Then
                  
                  If .Age_First_Transplantation = 0 Then .Age_First_Transplantation = .Age

            End If

                  
            'accumulate costs
            .Agg_Cost = .Agg_Cost + NASH_DisU_Costs(NASH_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(NASH_DisU_Costs(NASH_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + NASH_Cycle / Cycle_Length * NASH_DisU_Costs(NASH_CHS + 1, 2)

      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - NASH_Death_TPM(NASH_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub CHD(patient As patient)

'Capture patient ID
CHD_CPID = patient.ID

'loop through the model

Utility_Temp = 0

For i = 1 To CHD_Repeat
      
      With patient
            'if new patient flush previous health state
            If CHD_PPID <> CHD_CPID Then
                  
                  CHD_CHS = CHD_Starting_HS
                  CHD_PPID = CHD_CPID
                        
            Else
            
                  CHD_CHS = NextHS(CHD_TPM, CHD_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 27))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + CHD_DisU_Costs(CHD_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(CHD_DisU_Costs(CHD_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + CHD_Cycle / Cycle_Length * CHD_DisU_Costs(CHD_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - CHD_Death_TPM(CHD_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub CKD(patient As patient)

'Capture patient ID
CKD_CPID = patient.ID

'loop through the model

Utility_Temp = 0

For i = 1 To CKD_Repeat

      With patient
            'if new patient flush previous health state
            If CKD_PPID <> CKD_CPID Then
                  
                  CKD_CHS = CKD_Starting_HS
                  CKD_PPID = CKD_CPID
                        
            Else
            
                  CKD_CHS = NextHS(CKD_TPM, CKD_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 28))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + CKD_DisU_Costs(CKD_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(CKD_DisU_Costs(CKD_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + CKD_Cycle / Cycle_Length * CKD_DisU_Costs(CKD_CHS + 1, 2)
            
            If CKD_CHS = 7 Then
                  
                  .Dialysis = True
                  If .Age_First_Dialysis = 0 Then .Age_First_Dialysis = .Age

            End If
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - CKD_Death_TPM(CKD_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub DLP(patient As patient)

'Capture patient ID
DLP_CPID = patient.ID

'loop through the model

Utility_Temp = 0

For i = 1 To DLP_Repeat
      
      With patient
            
            'if new patient flush previous health state
            If DLP_PPID <> DLP_CPID Then
                  
                  DLP_CHS = DLP_Starting_HS
                  DLP_PPID = DLP_CPID
                        
            Else
            
                  DLP_CHS = NextHS(DLP_TPM, DLP_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 29))
            
            End If

            'accumulate costs
            .Agg_Cost = .Agg_Cost + DLP_DisU_Costs(DLP_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(DLP_DisU_Costs(DLP_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + DLP_Cycle / Cycle_Length * DLP_DisU_Costs(DLP_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - DLP_Death_TPM(DLP_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub DM(patient As patient)

Dim Temp_Drug_Cost As Double

'Capture patient ID

Utility_Temp = 0

DM_CPID = patient.ID

'loop through the model
For i = 1 To DM_Repeat

      With patient
      
            'if new patient flush previous health state
            If DM_PPID <> DM_CPID Then
                  
                  DM_CHS = DM_Starting_HS
                  DM_PPID = DM_CPID
                        
            Else
            
                  DM_CHS = NextHS(DM_TPM, DM_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 30))
            
            End If
            
            
            If .DM_Treated = True Then Call Diabetes_Treatment_Sequence_Update(patient)
   
            'accumulate costs
            .Agg_Cost = .Agg_Cost + DM_DisU_Costs(DM_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(DM_DisU_Costs(DM_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'add drug cost
            'adjust drug cost to consider not double counting GLP cost if it is already used for obesity
            If ActiveIntervention.GLP And .Diabetes_Drug.GLP And ActiveIntervention.Maint_Duration >= .time_elapsed Then
                  
                  Temp_Drug_Cost = .Diabetes_Drug.Cost_Daily * Cycle_Length * 365 - ActiveIntervention.Maint_Cost
                  
                  If Temp_Drug_Cost < 0 Then Temp_Drug_Cost = 0
            
            Else
            
                  Temp_Drug_Cost = .Diabetes_Drug.Cost_Daily * Cycle_Length * 365
            
            End If
            
            'aggregate discounted and undiscounted cost
            If .DM_Treated = True Then .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue((Temp_Drug_Cost), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            If .DM_Treated = True Then .Agg_Cost = .Agg_Cost + Temp_Drug_Cost
                        
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + DM_Cycle / Cycle_Length * DM_DisU_Costs(DM_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - DM_Death_TPM(DM_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub HTN(patient As patient)

'Capture patient ID

Utility_Temp = 0

HTN_CPID = patient.ID

'loop through the model
For i = 1 To HTN_Repeat

      With patient
            'if new patient flush previous health state
            If HTN_PPID <> HTN_CPID Then
                  
                  HTN_CHS = HTN_Starting_HS
                  HTN_PPID = HTN_CPID
                        
            Else
            
                  HTN_CHS = NextHS(HTN_TPM, HTN_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 31))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + HTN_DisU_Costs(HTN_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(HTN_DisU_Costs(HTN_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + HTN_Cycle / Cycle_Length * HTN_DisU_Costs(HTN_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - HTN_Death_TPM(HTN_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub MA(patient As patient)

'Capture patient ID

Utility_Temp = 0

MA_CPID = patient.ID

'loop through the model
For i = 1 To MA_Repeat

      With patient
            'if new patient flush previous health state
            If MA_PPID <> MA_CPID Then
                  
                  MA_CHS = MA_Starting_HS
                  MA_PPID = MA_CPID
                        
            Else
            
                  MA_CHS = NextHS(MA_TPM, MA_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 32))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + MA_DisU_Costs(MA_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(MA_DisU_Costs(MA_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + MA_Cycle / Cycle_Length * MA_DisU_Costs(MA_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - MA_Death_TPM(MA_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub Neuro(patient As patient)

'Capture patient ID

Utility_Temp = 0

Neuro_CPID = patient.ID

'loop through the model
For i = 1 To Neuro_Repeat

      With patient
            
            'if new patient flush previous health state
            If Neuro_PPID <> Neuro_CPID Then
                  
                  Neuro_CHS = Neuro_Starting_HS
                  Neuro_PPID = Neuro_CPID
                        
            Else
            
                  Neuro_CHS = NextHS(Neuro_TPM, Neuro_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 33))
            
            End If
            
            'accumulate costs
            .Agg_Cost = .Agg_Cost + Neuro_DisU_Costs(Neuro_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(Neuro_DisU_Costs(Neuro_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + Neuro_Cycle / Cycle_Length * Neuro_DisU_Costs(Neuro_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - Neuro_Death_TPM(Neuro_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub OA(patient As patient)

'Capture patient ID

Utility_Temp = 0

OA_CPID = patient.ID

'loop through the model
For i = 1 To OA_Repeat

      With patient
            'if new patient flush previous health state
            If OA_PPID <> OA_CPID Then
                  
                  OA_CHS = OA_Starting_HS
                  OA_PPID = OA_CPID
                        
            Else
            
                  OA_CHS = NextHS(OA_TPM, OA_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 34))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + OA_DisU_Costs(OA_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(OA_DisU_Costs(OA_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + OA_Cycle / Cycle_Length * OA_DisU_Costs(OA_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - OA_Death_TPM(OA_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub OSA(patient As patient)

'Capture patient ID

Utility_Temp = 0

OSA_CPID = patient.ID

'loop through the model
For i = 1 To OSA_Repeat
      
      With patient
            
            'if new patient flush previous health state
            If OSA_PPID <> OSA_CPID Then
                  
                  OSA_CHS = OSA_Starting_HS
                  OSA_PPID = OSA_CPID
                        
            Else
            
                  OSA_CHS = NextHS(OSA_TPM, OSA_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 35))
            
            End If
            
            'accumulate costs
            .Agg_Cost = .Agg_Cost + OSA_DisU_Costs(OSA_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(OSA_DisU_Costs(OSA_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + OSA_Cycle / Cycle_Length * OSA_DisU_Costs(OSA_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - OSA_Death_TPM(OSA_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub PVD(patient As patient)

'Capture patient ID

Utility_Temp = 0

PVD_CPID = patient.ID

'loop through the model
For i = 1 To PVD_Repeat
      
      With patient
            
            'if new patient flush previous health state
            If PVD_PPID <> PVD_CPID Then
                  
                  PVD_CHS = PVD_Starting_HS
                  PVD_PPID = PVD_CPID
                        
            Else
            
                  PVD_CHS = NextHS(PVD_TPM, PVD_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 36))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + PVD_DisU_Costs(PVD_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(PVD_DisU_Costs(PVD_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + PVD_Cycle / Cycle_Length * PVD_DisU_Costs(PVD_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - PVD_Death_TPM(PVD_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub Retino(patient As patient)

'Capture patient ID

Utility_Temp = 0

Retino_CPID = patient.ID

'loop through the model
For i = 1 To Retino_Repeat
      
      With patient
            'if new patient flush previous health state
            If Retino_PPID <> Retino_CPID Then
                  
                  Retino_CHS = Retino_Starting_HS
                  Retino_PPID = Retino_CPID
                        
            Else
            
                  Retino_CHS = NextHS(Retino_TPM, Retino_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 37))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + Retino_DisU_Costs(Retino_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(Retino_DisU_Costs(Retino_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + Retino_Cycle / Cycle_Length * Retino_DisU_Costs(Retino_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - Retino_Death_TPM(Retino_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub
Sub Stroke(patient As patient)

'Capture patient ID

Utility_Temp = 0

Stroke_CPID = patient.ID

'loop through the model
For i = 1 To Stroke_Repeat
      
      With patient
      
            'if new patient flush previous health state
            If Stroke_PPID <> Stroke_CPID Then
                  
                  Stroke_CHS = Stroke_Starting_HS
                  Stroke_PPID = Stroke_CPID
                        
            Else
            
                  Stroke_CHS = NextHS(Stroke_TPM, Stroke_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 38))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + Stroke_DisU_Costs(Stroke_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(Stroke_DisU_Costs(Stroke_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + Stroke_Cycle / Cycle_Length * Stroke_DisU_Costs(Stroke_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - Stroke_Death_TPM(Stroke_CHS + 1, 2)) ^ (1 / Cycle_Length)

End Sub


Sub MI(patient As patient)

'Capture patient ID

Utility_Temp = 0

MI_CPID = patient.ID

'loop through the model
For i = 1 To MI_Repeat

      With patient
            'if new patient flush previous health state
            If MI_PPID <> MI_CPID Then
                  
                  MI_CHS = MI_Starting_HS
                  MI_PPID = MI_CPID
                        
            Else
            
                  MI_CHS = NextHS(MI_TPM, MI_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 39))
            
            End If
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + MI_DisU_Costs(MI_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(MI_DisU_Costs(MI_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + MI_Cycle / Cycle_Length * MI_DisU_Costs(MI_CHS + 1, 2)
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - MI_Death_TPM(MI_CHS + 1, 2)) ^ (1 / Cycle_Length)


End Sub

Sub HypoGly(patient As patient)

'Always start patient with hypoglycemia
HypoGly_CHS = HypoGly_Starting_HS

'loop through the model

Utility_Temp = 0

For i = 1 To HypoGly_Repeat
                  
      With patient
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + HypoGly_DisU_Costs(HypoGly_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(HypoGly_DisU_Costs(HypoGly_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + HypoGly_Cycle / Cycle_Length * HypoGly_DisU_Costs(HypoGly_CHS + 1, 2)
      
            HypoGly_CHS = NextHS(HypoGly_TPM, HypoGly_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 40))
      
      End With
Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - HypoGly_Death_TPM(HypoGly_CHS + 1, 2)) ^ (1 / Cycle_Length)


End Sub

Sub DKA(patient As patient)

'Always start patient with DKA
DKA_CHS = DKA_Starting_HS

'loop through the model

Utility_Temp = 0

For i = 1 To DKA_Repeat
                  
      With patient
      
            'accumulate costs
            .Agg_Cost = .Agg_Cost + DKA_DisU_Costs(DKA_CHS + 1, 3)
            .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(DKA_DisU_Costs(DKA_CHS + 1, 3)), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
            
            'Add utilities to the aggregator module
            Utility_Temp = Utility_Temp + DKA_Cycle / Cycle_Length * DKA_DisU_Costs(DKA_CHS + 1, 2)
            
            DKA_CHS = NextHS(DKA_TPM, DKA_CHS, RandArray(.ID, .time_elapsed / Cycle_Length, 41))
      
      End With

Next i

'Capture disutility caused in this submodel
ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
Disutility_Arr(UBound(Disutility_Arr)) = Utility_Temp

'Capture probability of death
ReDim Preserve Mortality_Arr(UBound(Mortality_Arr) + 1)
Mortality_Arr(UBound(Mortality_Arr)) = 1 - (1 - DKA_Death_TPM(DKA_CHS + 1, 2)) ^ (1 / Cycle_Length)


End Sub
Function NextHS(TPMatrix As Variant, CurrentHS As Byte, Optional RandomNumber As Single) As Byte

      NextHS = Application.WorksheetFunction.Match(RandomNumber, Application.index(TPMatrix, (CurrentHS + 1), 0), 1) - 1

End Function
