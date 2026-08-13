Attribute VB_Name = "PatientModule"
Option Explicit

'the type below is be part of the patient characteristics
Type Complication_Status
'Later we need to check if we need to store all the details below
'complication name
name As String
'complication ID
ID As Integer
'complication happend before or not
History As Boolean
'utility decrement
Utility_Decrement As Double
'Length of complication in weeks
Length As Single
'Complication cost
Cost As Single
'Is the patient currently affected by the complication
Affected As Boolean
'First date of onset of the complication
FirstOnset As Single

End Type

Type Patient                              ' Create user-defined type.
      
      'Patient characteristics
      ID As Long                          'Define elements of data type.
      Age As Single                       'Age at diagnosis of diabetes, per year
      BMI As Single                       '=Weight/hight in meters squared
      BMI_Baseline As Single              'BMI at baseline
      Female As Boolean                   'True = female, False = Male
      Dead As Boolean                     'True = dead, false = alive
      Intervention_ID As Byte             'Intervention ID to be able to link both tables
      
      'Patient behavioural characteristics
      smoking As Boolean                  'True = smoker, False = nonsmoker
      Number_of_cigaretts As Byte          ' Number of cigarrets per day
      smoking_cess As Boolean              'True = joined smoking cessation program, false= didn't join smoking cessation program
      smoking_cess_success As Boolean          'True= succeeded in quitting smoking, False= didn't succeed
      rehab_drinking As Boolean            'True=Rehabilitation from drinking , False= no rehabilitation
      rehab_drinking_success As Boolean            'True= Rehabilitation succeeded , False= Rehabilitation failed
      drinking As Boolean                 'True = Alcohol consumption, False = no Alcohol consumption
      Number_of_alcohol As Byte            ' number of alcohol galsses per day
      alcohol_categories As String        'drinks alcohol daily, drinks alcohol almost daily, drinks alcohol 3-4 times a week, drinks alcohol 2 times a week. drinks alcohol once a week. drinks alcohol 2-3 times a month. drinks alcohol once a month. drinks alcohol 7-11 times a year. drinks alcohol 3-6 times a year. drinks alcohol 1-2 times a year.doesn’t drink alcohol
      
      'Lab
      time_elapsed As Double                       'Duration the patient spent in the model
      AF As Boolean                       'Atrial fibrillation
      SBP As Single                       'Systolic blood pressure  mmHg
      DBP As Single                       'Diastolic blood pressure  mmHg
      WC As Single                        'Waist circumfernce in cm
      occupational_risk_OA As String      'Categorical (never, seldom, sometimes, often, always)
      family_history_OA As Boolean        'True=family history of OA ,False= No history
      knee_injury As Boolean              'True=history of knee injury ,False= No history
      age_menopause As Single             'Female age at menopause
      Menopause As Boolean                'True= menopause occured, False=menopause didn't occur
      Metformin_intolerance As Boolean    'True=Metformin intolerance occured, False=Metformin intolerance didn't occur
      
      'Medications
      Diabetes_Treatment_Sequence As String       'ID of the treatment sequence
      Diabetes_treatment_ID As Integer        ' the class of DM medication
      Diabetes_Drug As Diabetes_Medication 'All details about diabetes medication
      'Sulfonylurea As Boolean                'using SU or not
      'Metformin As Boolean                   'using Met or not
      'Sulfonylurea_Pioglitazone As Boolean   'using the combination or not
      'Sulfonylurea_DDP4I As Boolean          'using the combination or not
      'Sulfonylurea_GLP1 As Boolean           'using the combination or not
      'Metformin_Pioglitazone As Boolean      'using the combination or not
      'Metformin_DDP4I As Boolean             'using the combination or not
      'Metformin_GLP1 As Boolean              'using the combination or not
      'Sulfonylurea_basal_insulin As Boolean  'using the combination or not
      'Metformin_basal_insulin As Boolean     'using the combination or not
      'Multiple_doses_insulin As Boolean      'using it or not
      Insulin As Boolean                      ' if the patient is using insulin
      DM_Remission As Boolean                 'if true means this patient was diabetic and went into remission
      
      'Biomedical Markers
      TG   As Single                      'Triglycerides mg/dL
      TC  As Single                       'total cholesterol in mg/dL
      LDL   As Single                     'Low density lipoprotein mg/dL
      HDL   As Single                     'high density lipoprotein mg/dL
      HbA1C  As Single                    'Hemoglubin A 1 c
      FBS As Single                       'fasting blood glucose mg/dl
      Uric_Acid As Single                  'uric acid in blood mg/dl
      ALT As Single                        'ALT in IU/L
      AST As Single                        'AST in IU/L
      HCT As Single                        'Hematocrit measured in %
      GGT As Single                         ' gamma-glutamyl transferase SAME AS Gamma-glutamyl transpeptidase in IU/L
      Scr As Double                         ' serum creatinine in mg/dl
       
      
      'Diabetes
      DM As Boolean                        'True= Has diabetes , Fasle= NO diabtes
      DM_Diagnosis_Age As Single          'Age at diagnosis of diabetes, per year
      DM_recognized As Boolean             'True=Patients know that they are diabetic , False= patients don't know
      DM_Treated As Boolean                'True = Pateint is on treatment for DM, False = patient isn't on treatment
      DM_type1 As Boolean                  ' TRUE= type 1 , false= type 2
      
      'Comorbidities
      DLP As Boolean                   ' true= hyperlipdemia present, false=not present      *************do we use this??
      ulcer_amput_history As Boolean        ' true= any history of ulcer or amuptation , false= no ulcer or amputation
      Ulcer As Boolean                      ' true=present , false= absent
      ASCVD As Boolean                      ' if patient has risk of ASCVD

      Nephro As Boolean                     ' true= diabetic nephropathy present, false=not present
      Hypertension As Boolean              'calculated from SBP & DBP,if SBP>=140 AND/OR DBP>=90
      anti_htn_drugs As Boolean            'True=On Htn drugs, False= not on HTN drugs
      
      'anti_htn_regular As Boolean
      CHD As Boolean                       'True= has cronary heart disease, False= No CHD
      LVH As Boolean                       'True= has left venticular hipertrophy, False= NO LVH
      physical_activity As Boolean         'true=enough physical activity, false= not enough physical activity
      Daily_fruit_consumption As Boolean   'true= enough fruit consumption, false= not enough fruit consumption
      'High_sugar_hist As Boolean          'True= history of high blood sugar , false= no history
      previousHypoGly As Boolean           ' if he had previous epdisode of hypoglycemia
      Keto  As Boolean                      ' if patient has keto
      'Hight_TC_ecognized As Boolean       'True=Patients know that they have high total cholesterol , False= patients don't know
      PVD As Boolean                       'Patient as peripheral vascular disease
      OA As Boolean                        'Patient has osteoarthritis
      OSA As Boolean                        'if patient has obstructive sleep apnea
      MA As Boolean                         ' if patient has macular edema
      Retino As Boolean                     ' if the patient has diabetic  retinopathy
      Neuro As Boolean                      ' if the patient has diabetic neuropathy
      HypoGly As Boolean                    ' does the patient has hypoglycemia
      Keto_history As Boolean               'Patient had a previous ketoacidosis or not
      NASH As Boolean                       ' patient has NASH or not
      Stroke As Boolean                     ' stroke present or not
      Stroke_history As Boolean             ' previous stroke present or not
      MI As Boolean                         'MI present or not
      MI_history As Boolean                 'previous MI present or not
      HF As Boolean                          'HF present or not
      CKD As Boolean                        'CKD present or not
      Dialysis As Boolean                   'Dialysis patients
      
      'Family history
       family_history_DM As Boolean         ' True= history of DM in family, false= no history
       First_degree_hist_DM As Boolean      'True= history of first degree relatives with DM , false= no history
       Second_degree_hist_DM As Boolean     'True= history of first degree relatives with DM , false= no history
       family_hist_CHD As Boolean           'True=Family history of CHD, False= Family history of CHD
       
       
       'Age of onset of diseases
      Age_First_DM As Single
      Age_First_CHD As Single              'Age as first coronary event
      Age_First_Ulcer As Single
      Age_First_Amputation As Single
      Age_First_MI As Single
      Age_First_Stroke As Single
      Age_First_NASH As Single
      Age_First_OA As Single
      Age_First_OSA As Single
      Age_First_Neuro As Single
      Age_First_HTN As Single
      Age_First_PVD As Single
      Age_First_Retino As Single
      Age_First_CKD As Single
      Age_First_Dialysis As Single
      Age_First_CC As Single
      Age_First_Transplantation As Single 'Liver transplantation
       
       
      Agg_Cost As Double
      Agg_Cost_Disc As Double
      Agg_QALYs As Double
      Agg_QALYs_Disc As Double
      
      'Complications
      Complication_Status() As Complication_Status
      
      'Minor_bleeding As Boolean
      'Major_bleeding As Boolean
      'Intestinal_leak As Boolean
      'Intestinal_perforation As Boolean
      'Intestinal_obstruction As Boolean
      'Cholelithiasis As Boolean
      'Internal_hernia As Boolean
      'Incisional_hernia As Boolean
      'Abdominal_abscess As Boolean
      'GERD As Boolean
      'Barretts_esophagus As Boolean
      'Esophageal_cc As Boolean

End Type
Sub PreparePatientCohort()

'load patient cohort as a matrix
Patient_Cohort_Matrix = Range("Patient_Cohort_Table")

'set the patients array and the number of patients
'get the number of patients
'Dim N_Patients As Long
NPatients = UBound(Patient_Cohort_Matrix) - 1
ReDim Patients(1 To NPatients)

End Sub
Sub Load_Patient_Characteristics()

'loop through patient characteristics
Dim i As Long
Dim j As Long

'loop through patients to load their data
            For i = 2 To UBound(Patients) + 1
      
                  With Patients(i - 1)
                  
                        .ID = Patient_Cohort_Matrix(i, 1)
                        .Age = Patient_Cohort_Matrix(i, 2)
                        .BMI = Patient_Cohort_Matrix(i, 3)
                        .BMI_Baseline = Patient_Cohort_Matrix(i, 4)
                        .Female = Patient_Cohort_Matrix(i, 5)
                        .Dead = Patient_Cohort_Matrix(i, 6)
                        .Intervention_ID = Patient_Cohort_Matrix(i, 7)
                        .smoking = Patient_Cohort_Matrix(i, 8)
                        .Number_of_cigaretts = Patient_Cohort_Matrix(i, 9)
                        .smoking_cess = Patient_Cohort_Matrix(i, 10)
                        .smoking_cess_success = Patient_Cohort_Matrix(i, 11)
                        .rehab_drinking = Patient_Cohort_Matrix(i, 12)
                        .rehab_drinking_success = Patient_Cohort_Matrix(i, 13)
                        .drinking = Patient_Cohort_Matrix(i, 14)
                        .Number_of_alcohol = Patient_Cohort_Matrix(i, 15)
                        .alcohol_categories = Patient_Cohort_Matrix(i, 16)
                        .time_elapsed = Patient_Cohort_Matrix(i, 17)
                        .AF = Patient_Cohort_Matrix(i, 18)
                        .SBP = Patient_Cohort_Matrix(i, 19)
                        .DBP = Patient_Cohort_Matrix(i, 20)
                        .WC = Patient_Cohort_Matrix(i, 21)
                        .occupational_risk_OA = Patient_Cohort_Matrix(i, 22)
                        .family_history_OA = Patient_Cohort_Matrix(i, 23)
                        .knee_injury = Patient_Cohort_Matrix(i, 24)
                        .age_menopause = Patient_Cohort_Matrix(i, 25)
                        .Menopause = Patient_Cohort_Matrix(i, 26)
                        .Metformin_intolerance = Patient_Cohort_Matrix(i, 27)
                        .Diabetes_Treatment_Sequence = Patient_Cohort_Matrix(i, 28)
                        .Diabetes_treatment_ID = Patient_Cohort_Matrix(i, 29)
                        .Insulin = Patient_Cohort_Matrix(i, 30)
                        .TG = Patient_Cohort_Matrix(i, 31)
                        .TC = Patient_Cohort_Matrix(i, 32)
                        .LDL = Patient_Cohort_Matrix(i, 33)
                        .HDL = Patient_Cohort_Matrix(i, 34)
                        .HbA1C = Patient_Cohort_Matrix(i, 35)
                        .FBS = Patient_Cohort_Matrix(i, 36)
                        .Uric_Acid = Patient_Cohort_Matrix(i, 37)
                        .ALT = Patient_Cohort_Matrix(i, 38)
                        .AST = Patient_Cohort_Matrix(i, 39)
                        .HCT = Patient_Cohort_Matrix(i, 40)
                        .GGT = Patient_Cohort_Matrix(i, 41)
                        .Scr = Patient_Cohort_Matrix(i, 42)
                        .DM = Patient_Cohort_Matrix(i, 43)
                        .DM_Diagnosis_Age = Patient_Cohort_Matrix(i, 44)
                        .DM_recognized = Patient_Cohort_Matrix(i, 45)
                        .DM_Treated = Patient_Cohort_Matrix(i, 46)
                        .Age_First_CHD = Patient_Cohort_Matrix(i, 47)
                        .DM_type1 = Patient_Cohort_Matrix(i, 48)
                        .DLP = Patient_Cohort_Matrix(i, 49)
                        .ulcer_amput_history = Patient_Cohort_Matrix(i, 50)
                        .Ulcer = Patient_Cohort_Matrix(i, 51)
                        .ASCVD = Patient_Cohort_Matrix(i, 52)
                        .Nephro = Patient_Cohort_Matrix(i, 53)
                        .Hypertension = Patient_Cohort_Matrix(i, 54)
                        .anti_htn_drugs = Patient_Cohort_Matrix(i, 55)
                        .CHD = Patient_Cohort_Matrix(i, 56)
                        .LVH = Patient_Cohort_Matrix(i, 57)
                        .physical_activity = Patient_Cohort_Matrix(i, 58)
                        .Daily_fruit_consumption = Patient_Cohort_Matrix(i, 59)
                        .previousHypoGly = Patient_Cohort_Matrix(i, 60)
                        .Keto = Patient_Cohort_Matrix(i, 61)
                        .PVD = Patient_Cohort_Matrix(i, 62)
                        .OA = Patient_Cohort_Matrix(i, 63)
                        .OSA = Patient_Cohort_Matrix(i, 64)
                        .MA = Patient_Cohort_Matrix(i, 65)
                        .Retino = Patient_Cohort_Matrix(i, 66)
                        .Neuro = Patient_Cohort_Matrix(i, 67)
                        .HypoGly = Patient_Cohort_Matrix(i, 68)
                        .Keto_history = Patient_Cohort_Matrix(i, 69)
                        .NASH = Patient_Cohort_Matrix(i, 70)
                        .Stroke = Patient_Cohort_Matrix(i, 71)
                        .Stroke_history = Patient_Cohort_Matrix(i, 72)
                        .MI = Patient_Cohort_Matrix(i, 73)
                        .MI_history = Patient_Cohort_Matrix(i, 74)
                        .HF = Patient_Cohort_Matrix(i, 75)
                        .CKD = Patient_Cohort_Matrix(i, 76)
                        .Dialysis = Patient_Cohort_Matrix(i, 77)
                        .family_history_DM = Patient_Cohort_Matrix(i, 78)
                        .First_degree_hist_DM = Patient_Cohort_Matrix(i, 79)
                        .Second_degree_hist_DM = Patient_Cohort_Matrix(i, 80)
                        .family_hist_CHD = Patient_Cohort_Matrix(i, 81)
                        .Agg_Cost = Patient_Cohort_Matrix(i, 82)
                        .Agg_Cost_Disc = Patient_Cohort_Matrix(i, 83)
                        .Agg_QALYs = Patient_Cohort_Matrix(i, 84)
                        .Agg_QALYs_Disc = Patient_Cohort_Matrix(i, 85)
                        .LDL = (.TC / 0.948) - (.HDL / 0.971) - ((.TG / 8.56) + (.TG * (.TC - .HDL) / 2140) - (.TG ^ 2) / 16100) - 9.44

                        If .Diabetes_Treatment_Sequence <> "0" Then Call Get_Medication_Row(Patients(i - 1))
                        
                  End With
            
            Next
            
      'Load_Patient_Characteristics = Patients

End Sub

Public Sub Paste_Patient_Characteristics(Patients() As Patient, InterventionCount As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Patient Outcomes")
    Dim i As Long
    Dim row As Long
    row = 3 + (InterventionCount - 1) * UBound(Patients)

    For i = LBound(Patients) To UBound(Patients)
        With ws
            .Cells(row, 1).Value = Patients(i).ID
            .Cells(row, 2).Value = Patients(i).Age
            .Cells(row, 3).Value = Patients(i).BMI
            .Cells(row, 4).Value = Patients(i).BMI_Baseline
            .Cells(row, 5).Value = Patients(i).Female
            .Cells(row, 6).Value = Patients(i).Dead
            .Cells(row, 7).Value = Patients(i).Intervention_ID
            .Cells(row, 8).Value = Patients(i).smoking
            .Cells(row, 9).Value = Patients(i).Number_of_cigaretts
            .Cells(row, 10).Value = Patients(i).smoking_cess
            .Cells(row, 11).Value = Patients(i).smoking_cess_success
            .Cells(row, 12).Value = Patients(i).rehab_drinking
            .Cells(row, 13).Value = Patients(i).rehab_drinking_success
            .Cells(row, 14).Value = Patients(i).drinking
            .Cells(row, 15).Value = Patients(i).Number_of_alcohol
            .Cells(row, 16).Value = Patients(i).alcohol_categories
            .Cells(row, 17).Value = Patients(i).time_elapsed
            .Cells(row, 18).Value = Patients(i).AF
            .Cells(row, 19).Value = Patients(i).SBP
            .Cells(row, 20).Value = Patients(i).DBP
            .Cells(row, 21).Value = Patients(i).WC
            .Cells(row, 22).Value = Patients(i).occupational_risk_OA
            .Cells(row, 23).Value = Patients(i).family_history_OA
            .Cells(row, 24).Value = Patients(i).knee_injury
            .Cells(row, 25).Value = Patients(i).age_menopause
            .Cells(row, 26).Value = Patients(i).Menopause
            .Cells(row, 27).Value = Patients(i).Metformin_intolerance
            .Cells(row, 28).Value = Patients(i).Diabetes_Treatment_Sequence
            .Cells(row, 29).Value = Patients(i).Diabetes_treatment_ID
            .Cells(row, 30).Value = Patients(i).Insulin
            .Cells(row, 31).Value = Patients(i).TG
            .Cells(row, 32).Value = Patients(i).TC
            .Cells(row, 33).Value = Patients(i).LDL
            .Cells(row, 34).Value = Patients(i).HDL
            .Cells(row, 35).Value = Patients(i).HbA1C
            .Cells(row, 36).Value = Patients(i).FBS
            .Cells(row, 37).Value = Patients(i).Uric_Acid
            .Cells(row, 38).Value = Patients(i).ALT
            .Cells(row, 39).Value = Patients(i).AST
            .Cells(row, 40).Value = Patients(i).HCT
            .Cells(row, 41).Value = Patients(i).GGT
            .Cells(row, 42).Value = Patients(i).Scr
            .Cells(row, 43).Value = Patients(i).DM
            .Cells(row, 44).Value = Patients(i).DM_Diagnosis_Age
            .Cells(row, 45).Value = Patients(i).DM_recognized
            .Cells(row, 46).Value = Patients(i).DM_Treated
            .Cells(row, 47).Value = Patients(i).Age_First_CHD
            .Cells(row, 48).Value = Patients(i).DM_type1
            .Cells(row, 49).Value = Patients(i).DLP
            .Cells(row, 50).Value = Patients(i).ulcer_amput_history
            .Cells(row, 51).Value = Patients(i).Ulcer
            .Cells(row, 52).Value = Patients(i).ASCVD
            .Cells(row, 53).Value = Patients(i).Nephro
            .Cells(row, 54).Value = Patients(i).Hypertension
            .Cells(row, 55).Value = Patients(i).anti_htn_drugs
            .Cells(row, 56).Value = Patients(i).CHD
            .Cells(row, 57).Value = Patients(i).LVH
            .Cells(row, 58).Value = Patients(i).physical_activity
            .Cells(row, 59).Value = Patients(i).Daily_fruit_consumption
            .Cells(row, 60).Value = Patients(i).previousHypoGly
            .Cells(row, 61).Value = Patients(i).Keto
            .Cells(row, 62).Value = Patients(i).PVD
            .Cells(row, 63).Value = Patients(i).OA
            .Cells(row, 64).Value = Patients(i).OSA
            .Cells(row, 65).Value = Patients(i).MA
            .Cells(row, 66).Value = Patients(i).Retino
            .Cells(row, 67).Value = Patients(i).Neuro
            .Cells(row, 68).Value = Patients(i).HypoGly
            .Cells(row, 69).Value = Patients(i).Keto_history
            .Cells(row, 70).Value = Patients(i).NASH
            .Cells(row, 71).Value = Patients(i).Stroke
            .Cells(row, 72).Value = Patients(i).Stroke_history
            .Cells(row, 73).Value = Patients(i).MI
            .Cells(row, 74).Value = Patients(i).MI_history
            .Cells(row, 75).Value = Patients(i).HF
            .Cells(row, 76).Value = Patients(i).CKD
            .Cells(row, 77).Value = Patients(i).Dialysis
            .Cells(row, 78).Value = Patients(i).family_history_DM
            .Cells(row, 79).Value = Patients(i).First_degree_hist_DM
            .Cells(row, 80).Value = Patients(i).Second_degree_hist_DM
            .Cells(row, 81).Value = Patients(i).family_hist_CHD
            .Cells(row, 82).Value = Patients(i).Agg_Cost
            .Cells(row, 83).Value = Patients(i).Agg_Cost_Disc
            .Cells(row, 84).Value = Patients(i).Agg_QALYs
            .Cells(row, 85).Value = Patients(i).Agg_QALYs_Disc
            .Cells(row, 86).Value = Patients(i).Age_First_Ulcer
            .Cells(row, 87).Value = Patients(i).Age_First_Amputation
            .Cells(row, 88).Value = Patients(i).Age_First_MI
            .Cells(row, 89).Value = Patients(i).Age_First_Stroke
            .Cells(row, 90).Value = Patients(i).Age_First_NASH
            .Cells(row, 91).Value = Patients(i).Age_First_OA
            .Cells(row, 92).Value = Patients(i).Age_First_OSA
            .Cells(row, 93).Value = Patients(i).Age_First_Neuro
            .Cells(row, 94).Value = Patients(i).Age_First_HTN
            .Cells(row, 95).Value = Patients(i).Age_First_PVD
            .Cells(row, 96).Value = Patients(i).Age_First_Retino
            .Cells(row, 97).Value = Patients(i).Age_First_CKD
            .Cells(row, 98).Value = Patients(i).Age_First_Dialysis
            .Cells(row, 99).Value = Patients(i).Age_First_CC
            .Cells(row, 100).Value = Patients(i).Age_First_Transplantation
 
        End With
        row = row + 1
    Next i

End Sub

Public Sub StorePatientArray(Patients() As Patient, NInterventions As Byte, NPatients As Long, Niterations As Long)
        
    Dim i As Long

    For i = 1 To NPatients
        
        With Patients(i)
        
            PastePatientArray(i + (NInterventions - 1) * NPatients, 1) = .ID
            PastePatientArray(i + (NInterventions - 1) * NPatients, 2) = .Age
            PastePatientArray(i + (NInterventions - 1) * NPatients, 3) = .BMI
            PastePatientArray(i + (NInterventions - 1) * NPatients, 4) = .BMI_Baseline
            PastePatientArray(i + (NInterventions - 1) * NPatients, 5) = .Female
            PastePatientArray(i + (NInterventions - 1) * NPatients, 6) = .Dead
            PastePatientArray(i + (NInterventions - 1) * NPatients, 7) = .Intervention_ID
            PastePatientArray(i + (NInterventions - 1) * NPatients, 8) = .smoking
            PastePatientArray(i + (NInterventions - 1) * NPatients, 9) = .Number_of_cigaretts
            PastePatientArray(i + (NInterventions - 1) * NPatients, 10) = .smoking_cess
            PastePatientArray(i + (NInterventions - 1) * NPatients, 11) = .smoking_cess_success
            PastePatientArray(i + (NInterventions - 1) * NPatients, 12) = .rehab_drinking
            PastePatientArray(i + (NInterventions - 1) * NPatients, 13) = .rehab_drinking_success
            PastePatientArray(i + (NInterventions - 1) * NPatients, 14) = .drinking
            PastePatientArray(i + (NInterventions - 1) * NPatients, 15) = .Number_of_alcohol
            PastePatientArray(i + (NInterventions - 1) * NPatients, 16) = .alcohol_categories
            PastePatientArray(i + (NInterventions - 1) * NPatients, 17) = .time_elapsed
            PastePatientArray(i + (NInterventions - 1) * NPatients, 18) = .AF
            PastePatientArray(i + (NInterventions - 1) * NPatients, 19) = .SBP
            PastePatientArray(i + (NInterventions - 1) * NPatients, 20) = .DBP
            PastePatientArray(i + (NInterventions - 1) * NPatients, 21) = .WC
            PastePatientArray(i + (NInterventions - 1) * NPatients, 22) = .occupational_risk_OA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 23) = .family_history_OA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 24) = .knee_injury
            PastePatientArray(i + (NInterventions - 1) * NPatients, 25) = .age_menopause
            PastePatientArray(i + (NInterventions - 1) * NPatients, 26) = .Menopause
            PastePatientArray(i + (NInterventions - 1) * NPatients, 27) = .Metformin_intolerance
            PastePatientArray(i + (NInterventions - 1) * NPatients, 28) = .Diabetes_Treatment_Sequence
            PastePatientArray(i + (NInterventions - 1) * NPatients, 29) = .Diabetes_treatment_ID
            PastePatientArray(i + (NInterventions - 1) * NPatients, 30) = .Insulin
            PastePatientArray(i + (NInterventions - 1) * NPatients, 31) = .TG
            PastePatientArray(i + (NInterventions - 1) * NPatients, 32) = .TC
            PastePatientArray(i + (NInterventions - 1) * NPatients, 33) = .LDL
            PastePatientArray(i + (NInterventions - 1) * NPatients, 34) = .HDL
            PastePatientArray(i + (NInterventions - 1) * NPatients, 35) = .HbA1C
            PastePatientArray(i + (NInterventions - 1) * NPatients, 36) = .FBS
            PastePatientArray(i + (NInterventions - 1) * NPatients, 37) = .Uric_Acid
            PastePatientArray(i + (NInterventions - 1) * NPatients, 38) = .ALT
            PastePatientArray(i + (NInterventions - 1) * NPatients, 39) = .AST
            PastePatientArray(i + (NInterventions - 1) * NPatients, 40) = .HCT
            PastePatientArray(i + (NInterventions - 1) * NPatients, 41) = .GGT
            PastePatientArray(i + (NInterventions - 1) * NPatients, 42) = .Scr
            PastePatientArray(i + (NInterventions - 1) * NPatients, 43) = .DM
            PastePatientArray(i + (NInterventions - 1) * NPatients, 44) = .DM_Diagnosis_Age
            PastePatientArray(i + (NInterventions - 1) * NPatients, 45) = .DM_recognized
            PastePatientArray(i + (NInterventions - 1) * NPatients, 46) = .DM_Treated
            PastePatientArray(i + (NInterventions - 1) * NPatients, 47) = .Age_First_CHD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 48) = .DM_type1
            PastePatientArray(i + (NInterventions - 1) * NPatients, 49) = .DLP
            PastePatientArray(i + (NInterventions - 1) * NPatients, 50) = .ulcer_amput_history
            PastePatientArray(i + (NInterventions - 1) * NPatients, 51) = .Ulcer
            PastePatientArray(i + (NInterventions - 1) * NPatients, 52) = .ASCVD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 53) = .Nephro
            PastePatientArray(i + (NInterventions - 1) * NPatients, 54) = .Hypertension
            PastePatientArray(i + (NInterventions - 1) * NPatients, 55) = .anti_htn_drugs
            PastePatientArray(i + (NInterventions - 1) * NPatients, 56) = .CHD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 57) = .LVH
            PastePatientArray(i + (NInterventions - 1) * NPatients, 58) = .physical_activity
            PastePatientArray(i + (NInterventions - 1) * NPatients, 59) = .Daily_fruit_consumption
            PastePatientArray(i + (NInterventions - 1) * NPatients, 60) = .previousHypoGly
            PastePatientArray(i + (NInterventions - 1) * NPatients, 61) = .Keto
            PastePatientArray(i + (NInterventions - 1) * NPatients, 62) = .PVD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 63) = .OA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 64) = .OSA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 65) = .MA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 66) = .Retino
            PastePatientArray(i + (NInterventions - 1) * NPatients, 67) = .Neuro
            PastePatientArray(i + (NInterventions - 1) * NPatients, 68) = .HypoGly
            PastePatientArray(i + (NInterventions - 1) * NPatients, 69) = .Keto_history
            PastePatientArray(i + (NInterventions - 1) * NPatients, 70) = .NASH
            PastePatientArray(i + (NInterventions - 1) * NPatients, 71) = .Stroke
            PastePatientArray(i + (NInterventions - 1) * NPatients, 72) = .Stroke_history
            PastePatientArray(i + (NInterventions - 1) * NPatients, 73) = .MI
            PastePatientArray(i + (NInterventions - 1) * NPatients, 74) = .MI_history
            PastePatientArray(i + (NInterventions - 1) * NPatients, 75) = .HF
            PastePatientArray(i + (NInterventions - 1) * NPatients, 76) = .CKD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 77) = .Dialysis
            PastePatientArray(i + (NInterventions - 1) * NPatients, 78) = .family_history_DM
            PastePatientArray(i + (NInterventions - 1) * NPatients, 79) = .First_degree_hist_DM
            PastePatientArray(i + (NInterventions - 1) * NPatients, 80) = .Second_degree_hist_DM
            PastePatientArray(i + (NInterventions - 1) * NPatients, 81) = .family_hist_CHD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 82) = .Agg_Cost
            PastePatientArray(i + (NInterventions - 1) * NPatients, 83) = .Agg_Cost_Disc
            PastePatientArray(i + (NInterventions - 1) * NPatients, 84) = .Agg_QALYs
            PastePatientArray(i + (NInterventions - 1) * NPatients, 85) = .Agg_QALYs_Disc
            PastePatientArray(i + (NInterventions - 1) * NPatients, 86) = .Age_First_Ulcer
            PastePatientArray(i + (NInterventions - 1) * NPatients, 87) = .Age_First_Amputation
            PastePatientArray(i + (NInterventions - 1) * NPatients, 88) = .Age_First_MI
            PastePatientArray(i + (NInterventions - 1) * NPatients, 89) = .Age_First_Stroke
            PastePatientArray(i + (NInterventions - 1) * NPatients, 90) = .Age_First_NASH
            PastePatientArray(i + (NInterventions - 1) * NPatients, 91) = .Age_First_OA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 92) = .Age_First_OSA
            PastePatientArray(i + (NInterventions - 1) * NPatients, 93) = .Age_First_Neuro
            PastePatientArray(i + (NInterventions - 1) * NPatients, 94) = .Age_First_HTN
            PastePatientArray(i + (NInterventions - 1) * NPatients, 95) = .Age_First_PVD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 96) = .Age_First_Retino
            PastePatientArray(i + (NInterventions - 1) * NPatients, 97) = .Age_First_CKD
            PastePatientArray(i + (NInterventions - 1) * NPatients, 98) = .Age_First_Dialysis
            PastePatientArray(i + (NInterventions - 1) * NPatients, 99) = .Age_First_CC
            PastePatientArray(i + (NInterventions - 1) * NPatients, 100) = .Age_First_Transplantation
 
        End With

    Next i

End Sub
'
'Public Sub Load_Stroke_Risk_Reference_Values()
'
'    'Builds age-group-specific SBP, DBP, and BMI reference values required by the Hunter et al. stroke equation
'
'    'Model methodology:
'    '   1. Patient_Cohort_Matrix is already loaded
'    '   2. This procedure creates a temporary VBA array with string age groups
'    '   3. FilterArray is used to filter by age group
'    '   4. ConvUnivariant is used to extract SBP, DBP, or BMI
'    '   5. Mean and sample SD are calculated once and stored in public variables
'
'    Dim StrokeReferenceArray As Variant
'
'    StrokeReferenceArray = Build_Stroke_Reference_Array
'
'    Stroke_Mean_SBP_Under50 = Stroke_Reference_Mean(StrokeReferenceArray, "<50", 2)
'    Stroke_SD_SBP_Under50 = Stroke_Reference_SD(StrokeReferenceArray, "<50", 2)
'    Stroke_Mean_DBP_Under50 = Stroke_Reference_Mean(StrokeReferenceArray, "<50", 3)
'    Stroke_SD_DBP_Under50 = Stroke_Reference_SD(StrokeReferenceArray, "<50", 3)
'    Stroke_Mean_BMI_Under50 = Stroke_Reference_Mean(StrokeReferenceArray, "<50", 4)
'    Stroke_SD_BMI_Under50 = Stroke_Reference_SD(StrokeReferenceArray, "<50", 4)
'
'    Stroke_Mean_SBP_50_59 = Stroke_Reference_Mean(StrokeReferenceArray, "50-59", 2)
'    Stroke_SD_SBP_50_59 = Stroke_Reference_SD(StrokeReferenceArray, "50-59", 2)
'    Stroke_Mean_DBP_50_59 = Stroke_Reference_Mean(StrokeReferenceArray, "50-59", 3)
'    Stroke_SD_DBP_50_59 = Stroke_Reference_SD(StrokeReferenceArray, "50-59", 3)
'    Stroke_Mean_BMI_50_59 = Stroke_Reference_Mean(StrokeReferenceArray, "50-59", 4)
'    Stroke_SD_BMI_50_59 = Stroke_Reference_SD(StrokeReferenceArray, "50-59", 4)
'
'    Stroke_Mean_SBP_60_69 = Stroke_Reference_Mean(StrokeReferenceArray, "60-69", 2)
'    Stroke_SD_SBP_60_69 = Stroke_Reference_SD(StrokeReferenceArray, "60-69", 2)
'    Stroke_Mean_DBP_60_69 = Stroke_Reference_Mean(StrokeReferenceArray, "60-69", 3)
'    Stroke_SD_DBP_60_69 = Stroke_Reference_SD(StrokeReferenceArray, "60-69", 3)
'    Stroke_Mean_BMI_60_69 = Stroke_Reference_Mean(StrokeReferenceArray, "60-69", 4)
'    Stroke_SD_BMI_60_69 = Stroke_Reference_SD(StrokeReferenceArray, "60-69", 4)
'
'    Stroke_Mean_SBP_70Plus = Stroke_Reference_Mean(StrokeReferenceArray, "70+", 2)
'    Stroke_SD_SBP_70Plus = Stroke_Reference_SD(StrokeReferenceArray, "70+", 2)
'    Stroke_Mean_DBP_70Plus = Stroke_Reference_Mean(StrokeReferenceArray, "70+", 3)
'    Stroke_SD_DBP_70Plus = Stroke_Reference_SD(StrokeReferenceArray, "70+", 3)
'    Stroke_Mean_BMI_70Plus = Stroke_Reference_Mean(StrokeReferenceArray, "70+", 4)
'    Stroke_SD_BMI_70Plus = Stroke_Reference_SD(StrokeReferenceArray, "70+", 4)
'
'
'Call Fill_Missing_Stroke_References
'
'If Stroke_SD_SBP_Under50 = 0 Or Stroke_SD_DBP_Under50 = 0 Or Stroke_SD_BMI_Under50 = 0 _
'Or Stroke_SD_SBP_50_59 = 0 Or Stroke_SD_DBP_50_59 = 0 Or Stroke_SD_BMI_50_59 = 0 _
'Or Stroke_SD_SBP_60_69 = 0 Or Stroke_SD_DBP_60_69 = 0 Or Stroke_SD_BMI_60_69 = 0 _
'Or Stroke_SD_SBP_70Plus = 0 Or Stroke_SD_DBP_70Plus = 0 Or Stroke_SD_BMI_70Plus = 0 Then
'
'Call Fill_Missing_Stroke_References
'
'End If
'
'End Sub
'Private Sub Fill_Missing_Stroke_References()
'
'    'If an age group is missing from the baseline cohort, its SD is zero
'    'Patients can still age into that group during the model
'    'To avoid division by zero in ProbStroke, missing groups borrow the nearest available age-group reference values
'
'    If Stroke_SD_SBP_60_69 = 0 Then
'        Stroke_Mean_SBP_60_69 = Stroke_Mean_SBP_50_59
'        Stroke_SD_SBP_60_69 = Stroke_SD_SBP_50_59
'        Stroke_Mean_DBP_60_69 = Stroke_Mean_DBP_50_59
'        Stroke_SD_DBP_60_69 = Stroke_SD_DBP_50_59
'        Stroke_Mean_BMI_60_69 = Stroke_Mean_BMI_50_59
'        Stroke_SD_BMI_60_69 = Stroke_SD_BMI_50_59
'    End If
'
'    If Stroke_SD_SBP_70Plus = 0 Then
'        Stroke_Mean_SBP_70Plus = Stroke_Mean_SBP_60_69
'        Stroke_SD_SBP_70Plus = Stroke_SD_SBP_60_69
'        Stroke_Mean_DBP_70Plus = Stroke_Mean_DBP_60_69
'        Stroke_SD_DBP_70Plus = Stroke_SD_DBP_60_69
'        Stroke_Mean_BMI_70Plus = Stroke_Mean_BMI_60_69
'        Stroke_SD_BMI_70Plus = Stroke_SD_BMI_60_69
'    End If
'
'End Sub
