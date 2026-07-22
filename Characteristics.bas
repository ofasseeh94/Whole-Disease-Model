Attribute VB_Name = "Characteristics"
Option Explicit

Public Previous_BMI As Double

Public Sub Characteristics_Progression(Patient As Patient)
'''''''Cycle length must be provided in years'''''''''

'some of the patient characteristics are instantly updated by other functions or modules the rest are updated here

'Go through patient characteristics and update what is required
'patient characteristics that is impossible to change are excluded from here (ex. patient ID, Gender)

'!!!!!*******VERY IMPORTANT*******************************
'This module code must be written in order of dependency (start with parameters not dependent on others then move to those dependent on them)

'Define general index variables
Dim i As Long
      
    'store last cycle age, BMI, and SBP for the delta SBP calculation later
    
        Dim OldAge As Single
        Dim OldBMI As Single
        Dim OldSBP As Single
        
        With Patient
        
            OldAge = .Age
            OldBMI = .BMI
            OldSBP = .SBP
              
            'Based on the model cycle length
            .time_elapsed = .time_elapsed + Cycle_Length 'As Double     'Duration the patient spent in the model
            .Age = .Age + Cycle_Length
            
        
'''''''''''''''''''''MENOPAUSE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '.Menopause 'As Boolean    'True= menopause occured, False=menopause didn't occur
      '.age_menopause 'As Single 'Female age at menopause
      'AGE OF MENOPAUSE
      'SOURCE for age of menopause equation: Tehrani, F. R., Solaymani-Dodaran, M., Tohidi, M., Gohari, M. R., & Azizi, F. (2013). Modeling age at menopause using serum concentration of anti-mullerian hormone. The Journal of clinical endocrinology and metabolism, 98(2), 729–735. https://doi.org/10.1210/jc.2012-3176
If .Female = True And (IsEmpty(.age_menopause) Or .age_menopause = 0) Then
    
      Dim AMH As Double
      Dim ActiveAge As Single
      
      ActiveAge = Application.WorksheetFunction.Max(20, .Age)
      'Source for log_AMH: Lee JY, Jee BC, Lee JR, Kim CH, Park T, Yeon BR, Seo SY, Lee WD, Suh CS, Kim SH. Age-related distributions of anti-Müllerian hormone level and anti-Müllerian hormone models. Acta obstetricia et gynecologica Scandinavica. 2012 Aug;91(8):970-5.
      ' in this study, AMH was ng/ml
      AMH = 10 ^ (0.205 * ActiveAge - 0.005 * (ActiveAge ^ 2) - 0.047)
      'AMH = AMH * 100
      'the equation for AMH isfrom 20-49 and
      ''age of menopause in egypt 46.7+- 5.44so we want to make it calculate if patient age is below 63(99.7% CI)
      'SOURCE FOR AGE OF MENOPAUSE IN EGYPY:https://applications.emro.who.int/emhj/0502/EMHJ_1999_5_2_307_319.pdf
        ' in this equation, AMH was ng/dl
      .age_menopause = Application.WorksheetFunction.Min(63, (-LN(0.5)) ^ (0.060388) * Exp(3.18019 + 0.1608897 * AMH + 0.016068 * (ActiveAge - .time_elapsed)))
           
End If
  If .age_menopause <= .Age Then
      
            .Menopause = True
      
      Else
      
            .Menopause = False

      End If

'''''''''''''''''''''BMI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Record previous BMI
      Previous_BMI = .BMI
      
      'Gradual change based on the surgery type
      If .time_elapsed <= UBound(ActiveIntervention.BMI_Change) Then
            'if the year is within the available data of BMI change use the provided data given the baseline
      
            .BMI = .BMI_Baseline * (1 + ActiveIntervention.BMI_Change(CInt(.time_elapsed + 0.000001)))
            'TO CHANGE THE INTERVENTION AS VARIABLE
      
      Else
            'once the duration in the model passes the available data use another model for BMI progression over time
            'Zhang, T., Whelton, P.K., Xi, B., Krousel-Wood, M., Bazzano, L.A., He, J., Chen, W. and Li, S. (2019). Rate of change in body mass index at different ages during childhood and adult obesity risk. [online] 14(7), pp.e12513–e12513. doi:https://doi.org/10.1111/ijpo.12513.
            'the paper provides data for Black and White population but race was excluded and white was used as the country is Egypt and the dominant race is white
            'The paper provides the estimated BMI at a certain age
            'We need the rate of change rather than the exact BMI so we will calculate the BMI at the current age and the previous age from the study _
             Then we will divide those values to get the % change and after that we will multiply it by the previous cycle BMI
            Dim BMI_Change As Double
            Dim Model_BMI_Previous As Single
            Dim Model_BMI_Current As Single
      
            'Source: Villareal DT, Apovian CM, Kushner RF, Klein S. Obesity in older adults: technical review and position statement of the American Society for Nutrition and NAASO, The Obesity Society. The American Journal of Clinical Nutrition [Internet]. 2005 Nov 1 [cited 2023 Dec 26];82(5):923–34. Available from: https://www.sciencedirect.com/science/article/pii/S0002916523296720#:~:text=Data%20from%20large%20population%20studies,and%20BMI%20tend%20to%20decrease.
            'Data from longitudinal cohort studies suggest that body weight and BMI do not change, or decreases only slightly, in older adults (60–70 y old at study entry) (7, 8, 9, 10)
            If .Age <= 65 Then
            
                  If .Female = True Then
                        
                        'case of female
                        Model_BMI_Current = 23.2 + 0.42 * (.Age - 20.1) - 0.08 * (.Age - 20.1) ^ 2 / 10 + 0.002 * (.Age - 20.1) ^ 3 / 20
                        Model_BMI_Previous = 23.2 + 0.42 * (.Age - Cycle_Length - 20.1) - 0.08 * (.Age - Cycle_Length - 20.1) ^ 2 / 10 + 0.002 * (.Age - Cycle_Length - 20.1) ^ 3 / 20
                        BMI_Change = Model_BMI_Current / Model_BMI_Previous
                        .BMI = .BMI * BMI_Change
                  
                  Else
                        
                        'in case of male
                        Model_BMI_Current = 23.9 + 0.48 * (.Age - 20.1) - 0.11 * (.Age - 20.1) ^ 2 / 10 + 0.002 * (.Age - 20.1) ^ 3 / 20
                        Model_BMI_Previous = 23.9 + 0.48 * (.Age - Cycle_Length - 20.1) - 0.11 * (.Age - Cycle_Length - 20.1) ^ 2 / 10 + 0.002 * (.Age - Cycle_Length - 20.1) ^ 3 / 20
                       BMI_Change = Model_BMI_Current / Model_BMI_Previous
                       .BMI = .BMI * BMI_Change
            
                  End If
            
            Else
            
                  .BMI = .BMI
            
            End If
      
      End If
      
'''''''''''''''''''''Waist Circumference''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      
      '.WC 'As Single'Waist circumfernce in cm
       
      If .BMI <= 40 Then
      'Bozeman, S.R., Hoaglin, D.C., Burton, T., Pashos, C.L., Rami Ben-Joseph and Hollenbeak, C.S. (2012). Predicting waist circumference from body mass index. [online] 12(1). doi:https://doi.org/10.1186/1471-2288-12-115.
      'ethnicity was excluded from the equation but it is available in the original model
      'Cof: Men Hispanic = -1.736731, Men Black = -3.703501, Women Hispanic = 0.1818819, Women Black = -0.6570163 ' Incase we need to add later

            If .Female = True Then
                  .WC = 28.81919 + 0.125975 * .Age + -3.688953 * IIf(.Age >= 35, 1, 0) + .BMI * 2.218007
            Else
                  .WC = 22.61306 + 0.1583812 * .Age + .BMI * 2.520738
            End If
      
      Else
      
      ' .WC = 28.81919 + 0.125975 * .Age * CInt(.Age >= 35) + -3.688953 * CInt(.Age >= 35) + .BMI * 2.218007
            If .Female = True Then
            'Source: Hassan, Nayera E.; El-Masry, Sahar A.; Elwakeel, Khaled H.; El Hussieny, Mohamed S.. Development of an easy-to-use prediction equation for waist circumference based on BMI and body weight among a sample of Egyptian women. Journal of The Arab Society for Medical Research 16(2):p 100-105, Jul–Dec 2021. | DOI: 10.4103/jasmr.jasmr_23_21
            'this equation is for females only
            'link: https://journals.lww.com/ASMR/Fulltext/2021/16020/Development_of_an_easy_to_use_prediction_equation.2.aspx
                 ' .WC = 48.44 + (1.471 * .BMI)
                 'Source: NHANES 2021-2023 equation as in the file NHANES waist circumference equation.
                 .WC = 57.3425343176783 + (.BMI * 1.54476875189187)
            Else
                  'Source: NHANES 2021-2023 equation as in the file NHANES waist circumference equation.
                 .WC = 55.6464203934405 + (.BMI * 1.8250772737945)
            
                  '.WC = 48.44 + (1.471 * .BMI)
            End If
      
      End If
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''ALT and AST'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'Source: regression conducted on NHANES database found in file "NHANES data Liver Function regression.xlsx"
 'Parameters were log transformed to the natural logarithm because the data were skewed

  
  With Patient
  
  .GGT = Exp(1.608046929 + .Age * 0.004841736 + .BMI * 0.017174985 + Abs(.Female) * -0.357500953 + .HbA1C * 0.072271899 + .TC * 0.002991914 + .HDL * -0.000764338)
  .AST = Exp(2.850595177 + .Age * 0.000876495 + .BMI * 0.000673665 + Abs(.Female) * -0.187585319 + .HbA1C * -0.006711421 + .TC * 0.000734709 + .HDL * 0.001467541)
  .ALT = Exp(2.236423928 + .Age * -0.0000232594 + .BMI * 0.013520515 + Abs(.Female) * -0.336130779 + .HbA1C * 0.020782428 + .TC * 0.002403868 + .HDL * -0.002475193)

 '5 times the Upper limit normal for ALT and 3 times for GGT
If .ALT > 280 Then .ALT = 280
If .GGT > 90 Then .GGT = 90


  End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      '.body_mass 'As Single     'Body Mass  '*******check if we are using still, if not remove. if we are using it needs to be updated in accordance to the BMI
            'REPLY: TO BE REMOVED
      '.AF 'As Boolean     'Atrial fibrillation
            
      '.TG   'As Single    'Triglycerides mg/dL
      '.TC  'As Single     'total cholesterol in mg/dL
      '.LDL   'As Single   'Low density lipoprotein mg/dL
      '.HDL   'As Single   'high density lipoprotein mg/dL
      
      '.Uric_Acid 'As Single'uric acid in blood mg/dl
      'Mean uric acid (µmol/L) in males was
      'Convert unit of uric acid from µmol/L to mg/dl
      'Scymed.com. (2023). Uric acid Unit Conversion Page :: MediCalculator ::: ScyMed ::: [online] Available at: http://www.scymed.com/en/smnxps/psxkc035_c.htm [Accessed 14 Jul. 2023].
      
      
'''''''''''''''''''''''Uric Acid''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      If .Age <= 14 Then
      'Dai, C., Wang, C., Xia, F.-Q., Liu, Z., Mo, Y., Shan, X. and Zhou, Y.-H. (2021). Age and Gender-Specific Reference Intervals for Uric Acid Level in Children Aged 5–14 Years in Southeast Zhejiang Province of China: Hyperuricemia in Children May Need Redefinition. [online] 9. doi:https://doi.org/10.3389/fped.2021.560720.
      'Source for uric acid conversion: http://www.scymed.com/en/smnxps/psxkc035_c.htm
            If .Female = True Then
                  .Uric_Acid = (177.848 + .BMI * 5.623 + .Age * 2.407)
            Else
                  .Uric_Acid = (109.321 + .BMI * 7.69 + .Age * 7.76)
            End If
      Else
      'Conen, D., Wietlisbach, V., Pascal Bovet, Shamlaye, C.F., Riesen, W.F., Paccaud, F. and Burnier, M. (2004). Prevalence of hyperuricemia and relation of serum uric acid with cardiovascular risk factors in a developing country. [online] 4(1). doi:https://doi.org/10.1186/1471-2458-4-9.
      '********Equation and source not validated yet. WE NEED TO CHECK FOR UNITS USED FOR EACH OF THE INPUT PARAMETERS AS WELL AS THE OUTCOME. mmol/L or mg/dL
'TGmmol = .TG * 0.01129

'A standard drink is equal to 14.0 grams (0.6 ounces) of pure alcohol. this is equivalent to 17.7ml. Source:https://www.cdc.gov/alcohol/faqs.htm


            If .Female = True Then
                  .Uric_Acid = (.Age * 0.62 + (.TG * 0.01129) * 46.54 + .BMI * 2.83 + CByte(.Number_of_alcohol) * 17.7 * 0.21 + Abs(.anti_htn_drugs) * 35.96 + 95.18)
            Else
                  .Uric_Acid = (.Age * 0.7 + (.TG * 0.01129) * 99.34 + .BMI * 2.08 + CByte(.Number_of_alcohol) * 17.7 * 0.16 + Abs(.anti_htn_drugs) * 45.93 + 160)
            End If
      
      .Uric_Acid = .Uric_Acid / 59.48
      
      End If
      
      
'''''''''''''''''''DBP''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'.DBP 'As Single     'Diastolic blood pressure  mmHg
'.SBP 'As Single     'Systolic blood pressure  mmHg

   Dim change_DBP As Double
   
  'ESTIMATING "CHANGE" IN DBP
  'Source: Sparrow, D., Garvey, A. J., Rosner, B., & Thomas, H. E., Jr (1982). Factors in predicting blood pressure change. Circulation, 65(4), 789–794. https://doi.org/10.1161/01.cir.65.4.789
  'this equation provides the annual change so we will adjust the change by simply dividing it by 2, this is not accurate but it is implemented for simplification
  
  change_DBP = 3.734 + 0.019 * .BMI + 0.022 * .HCT + -0.069 * .DBP
  change_DBP = change_DBP * Cycle_Length
  
  
  .DBP = .DBP + change_DBP
  
  If .DBP < 60 Then .DBP = 60
  
'************************************* SBP*******************************************
'Update SBP after BMI and age have been updated
'This calculation applies only the expected change on top of the base SBP to account for the original model situation, to account for the model conditions in treatment resistant or
'unctonrolled hypertension population

    Call BP_Update_SBP_By_Delta(Patient, OldAge, OldBMI, OldSBP)


'Since we now calculate the SBP and the DBP each seperately, we don't need to rely on the PP equation anymore as it does not make sense, we can calculate the actual PP
'by PP = SBP - DBP

Dim PP As Double

    PP = .SBP - .DBP

' 'Source:Skurnick, J. H., Aladjem, M., & Aviv, A. (2010). Sex differences in pulse pressure trends with age are cross-cultural. Hypertension (Dallas, Tex. : 1979), 55(1), 40–47. https://doi.org/10.1161/HYPERTENSIONAHA.109.139477
' Dim PP As Double
'
'  If .Female = True Then
'
'    PP = 41.9 + (.Age - 40) * 0.337 + (.Age - 40) ^ 2 * 0.0136
'
'  Else
'
'     PP = 43.4 + (.Age - 40) * 0.128 + (.Age - 40) ^ 2 * 0.0147
'
'  End If
'  'Since the SBP is the sum of DBP and PP, SBP will be always higher than DBP.
'  .SBP = PP + .DBP
  
  
'''''''''''''''''Lipid Profile''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NHANES regression models calculating difference and adding to previous value
.HDL = 72.438026272895 + Abs(.Female) * 8.62416337373032 + .Age * 0.148088531592338 + .BMI * -0.579309555164474 + .HbA1C * -2.26409805223787
'.TC = 117.91524684254 + .Age * 0.349863603149509 + .BMI * 0.383898943537148 + .HDL * 0.639981592368906
'.TG = 29.4285552800225 + .Age * 0.359719941182524 + .HbA1C * 8.10184684080546 + .HDL * -2.18061615079943 + .TC * 0.702458795238168
'.LDL = -4.00216664619377 + .Age * 7.76134075220858E-03 + .BMI * 1.96177895368991E-02 + .HDL * -0.98091869091262 + .TC * 1.00182922724858 + .TG * -0.167405984958064


  'REFERENCE: Nagy B, Zsólyom A, Nagyjanosi L, Merész G, Steiner T, Papp E, Dessewffy Z, Jermendy G, Winkler G, Kalo Z, Voko Z. Cost-effectiveness of a risk-based secondary screening programme of type 2 diabetes. Diabetes/Metabolism Research and Reviews. 2016 Oct;32(7):710-29.
If .Female = True Then
        
          .TC = Abs(181 - ((181 + -0.9726 * (.Age - Cycle_Length) + 0.0315 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .TC)) + (-0.9726 * .Age + 0.0315 * .Age ^ 2 + 0 * .Age ^ 3)
'          .HDL = Abs(42 - ((42 + -0.0435 * (.Age - Cycle_Length) + 0.0012 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .HDL)) + (-0.0435 * .Age + 0.0012 * .Age ^ 2 + 0 * .Age ^ 3)
         
  Else
          
          .TC = Abs(197 - ((197 + -0.0056 * (.Age - Cycle_Length) + 0.0052 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .TC)) + (-0.0056 * .Age + 0.0052 * .Age ^ 2 + 0 * .Age ^ 3)
'          .HDL = Abs(55 - ((55 + 0.0079 * (.Age - Cycle_Length) + 0.0004 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .HDL)) + (0.0079 * .Age + 0.0004 * .Age ^ 2 + 0 * .Age ^ 3)
  
End If
    
'Source: Takada, H., Harrell, J., Deng, S. et al. Eating habits, activity, lipids and body mass index in Japanese children: The Shiratori Children Study. Int J Obes 22, 470–476 (1998). https://doi.org/10.1038/sj.ijo.0800610
'We are using the equation from the paper to estimate the change in TG based on the change in BMI and adding it to the previous TG
.TG = .TG + (.BMI * 3.28 + Abs(.physical_activity) * -0.12 + 0 * 0.68 + Abs(.Female) * 11) - (Previous_BMI * 3.28 + Abs(.physical_activity) * -0.12 + 0 * 0.68 + Abs(.Female) * 11)
'.HDL = .HDL + (.BMI * -1.56 + Abs(.physical_activity) * 2 + 0 * -0.21 + Abs(.Female) * -4.26) - (Previous_BMI * -1.56 + Abs(.physical_activity) * 2 + 0 * -0.21 + Abs(.Female) * -4.26)
'.TC = .TC + (.BMI * 1.24 + Abs(.physical_activity) * 1.53 + 0 * -1.03 + Abs(.Female) * 0.64) - (Previous_BMI * 1.24 + Abs(.physical_activity) * 1.53 + 0 * -1.03 + Abs(.Female) * 0.64)
  
  
''''''''''''''''''''''LDL''''''''''''''''''''''''''''''''''
  
  'Source: Sampson M, Ling C, Sun Q, et al. A New Equation for Calculation of Low-Density Lipoprotein Cholesterol in Patients With Normolipidemia and/or Hypertriglyceridemia [published correction appears in JAMA Cardiol. 2020 May 1;5(5):613]. JAMA Cardiol. 2020;5(5):540-548. doi:10.1001/jamacardio.2020.0013
  ' NON-HDL= TC-HDL
  
.LDL = (.TC / 0.948) - (.HDL / 0.971) - ((.TG / 8.56) + (.TG * (.TC - .HDL) / 2140) - (.TG ^ 2) / 16100) - 9.44


''''''''''''Diabetes parameters are mostly MANAGED BY THE DIABETES MODEL
'''''''''''''''''''HbA1C'''''''''''''''''''''''''''''''''''
'HbA1C update
.HbA1C = HbA1CProg(Patient)

      '.DM_recognized 'As Boolean 'True=Patients know that they are diabetic , False= patients don't know
      If .DM = True Then
            'Source: Gopalan A, Mishra P, Alexeeff S, Blatchins MA, Kim E, Man A, et al. Prevalence and predictors of delayed clinical diagnosis of Type 2 diabetes: a longitudinal cohort study. Diabetic Medicine [Internet]. 2018 Sep 21 [cited 2023 Sep 26];35(12):1655–62. Available from: https://pubmed.ncbi.nlm.nih.gov/30175870/Print
            '30.2% remained undiagnosed with Type 2 diabetes 1 Year later
            'converted to be adjusted to 6 month cycle length then modified to diagnosed probability rather than undiagnosed
            'Diagnosis probability = 1 - undiagnosis probability
            If .DM_recognized = False Then
            
                  If RandArray(.ID, .time_elapsed / Cycle_Length, 2) < 0.450454733 Then
                        
                        .DM_recognized = True
                        '.DM_Diagnosis_Age 'As Single    'Age at diagnosis of diabetes, per year
                        .DM_Diagnosis_Age = .Age
                  
                  End If
                  
            End If
      
      End If
      
      '.DM_Treated 'As Boolean    'True = Pateint is on treatment for DM, False = patient isn't on treatment
      If .DM_recognized = True Then
      
      'in the below assumption all diagnosed patients are treated
            If .DM_Treated = False Then
            
                  If RandArray(.ID, .time_elapsed / Cycle_Length, 3) < 1 Then .DM_Treated = True
                           
            End If
            
      End If

  
 '''''''''''''''''FBS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
  
      '.FBS 'As Single     'fasting blood glucose mg/dl
      'Source: Hong, S., Kang, J.G., Kim, C.S. et al. Fasting plasma glucose concentrations for specified HbA1c goals in Korean populations: data from the Fifth Korea National Health and Nutrition Examination Survey (KNHANES V-2, 2011). Diabetol Metab Syndr 8, 62 (2016). https://doi.org/10.1186/s13098-016-0179-8
      'The equation was intended to calculate the HBA1C from FBS but we reverted the equation to propvide the FBS as the subject of formula
      
      
      .FBS = ((.HbA1C - 3.146) / 0.468) * 18 ' multiplying by 18 to change from mmol/L to mg/dL
      If .FBS < 60 Then .FBS = 60
      
      'Source: Reidpath, D.D., Jahan, N.K., Mohan, D. and Allotey, P., 2016. Single, community-based blood glucose readings may be a viable alternative for community surveillance of HbA1c and poor glycaemic control in people with known diabetes in resource-poor settings. Global health action, 9(1), p.31691.
      'dim FBSmmol/l as double
      'dim A1Cmmol as double
      'A1Cmmol= 10.929 * (.hba1c  - 2.15)
      'FBSmmol/l= (A1Cmmol-24.01)/3.99
      '.FBS= FBSmmol/l*18

  
      
'''''''''''''''''VARIABLES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
      'For the time being assume constant
      '.smoking = .smoking 'As Boolean'True = smoker, False = nonsmoker
      '.Number_of_cigaretts = .Number_of_cigaretts 'As Byte    ' Number of cigarrets per day
      '.smoking_cess = .smoking_cess 'As Boolean  'True = joined smoking cessation program, false= didn't join smoking cessation program
      '.smoking_cess_success = .smoking_cess_success 'As Boolean    'True= succeeded in quitting smoking, False= didn't succeed
      '.LVH 'As Boolean     'True= has left venticular hipertrophy, False= NO LVH   ******??????????
            
'Record disutility value based on intervention adverse events/side effects
'we check first if the duration at first is less than the decrement duration, if yes, we then check if the at end of duration(time elapsed) is still less than the decrement duration: this means the duration covers the whole cycle, then we apply the first equation
'if at end of duration (time elapsed), the decrement duration is less than the time elapsed, then we apply the part of the cycle as calculated by Decrement_Duration_Adj.
'if both conditions dont apply then no need to apply decrement duration adj.
If (.time_elapsed - Cycle_Length) * 12 < ActiveIntervention.Utility_Decrement_Duration Then
    
    If (.time_elapsed) * 12 < ActiveIntervention.Utility_Decrement_Duration Then
        
            'Capture disutility caused due to intervention it self
            ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
            Disutility_Arr(UBound(Disutility_Arr)) = Decrement_Duration_Adj(ActiveIntervention.Utility_Decrement, Cycle_Length, Cycle_Length)

    Else
    
            'Capture disutility caused due to intervention it self
            ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
            Disutility_Arr(UBound(Disutility_Arr)) = Decrement_Duration_Adj(ActiveIntervention.Utility_Decrement, ActiveIntervention.Utility_Decrement_Duration, Cycle_Length)

    End If
    
End If

'check if complication is still affecting the patient and Adjust utility value based on complications disutility accordingly. Also add costs if complication is still affecting the patient
'On Error GoTo Skip_Complications

      For i = 1 To UBound(Patient.Complication_Status)
            
      'Define when the complication should end in years
      Dim EndTime_Complication As Double
      
      EndTime_Complication = .Complication_Status(i).FirstOnset + .Complication_Status(i).Length / 52
            
            'check if complication still exists
            If EndTime_Complication <= Patient.time_elapsed - Cycle_Length Then
            
                  Patient.Complication_Status(i).Affected = False
            
            Else
                  
                  'Capture disutility caused due to intervention adverse events/side effects
                  ReDim Preserve Disutility_Arr(UBound(Disutility_Arr) + 1)
                  Disutility_Arr(UBound(Disutility_Arr)) = Decrement_Duration_Adj(.Complication_Status(i).Utility_Decrement, EndTime_Complication, Cycle_Length)
                  
                  'accumulate costs
                  .Agg_Cost = .Agg_Cost + .Complication_Status(i).Cost
                  .Agg_Cost_Disc = .Agg_Cost_Disc + DiscountedValue(CDbl(.Complication_Status(i).Cost), Disc_Costs, .time_elapsed - Cycle_Length, .time_elapsed)
                  
            End If
            
      Next i
      
Skip_Complications:
  
'Assess comorbidities
Call Assess_Comorbidities(Patient)

'HbA1C is updated after the treatment module of diabetes is run to consider the effect of the treatment in the HbA1C change
'HbA1C update
'.HbA1C = HbA1CProg(patient)
            

'************************************Doesn't change over time************************************************************************
      '.female
      '.ID
      '.DM_type1 'As Boolean' TRUE= type 1 , false= type 2
      '.Metformin_intolerance 'As Boolean    'True=Metformin intolerance occured, False=Metformin intolerance didn't occur
      
      'Family history can't change
      '.family_history_DM  'As Boolean   ' True= history of DM in family, false= no history
      '.First_degree_hist_DM  'As Boolean'True= history of first degree relatives with DM , false= no history
      '.Second_degree_hist_DM  'As Boolean     'True= history of first degree relatives with DM , false= no history
      '.family_hist_CHD  'As Boolean     'True=Family history of CHD, False= Family history of CHD
      '.family_history_OA 'As Boolean  'True=family history of OA ,False= No history
      
      'For the time being assume constant
      '.rehab_drinking = .rehab_drinking 'As Boolean'True=Rehabilitation from drinking , False= no rehabilitation
      '.rehab_drinking_success 'As Boolean'True= Rehabilitation succeeded , False= Rehabilitation failed
      '.drinking 'As Boolean     'True = Alcohol consumption, False = no Alcohol consumption
      '.Number_of_alcohol 'As Byte' number of alcohol galsses per day
      '.alcohol_categories 'As String  'drinks alcohol daily, drinks alcohol almost daily, drinks alcohol 3-4 times a week, drinks alcohol 2 times a week. drinks alcohol once a week. drinks alcohol 2-3 times a month. drinks alcohol once a month. drinks alcohol 7-11 times a year. drinks alcohol 3-6 times a year. drinks alcohol 1-2 times a year.doesn’t drink alcohol
      
      'For the time being assume constant
      '.physical_activity 'As Boolean   'true=enough physical activity, false= not enough physical activity
      '.Daily_fruit_consumption 'As Boolean   'true= enough fruit consumption, false= not enough fruit consumption
      '.occupational_risk_OA 'As String'Categorical (never, seldom, sometimes, often, always)
      '.knee_injury 'As Boolean  'True=history of knee injury ,False= No history



'***********************************DEATH***********************************************************
      'Check if the patient is dead or not
      'evaluate a random number against all death probabilities from all submodels + general mortality
      '**********VERY IMPORTANT****************
      'Combined - inverse multiplication excluding general mortality then taking the maximum between both
      'individual probabilities of death should be pooled from all submodels + the general mortality table
      If .Dead = False Then
            'aggregarte probabilities of dealth from the submodels
            'Pull the general mortality rate
            'choose the maximum mortality rate among both
            'adjust mortality rate to cycle length
            'compare the probability against a random number
            If (1 - (1 - WorksheetFunction.Max(General_Mortality(WorksheetFunction.RoundDown(.Age, 0), .Female, "OM", WorksheetFunction.RoundDown(Current_Year + .time_elapsed, 0)), Product_Inv_NotZero(Mortality_Arr))) ^ Cycle_Length) > RandArray(.ID, .time_elapsed / Cycle_Length, 44) Then .Dead = True

            'reset the global array so for the next cycle it would be empty
            Erase Mortality_Arr
      End If

      'calculate utility value based on distutilities throughout the cycle
      Utility_Instant = Agg_Utility(Disutility_Arr, Patient, Disutility_Method) * Cycle_Length
      
      'Discount and Aggregate QALYs
      .Agg_QALYs = .Agg_QALYs + Utility_Instant
      .Agg_QALYs_Disc = .Agg_QALYs_Disc + DiscountedValue(Utility_Instant, Disc_QALYs, .time_elapsed - Cycle_Length, .time_elapsed)
      
      
      'clear old disutility data at the end of the cycle
      Erase Disutility_Arr
      
End With

End Sub


Public Sub Assess_Comorbidities(Patient As Patient)

With Patient
      
      'BELOW comorbidities are updated through the comorbidities incidence Module
      
                                 ''''''''Events'''''''''
'The comorbidities below must be run in order of dependency to the biggest extent possible. _
otherwise the porbability of developing the comorbiditiy will rely on last year state of the other comorbidities.
            
      'Diabetes
            '.DM 'As Boolean'True= Has diabetes , Fasle= NO diabtes
            If .DM = True Then
                  
                  Call DM(Patient)
                        
            ElseIf (1 - (1 - ProbDiabetes(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 21) Then
            
                  Call DM(Patient)
                  .DM = True
                  If .Age_First_DM = 0 Then .Age_First_DM = .Age
                  
            End If
        
      '.Hypogly 'As Boolean  ' does the patient has hypoglycemia
      
      .HypoGly = False
            'Assess if the patient have a hypoglycemic event or not
            'HypoGly
            
            If (1 - (1 - ProbHypogly(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 5) Then
                  
                  '.HypoGly 'As Boolean    ' true=present , false= absent
                  .HypoGly = True
                  .previousHypoGly = True
                  
                  Call HypoGly(Patient)
                        
            End If

                  
       'Keto
       
      .Keto = False
            'Keto  As Boolean                      ' if patient has keto
      
                  '.Keto_history 'As Boolean   'Patient had a previous ketoacidosis or not
                        'Keto
            
            If (1 - (1 - probKeto(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 6) Then
                  
                  '.Keto 'As Boolean    ' true=present , false= absent
                  .Keto = True
            
                  Call DKA(Patient)
            
                  'Keto_history 'As Boolean  ' true= history of Keto
                  .Keto_history = True
                        
            End If

            
      
      
                                   ''''''''Continuous'''''''''
      
            '.PVD 'As Boolean     'Patient 'As peripheral vascular disease
            
            'PVD
            
            If .PVD = True Or (1 - (1 - ProbPVD(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 7) Then
                  
                  '.PVD 'As Boolean    ' true=present , false= absent
                  .PVD = True
                  
                  If .Age_First_PVD = 0 Then .Age_First_PVD = .Age

                  Call PVD(Patient)
                        
            End If

                  
                  'accumulate costs and utilities through out the period
            
                  'Add utilities and cost to the aggregator module
                        
                  'Capture probability of death
                
            
            
            '.OA 'As Boolean'Patient has osteoarthritis
                        'OA
            
            If .OA = True Or (1 - (1 - ProbOA(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 8) Then
                  
                  '.OA 'As Boolean    ' true=present , false= absent
                  .OA = True
                  
                  If .Age_First_OA = 0 Then .Age_First_OA = .Age

                  Call OA(Patient)
                        
            End If

            
            '.OSA 'As Boolean'if patient has obstructive sleep apnea
                        'OSA
            
            If .OSA = True Or (1 - (1 - ProbOSA(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 9) Then
                  
                  '.OSA 'As Boolean    ' true=present , false= absent
                  .OSA = True
                  
                  If .Age_First_OSA = 0 Then .Age_First_OSA = .Age

                  Call OSA(Patient)
                        
            End If

            
            '.MA 'As Boolean ' if patient has macular edema
                        'MA
            
            If .MA = True Or (1 - (1 - ProbMA(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 10) Then
                  
                  '.MA 'As Boolean    ' true=present , false= absent
                  .MA = True
            
                  Call MA(Patient)
                        
            End If

            
            '.Retino 'As Boolean   ' if the patient has diabetic  retinopathy
                        'Retino
            
            If .Retino = True Or (1 - (1 - ProbRetino(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 11) Then
                  
                  '.Retino 'As Boolean    ' true=present , false= absent
                  .Retino = True
                  
                  If .Age_First_Retino = 0 Then .Age_First_Retino = .Age

                  Call Retino(Patient)
                        
            End If

            
            '.Neuro 'As Boolean    ' if the patient has diabetic neuropathy
                        'Neuro
            
            If .Neuro = True Or (1 - (1 - ProbNeuro(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 12) Then
                  
                  '.Neuro 'As Boolean    ' true=present , false= absent
                  .Neuro = True
                  
                  If .Age_First_Neuro = 0 Then .Age_First_Neuro = .Age

                  Call Neuro(Patient)
                        
            End If

            
            '.DLP 'As Boolean 'true= hyperlipdemia present, false=not present
                        'DLP
            
            If .DLP = True Or (1 - (1 - ProbDLP(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 13) Then
                  
                  '.DLP 'As Boolean    ' true=present , false= absent
                  .DLP = True
            
                  Call DLP(Patient)
                        
            End If

            
            '.CHD 'As Boolean     'True= has cronary heart disease, False= No CHD
                        'CHD
            
            If .CHD = True Or (1 - (1 - ProbCHD(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 14) Then
                  
                  '.CHD 'As Boolean    ' true=present , false= absent
                  .CHD = True
            
                  Call CHD(Patient)
                  '.Age_First_CHD 'As Single  'Age 'As first
                  If .Age_First_CHD = 0 Then .Age_First_CHD = .Age
                        
            End If

            
                  
            
            '.hypertension 'As Boolean  'calculated from SBP & DBP,if SBP>=140 AND/OR DBP>=90 *********This needs further invistigation
                  'REPLY: >= 130/80 in ACC/AHA Guidelines, >=140/90 in ESC/ESH Guidelines (I recommend ESC)
                  'Source: https://www.acc.org/Latest-in-Cardiology/ten-points-to-remember/2022/08/22/16/48/Harmonization-of-the-ACC-AHA
                  
                  '.anti_htn_drugs 'As Boolean'True=On Htn drugs, False= not on HTN drugs
      
                  'HTN
            
            If .Hypertension = True Or (1 - (1 - ProbHTN(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 15) Then
                  
                  '.HTN 'As Boolean    ' true=present , false= absent
                  .Hypertension = True
                  
                  If .Age_First_HTN = 0 Then .Age_First_HTN = .Age
                  
                  Call HTN(Patient)

                        
            End If

      
                             '''''''''Pathways'''''''''
            '.NASH 'As Boolean     ' patient has NASH or not
            
                        'NASH
            
            If .NASH = True Or (1 - (1 - ProbNASH(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 16) Then
                  
                  '.NASH 'As Boolean    ' true=present , false= absent
                  .NASH = True
                  
                  If .Age_First_NASH = 0 Then .Age_First_NASH = .Age

                  Call NASH(Patient)
                        
            End If

            
            '.Nephro 'As Boolean   ' true= diabetic nephropathy present, false=not present
            
                        'Nephro
            
            'If .Nephro = True Or ProbNephro(patient) > RandArray(.ID, .time_elapsed / Cycle_Length, 17) Then
                  
                  '.Nephro 'As Boolean    ' true=present , false= absent
            '      .Nephro = True
            
            '      Call Nephro(patient)

            'End If

            
      
            'Ulcer
            
                  If .Ulcer = True Or (1 - (1 - ProbUlcer(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 18) Then
                        
                        '.Ulcer 'As Boolean    ' true=present , false= absent
                        .Ulcer = True
                  
                        Call Foot_Ulcer(Patient)
                  
                        '.ulcer_amput_history 'As Boolean  ' true= any history of ulcer or amuptation , false= no ulcer or amputation
                        If Foot_Ulcer_CHS = 4 And .ulcer_amput_history = False Then
                        
                              .ulcer_amput_history = True
                              .Age_First_Amputation = .Age
                        
                        End If
                        
                        If .Age_First_Ulcer = 0 Then .Age_First_Ulcer = .Age
                              
                  End If
            
            'CKD
            
                  If .CKD = True Or (1 - (1 - ProbCKD(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 45) Then
                        
                        '.CKD 'As Boolean    ' true=present , false= absent
                        .CKD = True
                        
                        If .Age_First_CKD = 0 Then .Age_First_CKD = .Age

                        Call CKD(Patient)
                                                
                  End If
      
      
      
      'MiniPathway
            '.Stroke 'As Boolean   ' stroke present or not
                  '.Stroke_history
            
                        'Stroke
            
            If .Stroke = True Or (1 - (1 - ProbStroke(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 19) Then
                  
                  '.Stroke 'As Boolean    ' true=present , false= absent
                  .Stroke = True
            
                  If .Age_First_Stroke = 0 Then .Age_First_Stroke = .Age
                  
                  Call Stroke(Patient)
            
                  'Stroke_history 'As Boolean  ' true= history of Stroke
                  .Stroke_history = True
                        
            End If

            'MI
            
            If .MI = True Or (1 - (1 - ProbMI(Patient)) ^ (Cycle_Length)) > RandArray(.ID, .time_elapsed / Cycle_Length, 20) Then
                  
                  '.MI 'As Boolean    ' true=present , false= absent
                  .MI = True
                  
                  If .Age_First_MI = 0 Then .Age_First_MI = .Age

                  Call MI(Patient)
            
                  'MI_history 'As Boolean  ' true= history of MI
                  .MI_history = True
                        
            End If
      
End With

End Sub

Function HbA1CProg(Patient As Patient) As Double
'this function intends to predict the progression of HbA1C of the patient over time
'the function is fed with the patient data and provides the new HbA1C after 6 month

With Patient

      If .DM = True Then
            'if patient is diabetic use this equation
                        
            'the equation provides annual steps of HbA1C so we adjust the change to the cycle length
            
            'Patients who are diabetic but not treated progress based on: Leal J, Alva M, Gregory V, Hayes A, Mihaylova B, Gray AM, Holman RR, Clarke P. Estimating risk factor progression equations for the UKPDS Outcomes Model 2 (UKPDS 90). Diabetic Medicine. 2021 Oct;38(10):e14656.
            'while diabetic treated patients progress based on: McEwan P, Bennett H, Qin L, Bergenheim K, Gordon J and Evans M. An alternative approach to modelling HbA1c trajectories in patients with type 2 diabetes mellitus. Diabetes Obes Metab. 2017;19:628– 634. https://doi.org/10.1111/dom.12865


            If .DM_Treated = False Then
            
                  HbA1CProg = .HbA1C + (((0.055 * Abs(.Female) + 0.063 * 0.5 + 0.04 * 0.5 + 0.679 * .HbA1C + 0.219 * LN(.Age + 1 - .Age_First_DM + 1) + 0.089 * BaseLine_HbA1C + 1.68) - (0.055 * Abs(.Female) + 0.063 * 0.5 + 0.04 * 0.5 + 0.679 * .HbA1C + 0.219 * LN(.Age - .Age_First_DM + 1) + 0.089 * BaseLine_HbA1C + 1.68)) * Cycle_Length)
                  'HbA1CProg = Hba1cProfileParametric(.time_elapsed - Time_Start_DM_Medication, Baseline_Before_Medication_HbA1C, treatment_effect_HbA1C, 0.980982, 9.3)
            Else
                  'The same equation for insulin and non-insulin is kept intentionally. The equation doesnt differentiate except for baseline hba1c and the treatment effect and these differe actually from patient to patient and according to the drug. both equations kept for the future if we find equations that predicts the trajectory by insulin use.
                  If .Insulin = True Then
                  
                        HbA1CProg = Hba1cProfileParametric(.time_elapsed - Time_Start_DM_Medication, Baseline_Before_Medication_HbA1C, treatment_effect_HbA1C, 0.980982, 9.3)
                  
                  Else
                  
                        HbA1CProg = Hba1cProfileParametric(.time_elapsed - Time_Start_DM_Medication, Baseline_Before_Medication_HbA1C, treatment_effect_HbA1C, 0.980982, 9.3)
                  
                  End If
                  
            End If
            
       Else
            'if patient is non-diabetic use this equation
            'Source: Rauh, S P., Heymans, M W., Koopman, A D., Nijpels, G., Stehouwer, C D. et al. (2017) Predicting glycated hemoglobin levels in the non-diabetic general population: development and validation of the DIRECT-DETECT prediction model - a DIRECT study. PLoS ONE, 12(2): e0171816 https://doi.org/10.1371/journal.pone.0171816
            Dim AgeCoff As Single, BMIcoff As Single, WCcoff As Single
            
            If .Female = True Then
            'female
                  'set age cofficient
                  If .Age >= 65 Then
                        AgeCoff = 0.307
                  ElseIf .Age >= 55 Then
                        AgeCoff = 0.213
                  ElseIf .Age >= 45 Then
                        AgeCoff = 0.173
                  Else
                        AgeCoff = 0
                  End If
            
                  'set BMI cofficient
                  If .BMI >= 30 Then
                        BMIcoff = 0.108
                  ElseIf .BMI >= 25 Then
                        BMIcoff = 0.033
                  Else
                        BMIcoff = 0
                  End If
            
                  'set WC cofficient
                  If .WC >= 88 Then
                        WCcoff = 0.106
                  ElseIf .WC >= 80 Then
                        WCcoff = 0.002
                  Else
                        WCcoff = 0
                  End If
                  
                  HbA1CProg = 5.398 + AgeCoff + BMIcoff + WCcoff + Abs(.anti_htn_drugs) * 0.05 + Abs(.smoking) * 0.093 + Abs(.family_history_DM) * 0.071
                  HbA1CProg = .HbA1C + ((HbA1CProg - .HbA1C) / 12)
            Else
            'male
                  'set age cofficient
                  If .Age >= 65 Then
                        AgeCoff = 0.188
                  ElseIf .Age >= 55 Then
                        AgeCoff = 0.091
                  ElseIf .Age >= 45 Then
                        AgeCoff = 0.053
                  Else
                        AgeCoff = 0
                  End If
            
                  'set BMI cofficient
                  If .BMI >= 30 Then
                        BMIcoff = 0.032
                  ElseIf .BMI >= 25 Then
                        BMIcoff = -0.028
                  Else
                        BMIcoff = 0
                  End If
            
                  'set WC cofficient
                  If .WC >= 102 Then
                        WCcoff = 0.115
                  ElseIf .WC >= 94 Then
                        WCcoff = 0.065
                  Else
                        WCcoff = 0
                  End If
                        
                  HbA1CProg = 5.502 + AgeCoff + BMIcoff + WCcoff + Abs(.anti_htn_drugs) * 0.05 + Abs(.smoking) * 0.133 + Abs(.family_history_DM) * 0.069
                  HbA1CProg = .HbA1C + ((HbA1CProg - .HbA1C) / 12)
            End If
      End If

End With

'Cap HbA1c to clinically feasible values to avoid outliers
If HbA1CProg > 20 Then HbA1CProg = 20

End Function

Function HbAlc_mmol(HbAlc_Perc As Single) As Single

HbAlc_mmol = (HbAlc_Perc - 2.15) * 10.929

End Function
Function HbAlc_Perc(HbAlc_mmol As Single) As Single

HbAlc_Perc = (HbAlc_mmol * 0.0915) * 2.25

End Function

Option Explicit

'===============================================================================
' Function: QRS_Estimate
'
' Purpose:
'   Estimates QRS duration in milliseconds using:
'       1. Age
'       2. BMI
'       3. Sex
'       4. Heart failure status
'       5. Obstructive sleep apnoea status
'
' Baseline QRS source:
'   Rao ACA, Ng ACC, Sy RW, et al.
'   "Electrocardiographic QRS duration is influenced by body mass index
'   and sex."
'   International Journal of Cardiology Heart & Vasculature.
'   2021;37:100884.
'
' Heart failure adjustment:
'   Approximately 25% of hospitalized heart failure patients were reported
'   to have QRS duration >=120 ms.
'
'   In this model:
'       25% of HF patients receive +40 ms
'       75% of HF patients receive +3 ms
'
'   The 25% probability is literature-based.
'   The +40 ms and +3 ms values are modelling assumptions.
'
' OSA adjustment:
'   Pressman GS, Orban M, Leinveber P, et al.
'   "Association between QRS duration and obstructive sleep apnea."
'   Journal of Clinical Sleep Medicine. 2012;8(6):649-654.
'
'   Mean QRS values reported:
'       No OSA:                85 ms
'       Mild-to-moderate OSA: 89 ms
'
'   Therefore, binary OSA receives an addition of:
'       89 - 85 = 4 ms
'
' Inputs:
'   Age:
'       Age in years.
'       Ages below 20 are treated as age 20.
'       Ages above 89 are treated as age 89.
'
'   BMI:
'       Body mass index in kg/m^2.
'
'   Female:
'       True  = female
'       False = male
'
'   HF:
'       True  = heart failure present
'       False = heart failure absent
'
'   OSA:
'       True  = obstructive sleep apnoea present
'       False = obstructive sleep apnoea absent
'
'   ID:
'       Patient identifier passed to RandArray.
'
'       RandArray(ID, 1, 1) is assumed to return a reproducible random number
'       between 0 and 1.
'
' Return:
'   Estimated QRS duration in milliseconds.
'
' Example:
'   =QRS_Estimate(65,32,TRUE,TRUE,TRUE,1001)
'
' Important limitation:
'   This is a population-level estimate and does not replace direct ECG
'   measurement. Bundle branch block, pacing, pre-excitation, electrolyte
'   disturbances and other conduction abnormalities are not explicitly modelled.
'===============================================================================

Public Function QRS_Estimate( _
    ByVal Age As Single, _
    ByVal BMI As Single, _
    ByVal Female As Boolean, _
    ByVal HF As Boolean, _
    ByVal OSA As Boolean, _
    ByVal ID As Variant) As Variant

    '===========================================================================
    ' Local fixed model parameters
    '
    ' These constants are declared inside the function so VBA can always find
    ' them, regardless of the module structure.
    '===========================================================================

    Const QRS_MINIMUM_AGE As Single = 20!
    Const QRS_MAXIMUM_AGE As Single = 89!

    Const QRS_HF_WIDE_PROBABILITY As Double = 0.25
    Const QRS_HF_WIDE_ADDITION As Double = 40#
    Const QRS_HF_NARROW_ADDITION As Double = 3#

    Const QRS_OSA_ADDITION As Double = 4#

    '===========================================================================
    ' Calculation variables
    '===========================================================================

    Dim EffectiveAge As Single
    Dim AgeCategory As Long
    Dim BMICategory As Long

    Dim BaseQRS As Double
    Dim HFAdjustment As Double
    Dim OSAAdjustment As Double
    Dim HFRandomNumber As Double

    On Error GoTo ErrorHandler

    '---------------------------------------------------------------------------
    ' Validate BMI
    '---------------------------------------------------------------------------

    If BMI <= 0 Then
        QRS_Estimate = CVErr(xlErrNA)
        Exit Function
    End If

    '---------------------------------------------------------------------------
    ' Normalize age
    '
    ' Ages below 20 are treated as age 20.
    ' Ages above 89 are treated as age 89.
    '---------------------------------------------------------------------------

    EffectiveAge = Age

    If EffectiveAge < QRS_MINIMUM_AGE Then

        EffectiveAge = QRS_MINIMUM_AGE

    ElseIf EffectiveAge > QRS_MAXIMUM_AGE Then

        EffectiveAge = QRS_MAXIMUM_AGE

    End If

    '---------------------------------------------------------------------------
    ' Assign age category
    '
    '   1 = 20-29
    '   2 = 30-39
    '   3 = 40-49
    '   4 = 50-59
    '   5 = 60-69
    '   6 = 70-79
    '   7 = 80-89
    '---------------------------------------------------------------------------

    Select Case EffectiveAge

        Case Is < 30
            AgeCategory = 1

        Case Is < 40
            AgeCategory = 2

        Case Is < 50
            AgeCategory = 3

        Case Is < 60
            AgeCategory = 4

        Case Is < 70
            AgeCategory = 5

        Case Is < 80
            AgeCategory = 6

        Case Else
            AgeCategory = 7

    End Select

    '---------------------------------------------------------------------------
    ' Assign BMI category
    '
    '   1 = BMI <18.5
    '   2 = BMI 18.5 to <25
    '   3 = BMI 25 to <30
    '   4 = BMI >=30
    '---------------------------------------------------------------------------

    Select Case BMI

        Case Is < 18.5
            BMICategory = 1

        Case Is < 25
            BMICategory = 2

        Case Is < 30
            BMICategory = 3

        Case Else
            BMICategory = 4

    End Select

    '---------------------------------------------------------------------------
    ' Retrieve baseline QRS duration according to sex, age category and BMI
    '---------------------------------------------------------------------------

    If Female Then

        '=======================================================================
        ' FEMALE BASELINE QRS DURATION, ms
        '=======================================================================

        Select Case AgeCategory

            Case 1      'Age 20-29

                Select Case BMICategory
                    Case 1: BaseQRS = 82.8
                    Case 2: BaseQRS = 83.4
                    Case 3: BaseQRS = 83.9
                    Case 4: BaseQRS = 85.6
                End Select

            Case 2      'Age 30-39

                Select Case BMICategory
                    Case 1: BaseQRS = 80.1
                    Case 2: BaseQRS = 83.1
                    Case 3: BaseQRS = 83.6
                    Case 4: BaseQRS = 85#
                End Select

            Case 3      'Age 40-49

                Select Case BMICategory
                    Case 1: BaseQRS = 80.5
                    Case 2: BaseQRS = 82.8
                    Case 3: BaseQRS = 83.5
                    Case 4: BaseQRS = 84.6
                End Select

            Case 4      'Age 50-59

                Select Case BMICategory
                    Case 1: BaseQRS = 80.3
                    Case 2: BaseQRS = 82.6
                    Case 3: BaseQRS = 83.2
                    Case 4: BaseQRS = 84.5
                End Select

            Case 5      'Age 60-69

                Select Case BMICategory
                    Case 1: BaseQRS = 80.8
                    Case 2: BaseQRS = 82.8
                    Case 3: BaseQRS = 83.5
                    Case 4: BaseQRS = 84.2
                End Select

            Case 6      'Age 70-79

                Select Case BMICategory
                    Case 1: BaseQRS = 81.2
                    Case 2: BaseQRS = 82.6
                    Case 3: BaseQRS = 82.9
                    Case 4: BaseQRS = 84#
                End Select

            Case 7      'Age 80-89

                Select Case BMICategory
                    Case 1: BaseQRS = 80.1
                    Case 2: BaseQRS = 82.4
                    Case 3: BaseQRS = 82.7
                    Case 4: BaseQRS = 89.9
                End Select

        End Select

    Else

        '=======================================================================
        ' MALE BASELINE QRS DURATION, ms
        '=======================================================================

        Select Case AgeCategory

            Case 1      'Age 20-29

                Select Case BMICategory
                    Case 1: BaseQRS = 91.3
                    Case 2: BaseQRS = 94.4
                    Case 3: BaseQRS = 94#
                    Case 4: BaseQRS = 94.3
                End Select

            Case 2      'Age 30-39

                Select Case BMICategory
                    Case 1: BaseQRS = 87.7
                    Case 2: BaseQRS = 92.3
                    Case 3: BaseQRS = 93.2
                    Case 4: BaseQRS = 93.1
                End Select

            Case 3      'Age 40-49

                Select Case BMICategory
                    Case 1: BaseQRS = 90.5
                    Case 2: BaseQRS = 91.8
                    Case 3: BaseQRS = 92.4
                    Case 4: BaseQRS = 92.9
                End Select

            Case 4      'Age 50-59

                Select Case BMICategory
                    Case 1: BaseQRS = 86.5
                    Case 2: BaseQRS = 90.2
                    Case 3: BaseQRS = 91.5
                    Case 4: BaseQRS = 91.8
                End Select

            Case 5      'Age 60-69

                Select Case BMICategory
                    Case 1: BaseQRS = 83#
                    Case 2: BaseQRS = 89.5
                    Case 3: BaseQRS = 90.7
                    Case 4: BaseQRS = 91.4
                End Select

            Case 6      'Age 70-79

                Select Case BMICategory
                    Case 1: BaseQRS = 86.1
                    Case 2: BaseQRS = 89.4
                    Case 3: BaseQRS = 90.5
                    Case 4: BaseQRS = 90.9
                End Select

            Case 7      'Age 80-89

                Select Case BMICategory
                    Case 1: BaseQRS = 88.5
                    Case 2: BaseQRS = 89.5
                    Case 3: BaseQRS = 89.6
                    Case 4: BaseQRS = 89.9
                End Select

        End Select

    End If

    '---------------------------------------------------------------------------
    ' Apply heart failure adjustment
    '
    ' RandArray is called only once so the patient remains in the same HF
    ' subgroup during the current function call.
    '---------------------------------------------------------------------------

    HFAdjustment = 0#

    If HF Then

        HFRandomNumber = CDbl(RandArray(ID, 1, 1))

        'Ensure RandArray returned a valid probability.
        If HFRandomNumber < 0# Or HFRandomNumber > 1# Then
            QRS_Estimate = CVErr(xlErrNum)
            Exit Function
        End If

        If HFRandomNumber <= QRS_HF_WIDE_PROBABILITY Then

            '25% of HF patients
            HFAdjustment = QRS_HF_WIDE_ADDITION

        Else

            'Remaining 75% of HF patients
            HFAdjustment = QRS_HF_NARROW_ADDITION

        End If

    End If

    '---------------------------------------------------------------------------
    ' Apply obstructive sleep apnoea adjustment
    '---------------------------------------------------------------------------

    OSAAdjustment = 0#

    If OSA Then
        OSAAdjustment = QRS_OSA_ADDITION
    End If

    '---------------------------------------------------------------------------
    ' Return final estimated QRS duration
    '---------------------------------------------------------------------------

    QRS_Estimate = BaseQRS + HFAdjustment + OSAAdjustment

    Exit Function

ErrorHandler:

    'Return Excel #N/A if an unexpected error occurs.
    QRS_Estimate = CVErr(xlErrNA)

End Function
