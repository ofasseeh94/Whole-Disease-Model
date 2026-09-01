Attribute VB_Name = "Characteristics"
Option Explicit


'Store the cycle-start age, BMI, and SBP, and DBP as module-level public variables.
Public OldAge As Single
Public OldBMI As Single
Public OldSBP As Single
Public OldDBP As Single
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

With Patient
      
'The SBP, and DBP progression equations use old and new age/BMI to calculate only the expected absolute SBP, and DBP change for this cycle. The actual patient.SBP, and
'.DBP is kept in this Characteristics_Progression section, so patient characteristics are altered in one structured location

OldAge = .Age
OldBMI = .BMI
OldSBP = .SBP
OldDBP = .DBP
      
      
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
'Not to exceed
If .BMI > 69.1 Then .BMI = 69.1
'Not to be less than
If .BMI < 15 Then .BMI = 15
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
'Not to exceed
If .WC > 200 Then .WC = 200
'Not to be less than
If .WC < 50 Then .WC = 50
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''ALT and AST'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'Source: regression conducted on NHANES database found in file "NHANES data Liver Function regression.xlsx"
 'Parameters were log transformed to the natural logarithm because the data were skewed

  
  With Patient
  
'  .GGT = Exp(1.608046929 + .Age * 0.004841736 + .BMI * 0.017174985 + Abs(.Female) * -0.357500953 + .HbA1C * 0.072271899 + .TC * 0.002991914 + .HDL * -0.000764338)
'  .AST = Exp(2.850595177 + .Age * 0.000876495 + .BMI * 0.000673665 + Abs(.Female) * -0.187585319 + .HbA1C * -0.006711421 + .TC * 0.000734709 + .HDL * 0.001467541)
'  .ALT = Exp(2.236423928 + .Age * -0.0000232594 + .BMI * 0.013520515 + Abs(.Female) * -0.336130779 + .HbA1C * 0.020782428 + .TC * 0.002403868 + .HDL * -0.002475193)
Dim logalt As Double
Dim gender_num As Integer
If .Female = False Then
    gender_num = 1
Else
    gender_num = 0
End If

logalt = 0.869729371643 _
    - 0.857315566176 * gender_num _
    + 0.019072702259 * .Age _
    - 0.004746368871 * gender_num * .Age _
    - 0.000175422566 * .Age * .Age _
    - 0.000028589762 * gender_num * .Age * .Age _
    + 0.027262219512 * .BMI _
    + 0.062220621608 * gender_num * .BMI _
    - 0.000289523 * .BMI * .BMI _
    - 0.000809878162 * gender_num * .BMI * .BMI _
    + 0.001266559103 * .HDL _
    + 0.005778417906 * gender_num * .HDL _
    - 0.000269235297 * .TC _
    - 0.000129487219 * gender_num * .TC _
    + 0.03976057495 * .HbA1c _
    - 0.037293983897 * gender_num * .HbA1c _
    + 0.145572713216 * Log(.TG) _
    + 0.076272540433 * gender_num * Log(.TG)

.ALT = 1.1291495142703 * Exp(logalt)
    
Dim logast As Double


logast = 2.663579890036 _
    + 0.078571809803 * gender_num _
    + 0.003609920484 * .Age _
    - 0.00449458095 * gender_num * .Age _
    + 0.000121908841 * .BMI _
    + 0.004514008919 * gender_num * .BMI _
    + 0.001141671184 * .HDL _
    + 0.003032110804 * gender_num * .HDL _
    + 0.000082453091 * .TC _
    + 0.000452513699 * gender_num * .TC _
    + 0.005526204277 * .HbA1c _
    - 0.010010722564 * gender_num * .HbA1c

.AST = 1.0801390328268 * Exp(logast)

Dim logggt As Double
Dim lntg As Double

lntg = Log(.TG)

logggt = -1.037754095807 _
    - 0.666877484897 * gender_num _
    + 0.021650880194 * .Age _
    - 0.001729998953 * gender_num * .Age _
    - 0.000187304557 * .Age * .Age _
    - 0.000035523903 * gender_num * .Age * .Age _
    + 0.040560163755 * .BMI _
    + 0.043847919229 * gender_num * .BMI _
    - 0.00038342176 * .BMI * .BMI _
    - 0.000617686274 * gender_num * .BMI * .BMI _
    + 0.004437880777 * .HDL _
    + 0.010076180848 * gender_num * .HDL _
    - 0.001041803282 * .TC _
    + 0.001545055732 * gender_num * .TC _
    + 0.126042883901 * .HbA1c _
    - 0.070027400341 * gender_num * .HbA1c _
    + 0.387868620771 * lntg _
    + 0.013554995466 * gender_num * lntg

.GGT = 1.2669564846961 * Exp(logggt)

 '5 times the Upper limit normal for ALT and 3 times for GGT
'Not to exceed
If .GGT > 90 Then .GGT = 90
'Not to be less than
If .GGT < 5 Then .GGT = 5

'Not to exceed
If .AST > 280 Then .AST = 280
'Not to be less than
If .AST < 5 Then .AST = 5

'Not to exceed
If .ALT > 280 Then .ALT = 280
'Not to be less than
If .ALT < 5 Then .ALT = 5

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
'Not to exceed
If .Uric_Acid > 10.3 Then .Uric_Acid = 10.3
'Not to be less than
If .Uric_Acid < 3.98 Then .Uric_Acid = 3.98
      

'************************************* SBP*******************************************
'Update SBP after BMI and age have been updated
'This calculation applies only the expected change on top of the base SBP to account for the original model situation, to account for the model conditions in treatment resistant or
'unctonrolled hypertension population

'Dim SBP_Absolute_Change As Single
'
'SBP_Absolute_Change = BP_SBP_Absolute_Change(Patient, OldAge, OldBMI)
'
'    'Apply only the model-predicted absolute SBP change on top of the patient's actual cycle-start SBP. This preserves baseline resistant/uncontrolled hypertension
'    'burden and prevents the equation from replacing actual patient SBP with anaverage predicted value.
'    If OldSBP > 0 Then
'        .SBP = OldSBP + SBP_Absolute_Change
'    End If
''Important:
''Table 2 reports coefficients for longitudinal tracking/change, not a full absolute PP prediction
''equation with an intercept. Therefore, we calculate the change from old patient state to current
''patient state, then apply that change to the patient's actual cycle-start PP.
'
'Dim OldPP As Double
'Dim OldPP_Cheng_Component As Double
'Dim NewPP_Cheng_Component As Double
'Dim PP_Absolute_Change As Double
'Dim PP As Double
'
'    OldPP = OldSBP - OldDBP
'
'    OldPP_Cheng_Component = _
'          4.656 * ((OldAge - 49) / 10) _
'        + 6.241 * Abs(Abs(.Female) - 1) _
'        - 1.31 * ((OldAge - 49) / 10) * Abs(Abs(.Female) - 1) _
'        + 1.739 * (OldBMI / 5) _
'        - 0.993 * Abs(Abs(.Female) - 1) * (OldBMI / 5) _
'        + 0.557 * Abs(.smoking) _
'        + 4.571 * Abs(.DM) _
'        - 0.719 * ((.TC / .HDL) / 2)
'
'    NewPP_Cheng_Component = _
'          4.656 * ((.Age - 49) / 10) _
'        + 6.241 * Abs(Abs(.Female) - 1) _
'        - 1.31 * ((.Age - 49) / 10) * Abs(Abs(.Female) - 1) _
'        + 1.739 * (.BMI / 5) _
'        - 0.993 * Abs(Abs(.Female) - 1) * (.BMI / 5) _
'        + 0.557 * Abs(.smoking) _
'        + 4.571 * Abs(.DM) _
'        - 0.719 * ((.TC / .HDL) / 2)
'
'    PP_Absolute_Change = NewPP_Cheng_Component - OldPP_Cheng_Component
'
'    PP = OldPP + PP_Absolute_Change
'
'    'Derive DBP from the updated SBP and estimated current PP.
'    .DBP = .SBP - PP
      
   
   
''HCT''
Dim Change_HCT As Double
Dim Male As Single

If .Female Then

    Male = 0
    
    Else
    
    Male = 1

End If

.HCT = 42.2435294 _
        - 0.293257202 * .Age _
        + 0.00788012742 * .Age ^ 2 _
        - 0.0000581892533 * .Age ^ 3 _
        + 7.00322478 * Male _
        + 0.143003991 * .BMI _
        - 0.00189879669 * .BMI ^ 2 _
        - 0.852816343 * .HbA1c _
        + 0.0545071927 * .HbA1c ^ 2 _
        + 0.910794546 * .smoking _
        - 0.0562697049 * Male * .Age
   
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

'Not to exceed
If .DBP > 120 Then .DBP = 120
'Not to be less than
If .DBP < 60 Then .DBP = 60
'
''************************************* SBP*******************************************
 'Source:Skurnick, J. H., Aladjem, M., & Aviv, A. (2010). Sex differences in pulse pressure trends with age are cross-cultural. Hypertension (Dallas, Tex. : 1979), 55(1), 40–47. https://doi.org/10.1161/HYPERTENSIONAHA.109.139477
 Dim PP As Double

  If .Female = True Then

    PP = 41.9 + (.Age - 40) * 0.337 + (.Age - 40) ^ 2 * 0.0136

  Else

     PP = 43.4 + (.Age - 40) * 0.128 + (.Age - 40) ^ 2 * 0.0147

  End If
  'Since the SBP is the sum of DBP and PP, SBP will be always higher than DBP.
  .SBP = PP + .DBP
If .SBP > 250 Then .SBP = 250
'Not to be less than
If .SBP < 50 Then .SBP = 50
  
'''''''''''''''''Lipid Profile''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NHANES regression models calculating difference and adding to previous value
.HDL = 72.438026272895 + Abs(.Female) * 8.62416337373032 + .Age * 0.148088531592338 + .BMI * -0.579309555164474 + .HbA1c * -2.26409805223787
'.TC = 117.91524684254 + .Age * 0.349863603149509 + .BMI * 0.383898943537148 + .HDL * 0.639981592368906
'.TG = 29.4285552800225 + .Age * 0.359719941182524 + .HbA1C * 8.10184684080546 + .HDL * -2.18061615079943 + .TC * 0.702458795238168
'.LDL = -4.00216664619377 + .Age * 7.76134075220858E-03 + .BMI * 1.96177895368991E-02 + .HDL * -0.98091869091262 + .TC * 1.00182922724858 + .TG * -0.167405984958064
'Not to exceed
If .HDL > 100 Then .HDL = 100
'Not to be less than
If .HDL < 10 Then .HDL = 10

  'REFERENCE: Nagy B, Zsólyom A, Nagyjanosi L, Merész G, Steiner T, Papp E, Dessewffy Z, Jermendy G, Winkler G, Kalo Z, Voko Z. Cost-effectiveness of a risk-based secondary screening programme of type 2 diabetes. Diabetes/Metabolism Research and Reviews. 2016 Oct;32(7):710-29.
If .Female = True Then
        
          .TC = Abs(181 - ((181 + -0.9726 * (.Age - Cycle_Length) + 0.0315 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .TC)) + (-0.9726 * .Age + 0.0315 * .Age ^ 2 + 0 * .Age ^ 3)
'          .HDL = Abs(42 - ((42 + -0.0435 * (.Age - Cycle_Length) + 0.0012 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .HDL)) + (-0.0435 * .Age + 0.0012 * .Age ^ 2 + 0 * .Age ^ 3)
         
  Else
          
          .TC = Abs(197 - ((197 + -0.0056 * (.Age - Cycle_Length) + 0.0052 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .TC)) + (-0.0056 * .Age + 0.0052 * .Age ^ 2 + 0 * .Age ^ 3)
'          .HDL = Abs(55 - ((55 + 0.0079 * (.Age - Cycle_Length) + 0.0004 * (.Age - Cycle_Length) ^ 2 + 0 * (.Age - Cycle_Length) ^ 3) - .HDL)) + (0.0079 * .Age + 0.0004 * .Age ^ 2 + 0 * .Age ^ 3)

End If
'Not to exceed
 If .TC > 248 Then .TC = 248
'Not to be less than
If .TC < 98.2 Then .TC = 98.2

'Source: Takada, H., Harrell, J., Deng, S. et al. Eating habits, activity, lipids and body mass index in Japanese children: The Shiratori Children Study. Int J Obes 22, 470–476 (1998). https://doi.org/10.1038/sj.ijo.0800610
'We are using the equation from the paper to estimate the change in TG based on the change in BMI and adding it to the previous TG
.TG = .TG + (.BMI * 3.28 + Abs(.physical_activity) * -0.12 + 0 * 0.68 + Abs(.Female) * 11) - (Previous_BMI * 3.28 + Abs(.physical_activity) * -0.12 + 0 * 0.68 + Abs(.Female) * 11)
'.HDL = .HDL + (.BMI * -1.56 + Abs(.physical_activity) * 2 + 0 * -0.21 + Abs(.Female) * -4.26) - (Previous_BMI * -1.56 + Abs(.physical_activity) * 2 + 0 * -0.21 + Abs(.Female) * -4.26)
'.TC = .TC + (.BMI * 1.24 + Abs(.physical_activity) * 1.53 + 0 * -1.03 + Abs(.Female) * 0.64) - (Previous_BMI * 1.24 + Abs(.physical_activity) * 1.53 + 0 * -1.03 + Abs(.Female) * 0.64)
'Not to exceed
If .TG > 344.5 Then .TG = 344.5
'Not to be less than
If .TG < 60 Then .TG = 60
  
''''''''''''''''''''''LDL''''''''''''''''''''''''''''''''''
  
  'Source: Sampson M, Ling C, Sun Q, et al. A New Equation for Calculation of Low-Density Lipoprotein Cholesterol in Patients With Normolipidemia and/or Hypertriglyceridemia [published correction appears in JAMA Cardiol. 2020 May 1;5(5):613]. JAMA Cardiol. 2020;5(5):540-548. doi:10.1001/jamacardio.2020.0013
  ' NON-HDL= TC-HDL
  
.LDL = (.TC / 0.948) - (.HDL / 0.971) - ((.TG / 8.56) + (.TG * (.TC - .HDL) / 2140) - (.TG ^ 2) / 16100) - 9.44
'Not to exceed
If .LDL > 700 Then .LDL = 700
'Not to be less than
If .LDL < 10 Then .LDL = 10

''''''''''''Diabetes parameters are mostly MANAGED BY THE DIABETES MODEL
'''''''''''''''''''HbA1C'''''''''''''''''''''''''''''''''''
'HbA1C update
.HbA1c = HbA1CProg(Patient)

      '.DM_recognized 'As Boolean 'True=Patients know that they are diabetic , False= patients don't know
      If .DM = True Then
            If .DM_recognized = False Then
            
                  If RandArray(.ID, .time_elapsed / Cycle_Length, 2) < (1 - (1 - AnnualDiabetesDiagnosisProbability(.HbA1c, .BMI, .Female)) ^ Cycle_Length) Then
                        
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
            
                  'If RandArray(.ID, .time_elapsed / Cycle_Length, 3) < 1 Then .DM_Treated = True
                  If RandArray(.ID, .time_elapsed / Cycle_Length, 3) < (1 - (1 - AnnualDiabetesTreatmentProbability(.Age, .HbA1c, .Female)) ^ Cycle_Length) Then .DM_Treated = True
                           
            End If
            
      End If

  
 '''''''''''''''''FBS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
  
      '.FBS 'As Single     'fasting blood glucose mg/dl
      'Source: Hong, S., Kang, J.G., Kim, C.S. et al. Fasting plasma glucose concentrations for specified HbA1c goals in Korean populations: data from the Fifth Korea National Health and Nutrition Examination Survey (KNHANES V-2, 2011). Diabetol Metab Syndr 8, 62 (2016). https://doi.org/10.1186/s13098-016-0179-8
      'The equation was intended to calculate the HBA1C from FBS but we reverted the equation to propvide the FBS as the subject of formula
      
      
      .FBS = ((.HbA1c - 3.146) / 0.468) * 18 ' multiplying by 18 to change from mmol/L to mg/dL
      If .FBS < 60 Then .FBS = 60
      
      'Source: Reidpath, D.D., Jahan, N.K., Mohan, D. and Allotey, P., 2016. Single, community-based blood glucose readings may be a viable alternative for community surveillance of HbA1c and poor glycaemic control in people with known diabetes in resource-poor settings. Global health action, 9(1), p.31691.
      'dim FBSmmol/l as double
      'dim A1Cmmol as double
      'A1Cmmol= 10.929 * (.hba1c  - 2.15)
      'FBSmmol/l= (A1Cmmol-24.01)/3.99
      '.FBS= FBSmmol/l*18
'Not to exceed
If .FBS > 1500 Then .FBS = 1500
'Not to be less than
If .FBS < 30 Then .FBS = 30
  
      
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
            
                  HbA1CProg = .HbA1c + (((0.055 * Abs(.Female) + 0.063 * 0.5 + 0.04 * 0.5 + 0.679 * .HbA1c + 0.219 * LN(.Age + 1 - .Age_First_DM + 1) + 0.089 * BaseLine_HbA1C + 1.68) - (0.055 * Abs(.Female) + 0.063 * 0.5 + 0.04 * 0.5 + 0.679 * .HbA1c + 0.219 * LN(.Age - .Age_First_DM + 1) + 0.089 * BaseLine_HbA1C + 1.68)) * Cycle_Length)
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
                  HbA1CProg = .HbA1c + ((HbA1CProg - .HbA1c) / 12)
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
                  HbA1CProg = .HbA1c + ((HbA1CProg - .HbA1c) / 12)
            End If
      End If

End With

'Cap HbA1c to clinically feasible values to avoid outliers
If HbA1CProg > 20 Then HbA1CProg = 20
If HbA1CProg < 5 Then HbA1CProg = 5

End Function

Function HbAlc_mmol(HbAlc_Perc As Single) As Single

HbAlc_mmol = (HbAlc_Perc - 2.15) * 10.929

End Function
Function HbAlc_Perc(HbAlc_mmol As Single) As Single

HbAlc_Perc = (HbAlc_mmol * 0.0915) * 2.25

End Function

Function AnnualDiabetesDiagnosisProbability(HbA1c As Single, BMI As Single, Female As Boolean) As Double

'===============================================================================
' FUNCTION:
' AnnualDiabetesDiagnosisProbability
'
' PURPOSE:
' Estimates the annual probability that a patient with previously undiagnosed
' diabetes will receive a clinical diabetes diagnosis during the current year.
'
' MODEL LOGIC:
' 1. HbA1c enters the disease model in NGSP/DCCT units (%).
'
' 2. HbA1c is converted internally from % to IFCC units (mmol/mol), because the
'    Cox regression used to parameterize diagnosis rates reported HbA1c in
'    mmol/mol categories.
'
'       IFCC HbA1c (mmol/mol) = (HbA1c [%] - 2.152) / 0.09148
'
'    This is equivalent to approximately:
'
'       IFCC HbA1c = (HbA1c [%] - 2.15) * 10.93
'
' 3. A hazard-ratio multiplier is constructed from patient characteristics:
'
'       Male sex:               HR = 1.12
'       Female sex:             HR = 1.00 [reference]
'
'       BMI <30 kg/m2:          HR = 1.00 [reference]
'       BMI >=30 kg/m2:         HR = 1.25
'
'       HbA1c 48.0-52.9 mmol/mol: HR = 1.00 [reference]
'       HbA1c 53.0-57.9 mmol/mol: HR = 2.13
'       HbA1c >=58.0 mmol/mol:     HR = 2.71
'
' 4. Cox proportional-hazards logic is used. The individual hazard is:
'
'       Individual Hazard = Baseline Hazard * Combined HR
'
'    where:
'
'       Combined HR =
'           HR(sex) * HR(BMI) * HR(HbA1c)
'
' 5. The hazard is converted into a 1-year probability:
'
'       P(Diagnosis during year) = 1 - Exp(-Individual Hazard)
'
' IMPORTANT:
' The published UK Biobank Cox model reported hazard ratios but did not provide
' a directly implementable absolute baseline hazard for the reference patient.
'
' Therefore, the baseline hazard below is CALIBRATED rather than taken directly
' from the published Cox coefficients. A baseline annual diagnosis probability
' of 25.5% is currently used:
'
'       BaselineHazard = -Ln(1 - 0.255)
'
' This value was derived from the approximately 23% remaining undiagnosed at
' 5 years reported in the UK Biobank analysis, assuming approximately constant
' annual hazard:
'
'       Annual survival = 0.23^(1/5) ~= 0.745
'
'       Annual diagnosis probability ~= 1 - 0.745 = 0.255
'
' Hence, the absolute probabilities generated by this function should be
' considered a calibrated approximation based on the published Cox relative
' effects, rather than an exact reproduction of the original Cox survival model.
'
'
' PRIMARY CLINICAL REFERENCE:
'
' Jones A, et al.
' "The impact of population-level HbA1c screening on reducing diabetes
' diagnostic delay in middle-aged adults: a UK Biobank analysis."
' Diabetologia. 2023.
'
' The study identified individuals with biochemical evidence of diabetes who
' did not have a previous clinical diagnosis and followed their subsequent
' diagnosis in routine care.
'
' Key findings used in this function:
'   - Male vs female: HR 1.12 (95% CI 1.00-1.25)
'   - BMI >=30 vs <30 kg/m2: HR 1.25 (95% CI 1.12-1.39)
'   - HbA1c 53.0-57.9 vs 48.0-52.9 mmol/mol:
'         HR 2.13 (95% CI 1.84-2.46)
'   - HbA1c >=58.0 vs 48.0-52.9 mmol/mol:
'         HR 2.71 (95% CI 2.37-3.09)
'
' Full text:
' https://pmc.ncbi.nlm.nih.gov/articles/PMC9807472/
'
'
' HbA1c CONVERSION REFERENCE:
'
' National Glycohemoglobin Standardization Program (NGSP) / International
' Federation of Clinical Chemistry (IFCC).
'
' Master equation:
'
'       NGSP (%) = 0.09148 * IFCC (mmol/mol) + 2.152
'
' Therefore:
'
'       IFCC (mmol/mol) = (NGSP [%] - 2.152) / 0.09148
'
' NGSP reference:
' https://ngsp.org/ifcc.asp
'
'
' INPUTS:
'
' HbA1cPercent
'   Patient HbA1c expressed in NGSP/DCCT percentage units.
'   Example: 7.2 means HbA1c = 7.2%.
'
' BMI
'   Patient body mass index in kg/m2.
'
' IsFemale
'   True  = female
'   False = male
'
'
' OUTPUT:
'
' Annual probability between 0 and 1 that an UNDIGNOSED diabetic patient
' becomes clinically diagnosed during the current annual model cycle.
'
' Example:
'       0.30 = 30% probability of diagnosis during the year.
'
'
' IMPORTANT APPLICATION:
'
' This function should ONLY be applied to patients who:
'   1. currently have diabetes, AND
'   2. have not yet been clinically diagnosed.
'
' Once a patient becomes diagnosed, this probability should no longer be
' evaluated for that patient.
'
'===============================================================================

    Dim HbA1cMMol As Double
    Dim HR As Double

    '---------------------------------------------------------------------------
    ' STEP 1: Convert HbA1c from NGSP percentage to IFCC mmol/mol.
    '
    ' Official NGSP/IFCC equation:
    ' NGSP = 0.09148 * IFCC + 2.152
    ' Therefore:
    ' IFCC = (NGSP - 2.152) / 0.09148
    '---------------------------------------------------------------------------

    HbA1cMMol = (HbA1c - 2.152) / 0.09148

    '---------------------------------------------------------------------------
    ' STEP 2: Initialise Cox hazard ratio.
    '
    ' Reference patient:
    '   Female
    '   BMI <30 kg/m2
    '   HbA1c 48.0-52.9 mmol/mol
    '---------------------------------------------------------------------------

    HR = 1

    '---------------------------------------------------------------------------
    ' STEP 3: Sex effect.
    '
    ' Jones et al.:
    ' Male vs female HR = 1.12.
    '
    ' Since IsFemale = True represents the reference group,
    ' the HR is applied only when IsFemale = False.
    '---------------------------------------------------------------------------

    If Not Female Then HR = HR * 1.12

    '---------------------------------------------------------------------------
    ' STEP 4: BMI effect.
    '
    ' Jones et al.:
    ' BMI >=30 kg/m2 vs BMI <30 kg/m2:
    ' HR = 1.25
    '---------------------------------------------------------------------------

    If BMI >= 30 Then HR = HR * 1.25

    '---------------------------------------------------------------------------
    ' STEP 5: HbA1c effect.
    '
    ' Jones et al. categories:
    '
    ' 48.0-52.9 mmol/mol = reference HR 1.00
    ' 53.0-57.9 mmol/mol = HR 2.13
    ' >=58.0 mmol/mol     = HR 2.71
    '---------------------------------------------------------------------------

    If HbA1cMMol >= 58 Then

        HR = HR * 2.71

    ElseIf HbA1cMMol >= 53 Then

        HR = HR * 2.13

    End If

    '---------------------------------------------------------------------------
    ' STEP 6: Baseline annual hazard.
    '
    ' Calibrated using an annual diagnosis probability of approximately 25.5%.
    '
    ' For a constant hazard:
    '
    ' Probability = 1 - Exp(-Hazard)
    '
    ' Therefore:
    '
    ' Hazard = -Ln(1 - Probability)
    '---------------------------------------------------------------------------

    AnnualDiabetesDiagnosisProbability = -Log(1 - 0.255)

    '---------------------------------------------------------------------------
    ' STEP 7: Apply Cox proportional hazard multiplier.
    '---------------------------------------------------------------------------

    AnnualDiabetesDiagnosisProbability = AnnualDiabetesDiagnosisProbability * HR


    '---------------------------------------------------------------------------
    ' STEP 8: Convert annual hazard into annual probability.
    '
    ' P = 1 - Exp(-hazard)
    '---------------------------------------------------------------------------

    AnnualDiabetesDiagnosisProbability = 1 - Exp(-AnnualDiabetesDiagnosisProbability)
    'Return annual probability of receiving a diabetes diagnosis.

End Function

'===============================================================================
' FUNCTION:
' AnnualDiabetesTreatmentProbability
'
' PURPOSE:
' Estimates the annual probability that a patient with diagnosed type 2
' diabetes who is currently NOT receiving pharmacological glucose-lowering
' therapy will initiate antihyperglycaemic medication during the current year.
'
' MODEL TRANSITION:
'
'   Diagnosed / recognized diabetes + not treated
'
'                         |
'                         | AnnualDiabetesTreatmentProbability()
'                         v
'
'   Diagnosed / recognized diabetes + pharmacologically treated
'
'
' IMPORTANT:
' "Treatment" in this function refers to initiation of pharmacological
' antihyperglycaemic medication.
'
'
' SOURCE:
'
' Sinclair AJ, Alexander CM, Davies MJ, Zhao C, Mavros P.
' Factors associated with initiation of antihyperglycaemic medication in UK
' patients with newly diagnosed type 2 diabetes.
' BMC Endocrine Disorders. 2012;12:1.
'
' DOI:
' 10.1186/1472-6823-12-1
'
' Study characteristics:
'   - Retrospective UK cohort
'   - N = 9,158 newly diagnosed type 2 diabetes patients
'   - Patients aged >=30 years
'   - 2-year follow-up after diabetes diagnosis
'   - Outcome = initiation of antihyperglycaemic medication
'   - Analysis = multivariable Cox proportional hazards regression
'
' Key observed treatment-initiation results:
'   - 51% of the overall cohort initiated treatment within 2 years.
'
'   Among patients with HbA1c >=7.5%:
'       73% initiated treatment within 180 days
'       81% initiated treatment within 1 year
'       87% initiated treatment within 2 years
'
'
' COX REGRESSION PARAMETERS USED:
'
' Age:
'   HR = 0.98 per additional year
' Older age was associated with a lower rate of treatment initiation.
'
' Sex:
'   Male vs Female HR = 0.91
' Female is therefore treated as the reference category in this function.
'
' HbA1c:
'   HbA1c >=7.5% vs <7.5%
'   HR = 2.44
' Elevated HbA1c was the strongest predictor of treatment initiation.
'
' Age x HbA1c interaction:
'   HR = 1.015 per year
'The original Cox model identified a significant interaction between age and
' HbA1c >=7.5%. Therefore, when HbA1c is >=7.5%, both the main HbA1c effect
' and the Age x HbA1c interaction are applied.
'
'
' ABSOLUTE-RISK CALIBRATION:
'
' Cox regression provides relative hazards but the publication does not provide
' an implementation-ready baseline survival function for direct use in the
' simulation.
'
' Therefore, the absolute baseline hazard was calibrated using the reported
' 1-year treatment initiation probability among patients with HbA1c >=7.5%.
'
' Reported:
'
'       P(Treatment by 1 year | HbA1c >=7.5%) = 0.81
'
' Calibration reference patient:
'
'       Age = 62.4 years
'       Female
'       HbA1c >=7.5%
'
' where 62.4 years is approximately the mean age of the study population.
'
'
' The calibrated baseline hazard is:
'
'       -Ln(1 - 0.81)
'       -------------------------
'       2.44 * (1.015 ^ 62.4)
'
'       ~= 0.2688
'
' The value 0.2688 is therefore used directly in the function to avoid
' recalculating the same constant every time the function is called.
'
'
' COX MODEL IMPLEMENTATION:
'
' For HbA1c <7.5%:
'
'       HR = 0.98 ^ (Age - 62.4)
'
' For HbA1c >=7.5%:
'
'       HR = 0.98 ^ (Age - 62.4)
'            * 2.44
'            * 1.015 ^ Age
'
' For males, the resulting HR is additionally multiplied by:
'
'       0.91
'
'
' The adjusted hazard is:
'
'       Adjusted Hazard = 0.2688 * HR
'
' and is converted to an annual probability using:
'
'       P = 1 - Exp(-Adjusted Hazard)
'
'
' INPUTS:
'
' Age
'   Current patient age in years.
'
' HbA1cPercent
'   Current HbA1c expressed in percentage units.
'   Example: 7.8 represents HbA1c = 7.8%.
'
' IsFemale
'   True  = Female
'   False = Male
'
'
' OUTPUT:
'
' A probability between 0 and 1 representing the probability that a diagnosed,
' currently pharmacologically untreated patient initiates antihyperglycaemic
' treatment during an annual cycle.
'
'
' MODEL USE:
'
' This function should ONLY be evaluated when:
'
'       DM_recognized = True
'       AND
'       DM_Treated = False
'
' Once DM_Treated becomes True, the function should no longer be evaluated
' unless treatment discontinuation/re-initiation is explicitly modelled.
'
'
' LIMITATION:
'
' The hazard ratios are based on the published Sinclair et al. Cox model,
' whereas the absolute baseline hazard (0.2688) is calibrated from the
' published 1-year treatment-initiation rate because an implementation-ready
' Cox baseline survival function was not reported.
'
'===============================================================================

Function AnnualDiabetesTreatmentProbability( _
    Age As Single, _
    HbA1c As Single, _
    Female As Boolean) As Double

    Dim HR As Double

    'Age effect:
    'HR = 0.98 per additional year.
    'Age is centered at the approximate study mean age (62.4 years).
    HR = 0.98 ^ (Age - 62.4)

    'Sex effect:
    'Female = reference category.
    'Male vs female HR = 0.91.
    If Not Female Then HR = HR * 0.91

    'HbA1c effect:
    'For HbA1c >=7.5%, apply:
    '   Main HbA1c effect HR = 2.44
    '   Age x HbA1c interaction HR = 1.015 per year
    If HbA1c >= 7.5 Then _
        HR = HR * 2.44 * (1.015 ^ Age)

    'Convert the patient-specific annual hazard into an annual probability.
    '
    '0.2688 = calibrated baseline annual hazard derived from the reported
    '81% 1-year treatment initiation among patients with HbA1c >=7.5%.
    AnnualDiabetesTreatmentProbability = _
        1 - Exp(-0.2688 * HR)

End Function
