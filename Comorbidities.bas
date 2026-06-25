Attribute VB_Name = "Comorbidities"
Option Explicit
Function CapProb(prob As Double) As Double
    ' Check if the input probability is greater than 1
    If prob > 1 Then
        ' Cap the probability to 1
        CapProb = 1
    Else
        ' Return the original probability if it is 1 or less
        CapProb = prob
    End If
End Function
Function ASCVD_check(patient As patient) As Boolean
'SOURCE: Muntner, P., Colantonio, L. D., Cushman, M., Goff, D. C., Jr, Howard, G., Howard, V. J., Kissela, B., Levitan, E. B., Lloyd-Jones, D. M., & Safford, M. M. (2014). Validation of the atherosclerotic cardiovascular disease Pooled Cohort risk equations. JAMA, 311(14), 1406–1415. https://doi.org/10.1001/jama.2014.2630
'EQUATION IN THE SUPPLEMENTARY

Dim risk As Byte
Dim risk_score As Boolean

With patient
    risk = Application.WorksheetFunction.Sum(Abs(.Hypertension), Abs(.DLP), Abs(.smoking), Abs(.BMI >= 30))

    If risk >= 2 Then risk_score = True

    If .CHD = True Or .MI = True Or .Stroke = True Or .PVD = True Or risk_score = True Then

      ASCVD_check = True
 
      
    End If
    
End With

End Function
Function ProbStroke(patient As patient)

'Lipid_ratio = Lipid ratio, T: H  - T:H indicates ratio of total:HDL cholesterol.
'Dim Lipid_ratio As Double
'Dim Q As Double
'Dim L As Double
'Dim a As Double
'Dim M As Double
'Dim T As Single
'Dim TCmmol As Double
'Dim HDLmmol As Double
Dim mu, sigma, U As Double


With patient
      
      
  'Source:Anderson, K. M., Odell, P. M., Wilson, P. W., & Kannel, W. B. (1991). Cardiovascular disease risk profiles. American heart journal, 121(1 Pt 2), 293–298. https://doi.org/10.1016/0002-8703(91)90861-b
  'This is the equation for i year probability of stroke
mu = 25.1067 _
    + Abs(.Female) * 0.1558 _
    + Log(.Age) * -3.0997 _
    + 0 * (Log(.Age) * Abs(.Female) _
        + 0 * ((Log(.Age)) ^ 2 * Abs(.Female))) _
    + -1.7556 * Log(.DBP) _
    + -0.3975 * Abs(.smoking) _
    + 0.0297 * Log(.TC / .HDL) _
    + -0.4047 * Abs(.DM) _
    + -0.2506 * (Abs(.DM) * Abs(.Female)) _
    + -0.2801 * .LVH _
    + 0 * (Abs(.LVH) * Abs(Abs(.Female) - 1))
          

          sigma = Exp(-0.4212 + 0 * mu)
          
          U = (Log(4) - mu) / sigma
          
          ProbStroke = 1 - Exp(-Exp(U))
          ProbStroke = 1 - (1 - ProbStroke) ^ (1 / 4)
      
          'adjusting the risk of stroke to BMI
          'source: Liu X, et al., A J-shaped relation of BMI and stroke: Systematic review and doseeresponse metaanalysis of 4.43 million participants, Nutrition, Metabolism & Cardiovascular Diseases (2018), https://doi.org/10.1016/ j.numecd.2018.07.004
          'the reported RR data was parametrized using curve expert professional to produce this equation
          'ProbStroke = ProbStroke * (2.44677054944366 + 1.45708742314377 * Cos(6.85415964516995E-02 * .BMI + 1.69241280922106))

      
      'Adjust TC and HDL from mg/dl to mmol/L
      'Reference: Haney EM, Huffman LH, Bougatsos C, et al. Screening for Lipid Disorders in Children and Adolescents [Internet]. Rockville (MD): Agency for Healthcare Research and Quality (US); 2007 Jul. (Evidence Syntheses, No. 47.) Appendix 2. Units of Measure Conversion Formulas.
      'TCmmol = .TC * 0.02586
      'HDLmmol = .HDL * 0.02586
      
     ' Lipid_ratio = .TC / .HDL
    '  T = .Age - .DM_Diagnosis_Age
      
     ' If .DM = True Then
            ' in case patient is diabetic
            'Reference: Kothari, V., Stevens, R. J., Adler, A. I., Stratton, I. M., Manley, S. E., Neil, H. A., & Holman, R. R. (2002). UKPDS 60: risk of stroke in type 2 diabetes estimated by the UK Prospective Diabetes Study risk engine. Stroke, 33(7), 1776-1781.
    '        Q = 0.00186 * 1.092 ^ (.DM_Diagnosis_Age - 55) * 0.7 ^ Abs(.Female) * 1.547 ^ Abs(.smoking) * 8.554 ^ Abs(.AF) * 1.122 ^ ((.SBP - 135.5) / 10) * 1.138 ^ (Lipid_ratio - 5.11)
            
   '         ProbStroke = 1 - Exp(-Q * 1.145 ^ T * (1 - 1.145 ^ 1) / (1 - 1.145))

   '   Else
            'in case patient is not diabetic
            ''Reference: Wolf, P. A., D'Agostino, R. B., Belanger, A. J., & Kannel, W. B. (1991). Probability of stroke: a risk profile from the Framingham Study. Stroke, 22(3), 312-318.
    '        If .Female = True Then
                  'in case of female
 '                 L = 0.0657 * .Age + 0.0197 * .SBP + 2.5432 * Abs(.anti_htn_drugs) - 0.0134 * .SBP * Abs(.anti_htn_drugs) + 0.5442 * Abs(.DM) + 0.5294 * Abs(.smoking) + 0.4326 * Abs(.CHD) + 1.1497 * Abs(.AF) + 0.8488 * Abs(.LVH)
  '                'is the average female patient value
  '                M = 7.5766
   '               a = L - M
                  ' one year = 0.9977 female
   '               ProbStroke = 1 - 0.9977 ^ Exp(a)
          '  Else
                  'in case of male
    '              L = 0.0505 * .Age + 0.014 * .SBP + 0.3263 * Abs(.anti_htn_drugs) + 0.3384 * Abs(.DM) + 0.5147 * Abs(.smoking) + 0.5195 * Abs(.CHD) + 0.6061 * Abs(.AF) + 0.8415 * Abs(.LVH)
                  'is the mean male patient value
      '            M = 5.677
       '           a = L - M
                  ' one year = 0.9948 male
         '        ProbStroke = 1 - 0.9948 ^ Exp(a)
       '     End If
            
     ' End If
'UKPDS 68 EQUATION: Patients must be diabetic
'    Dim lambda As Double
'    Dim rho As Double
'    Dim xb As Double
'    Dim integratedHazard As Double
'    Dim AgeInp As Double, HbA1CInp As Double, SBPInp As Double, TOTALHDL As Double
'
'
'    ' Weibull parameters for MI
'    lambda = -7.163
'    rho = 1.497
'
'    ' Apply UKPDS variable transformations
'    AgeInp = .DM_Diagnosis_Age - 52.59
'    HbA1CInp = .HbA1C - 7.09
'    SBPInp = (.SBP - 135.09) / 10
'    TOTALHDL = (.TC / .HDL - 5.23)
'
'    ' Linear predictor for MI
'    xb = lambda _
'        + 0.085 * AgeInp _
'        - 0.516 * Abs(.Female) _
'        + 0.355 * Abs(.smoking) _
'        + 0.128 * HbA1CInp _
'        + 0.276 * SBPInp _
'        + 1.19 * TOTALHDL _
'        + 1.742 * Abs(.HF) _
'        + 1.428 * Abs(.AF)

'    ' Integrated hazard from year t to t+1
'    integratedHazard = Exp(xb) * (((2) ^ rho) - (1 ^ rho))
'
'    ' Annual probability
'    ProbStroke = 1 - Exp(-integratedHazard)
'
'      End With

'ProbMI = Temp



End With

End Function

Function ProbDiabetes(patient As patient)
'Source: Lacy ME, Wellenius GA, Carnethon MR, Loucks EB, Carson AP, Luo X, et al. Racial Differences in the Performance of Existing Risk Prediction Models for Incident Type 2 Diabetes: The CARDIA Study. Diabetes Care [Internet]. 2016 Jan 11 [cited 2023 Dec 29];39(2):285–91. Available from: https://www.ncbi.nlm.nih.gov/pmc/articles/PMC4722943/ Print

With patient

'Dim HeightCoff As Single

'Height Coefficient
'Average heights of males and females in Egypt were used https://www.worlddata.info/average-bodyheight.php#google_vignette
'      If .Female = True Then
'      HeightCoff = -0.0326 * 160
'      Else
'      HeightCoff = -0.0326 * 173
'End If

'Race will be excluded from our calculation so we will multiply * 0
' This equation is used from Framingham as the patients characteristics are closest to out patients 30-74 years of age
'Source: Wilson PW, Meigs JB, Sullivan L, Fox CS, Nathan DM, D’Agostino RB. Prediction of incident diabetes mellitus in middle-aged adults: the Framingham Offspring Study. Archives of internal medicine. 2007 May 28;167(10):1068-74.

'ProbDiabetes = -9.9808 + 0.0173 * .Age + 0.4433 * 0 + 0.4981 * Abs(.family_history_DM) + 0.088 * .FBS + 0.0111 * .SBP + 0.0273 * .WC + HeightCoff - 0.0122 * .HDL + 0.00271 * .TG

ProbDiabetes = -18.607 - 0.0101 * .Age - 0.4308 * Abs(Abs(.Female) - 1) + 0.4383 * Abs(.family_history_DM) + 0.03922 * .BMI + 0.001 * .SBP - 0.0488 * .HDL + 0.0488 * .WC + 0.1398 * .FBS

'ProbDiabetes = -13.415 + 0.028 * .Age + 0.661 * Abs(.Female) + 0.412 * 0 + 0.079 * .FBS + 0.018 * .SBP - 0.039 * .HDL + 0.07 * .BMI + 0.481 * Abs(.family_history_DM)

ProbDiabetes = Exp(ProbDiabetes)
ProbDiabetes = ProbDiabetes / (1 + ProbDiabetes)
ProbDiabetes = 1 - (1 - ProbDiabetes) ^ (1 / 7)


End With


End Function


Function ProbHTN(patient As patient) As Double
'Source: Kawasoe M, Kawasoe S, Kubozono T, Ojima S, Kawabata T, Ikeda Y, Oketani N, Miyahara H, Tokushige K, Miyata M, Ohishi M. Development of a risk prediction score for hypertension incidence using Japanese health checkup data. Hypertension Research. 2022 Apr;45(4):730-40.

With patient

      ProbHTN = -16.428 + 0.0364 * .Age + 0.0632 * .BMI + 0.067 * .SBP + 0.0439 * .DBP + 0.3556 * Abs(.smoking) + 0.3639 * 0.5
      
      
      ProbHTN = Exp(ProbHTN)
      ProbHTN = ProbHTN / (1 + ProbHTN)
      ProbHTN = 1 - (1 - ProbHTN) ^ (1 / 5)

End With

End Function
Function ProbOSA(patient As patient) As Double

    ' Source: Tishler PV, Larkin EK, Schluchter MD, Redline S. Incidence of sleep-disordered breathing in an urban adult population: the relative importance of risk factors in the development of sleep-disordered breathing. Jama. 2003 May 7;289(17):2230-7.
    ' Threshold: AHI >= 15, variability-adjusted (7.5% 5-year incidence)

    Dim Age10 As Double
    Dim LP    As Double
    Dim P5yr  As Double
    Dim Male  As Double

    With patient
        Male = IIf(.Female, 0, 1)
        Age10 = .Age / 10

        LP = -13.129 _
           + 1.803 * Age10 _
           + 0.259 * .BMI _
           + 3.837 * Male _
           + (-0.7399 * Age10 * Male) _
           + (-0.0342 * Age10 * .BMI)

        P5yr = Exp(LP) / (1 + Exp(LP))
        If P5yr > 0.8 Then P5yr = 0.8

        ProbOSA = 1 - (1 - P5yr) ^ (1 / 5)
    End With

End Function
'Function ProbOSA(patient As patient) As Double
''Source:Ahlin, Sofie MD, PhDa,b,*; Manco, Melania MD, PhDc; Panunzi, Simona PhDd; Verrastro, Ornella MScb; Giannetti, Giulia MScb; Prete, Anna MDb; Guidone, Caterina MD, PhDb; Berardino, Alessandro Di Marco MDe; Viglietta, Luca MDe; Ferravante, Anna MDe; Mingrone, Geltrude MD, PhDb,f; Mormile, Flaminio MDe; Capristo, Esmeralda MD, PhDb. A new sensitive and accurate model to predict moderate to severe obstructive sleep apnea in patients with obesity. Medicine 98(32):p e16687, August 2019. | DOI: 10.1097/MD.0000000000016687
'    Dim x As Double
'    Dim Male As Double
'    Dim HasDM As Double
'
'    With patient
'
'        ' Explicit conversion
'        If .Female = True Then
'            Male = 0
'        Else
'            Male = 1
'        End If
'
'        If .DM = True Then
'            HasDM = 1
'        Else
'            HasDM = 0
'        End If
'
'        x = -9.144 + (0.087 * .Age) + (1.976 * Male) + (1.577 * HasDM) + (0.116 * .BMI)
'
'        ProbOSA = Exp(x) / (1 + Exp(x))
'
'    End With
'
'End Function
Function ProbOA(patient As patient) As Double
'Source: https://wrap.warwick.ac.uk/38642/1/WRAP_Muir_Ann_Rheum_Dis-2011-Zhang-1599-604.pdf
' Nottingham knee osteoarthritis risk prediction models
' Occupational_risk_OA is a string variable: 0=never, 1=seldom, 2=sometimes, 3=often, 4=always
Dim OA_risk As Integer


With patient

      Select Case .occupational_risk_OA
            Case Is = "never":            OA_risk = 0
            Case Is = "seldom":           OA_risk = 1
            Case Is = "sometimes":        OA_risk = 2
            Case Is = "often":            OA_risk = 3
            Case Is = "always":           OA_risk = 4
      End Select
 
           ProbOA = .Age * 0.056 + Abs(.Female) * 0.029 + .BMI * 0.089 + 0.245 * OA_risk + Abs(.family_history_OA) * 0.543 + Abs(.knee_injury) * 0.87 - 7.733
            
End With
            'ProbOA = Exp(ProbOA) / Exp(ProbOA) + 1
            ProbOA = Exp(ProbOA)
            ProbOA = ProbOA / (1 + ProbOA)
            'life time risk from 49-70 years
            'the risk equation is for 12 years. to calculate annual rate we apply the following equation
            ProbOA = 1 - ((1 - ProbOA) ^ (1 / 12))

End Function

Function ProbMI(patient As patient)
'This is the equation for 1 year probability of MI (Myocardial infaction)
'VALIDATED
'Source: D'Agostino, R. B., Russell, M. W., Huse, D. M., Ellison, R. C., Silbershatz, H., Wilson, P. W., & Hartz, S. C. (2000). Primary and subsequent coronary risk appraisal: new results from the Framingham study. American heart journal, 139(2 Pt 1), 272–281. https://doi.org/10.1067/mhj.2000.96469
'Source for diabetes part: Stevens, R. J., Kothari, V., Adler, A. I., Stratton, I. M., & United Kingdom Prospective Diabetes Study (UKPDS) Group (2001). The UKPDS risk engine: a model for the risk of coronary heart disease in Type II diabetes (UKPDS 56). Clinical science (London, England : 1979), 101(6), 671–679.
'T duration of diabetes
Dim T As Double
Dim temp As Double
Dim LipidRatio As Double
'CONTSTANT s=0.7764 in males ,0.7333 in females
Dim s As Double

Dim mu As Double
Dim sigma As Double
Dim U As Double

      With patient
'Source:Anderson, K. M., Odell, P. M., Wilson, P. W., & Kannel, W. B. (1991). Cardiovascular disease risk profiles. American heart journal, 121(1 Pt 2), 293–298. https://doi.org/10.1016/0002-8703(91)90861-b
'This is the equation for 1 year probability of MI
          mu = 11.0436 + Abs(.Female) * 5.1559 + Log(.Age) * -0.9302 + -2.631 * (Log(.Age) * Abs(.Female)) + _
          0.3472 * ((Log(.Age)) ^ 2 * Abs(.Female)) + _
          -0.5132 * Log(.DBP) + -0.2721 * Abs(.smoking) + -0.4228 * Log(.TC / .HDL) + -0.1764 * Abs(.DM) + -0.1184 * (Abs(.DM) * Abs(.Female)) + 0 * Abs(.LVH) + _
          -0.1702 * (Abs(.LVH) * Abs(Abs(.Female) - 1))


          sigma = Exp(3.4587 + -0.8647 * mu)

          U = (Log(4) - mu) / sigma

          ProbMI = 1 - Exp(-Exp(U))

          ProbMI = 1 - (1 - ProbMI) ^ (1 / 4)

            'T = .Age - .DM_Diagnosis_Age
            'LipidRatio = .TC / .HDL
            'If .Female = True Then
                  'Temp = 20.4049 - 0.0622 * .Age - 3.8236 * .Menopause + 0.0717 * .age_menopause - Log(.TC / .HDL) - 2.3607 * Log(.SBP) - 0.0097 * Log(.anti_htn_drugs * .SBP) - 0.5734 * .DM - 0.4041 * .smoking + 0.0461 * .drinking
               '   Temp = 20.4049 - 0.0622 * .Age - 3.8236 * Abs(.Menopause) + 0.0717 * (.Age * Abs(.Menopause)) - 0.8902 * Log(LipidRatio) - 2.3607 * Log(Patient.SBP) - 0.0097 * Abs(.anti_htn_drugs) * ((200 - .SBP) * (.SBP - 110) / 100) - 0.5734 * Abs(.DM) - 0.4041 * Abs(.smoking) + 0.0461 * .Number_of_alcohol
                  '     =20.4049 -0.0622  *CG77  -3.8236  *CG78             +0.0717  *(CG77*CG78)                -0.8902*(LN(CG79))         -2.3607  *LN(CG80)  -0.0097    *CG81                 *((200-CG80)*(CG80-110)/100)          -0.5734*CG82        -0.4041*CG83             +0.0461*CG84
           ' Else
           '       Temp = 12.7868 - 0.0405 * .Age - 0.9494 * Log(LipidRatio) - 1.0163 * Log(Patient.SBP) - 0.0161 * Abs(.anti_htn_drugs) * ((200 - .SBP) * (.SBP - 110) / 100) - 0.4412 * Abs(.DM) - 0.6042 * Abs(.smoking)
           ' End If
            
           ' If .Female = False Then s = 0.7764 Else s = 0.7333

         '  Temp = (Log(1) - Temp) / s
          ' Temp = 1 - Exp(-Exp(Temp))
            
            'If .DM = True Then
            'Temp = Temp * (1.183 ^ (.HbA1C - 6.72))
            'End If
            
            
    ' UKPDS 68 - Table 2: patients MUST be diabetic
    ' Equation 2: Myocardial infarction (MI)
    ' Weibull model annual probability

'    Dim lambda As Double
'    Dim rho As Double
'    Dim xb As Double
'    Dim integratedHazard As Double
'    Dim AgeInp As Double, HbA1CInp As Double, SBPInp As Double
'
'    ' Weibull parameters for MI
'    lambda = -4.977
'    rho = 1.257
'
'    ' Apply UKPDS variable transformations
'    AgeInp = .DM_Diagnosis_Age - 52.59
'    HbA1CInp = .HbA1C - 7.09
'    SBPInp = (.SBP - 135.09) / 10
'
'    ' Linear predictor for MI
'    xb = lambda _
'        + 0.055 * AgeInp _
'        - 0.826 * Abs(.Female) _
'        - 1.312 * 0 _
'        + 0.346 * Abs(.smoking) _
'        + 0.118 * HbA1CInp _
'        + 0.101 * SBPInp _
'        + 1.19 * Log((.TC / .HDL)) _
'        + 0.914 * Abs(.CHD) _
'        + 1.558 * Abs(.HF)
'
'    ' Integrated hazard from year t to t+1
'    integratedHazard = Exp(xb) * (((2) ^ rho) - (1 ^ rho))
'
'    ' Annual probability
'    ProbMI = 1 - Exp(-integratedHazard)
            
      End With

'ProbMI = Temp

End Function
Function ProbCHD(patient As patient)
'Source: Anderson KM, Odell PM, Wilson PW, Kannel WB. Cardiovascular disease risk profiles. American heart journal. 1991 Jan 1;121(1):293-8.
Dim mu As Double
Dim sigma As Double
Dim U As Double
With patient

 mu = 15.5222 + Abs(.Female) * 32.4811 + Log(.Age) * -1.6346 + -16.4933 * (Log(.Age) * Abs(.Female)) + _
          2.1059 * ((Log(.Age)) ^ 2 * Abs(.Female)) + _
          -0.867 * Log(.DBP) + -0.2789 * Abs(.smoking) + -0.7142 * Log(.TC / .HDL) + -0.2082 * Abs(.DM) + -0.1973 * (Abs(.DM) * Abs(.Female)) + -0.7195 * Abs(.LVH) _
          + 0 * (Abs(.LVH) * Abs(Abs(.Female) - 1))
          

          sigma = Exp(0.9341 + -0.2825 * mu)
          
          U = (Log(4) - mu) / sigma
          
          ProbCHD = 1 - Exp(-Exp(U))
          
          ProbCHD = 1 - (1 - ProbCHD) ^ (1 / 4)

End With

End Function

Function ProbNASH(patient As patient)
'Dim HDLmmol  As Single, UAmicromol As Single, TGmmol As Double
Dim UAmicromol As Double, HDLmmol As Double, TGmmol As Double
With patient

UAmicromol = .Uric_Acid * 59.48
HDLmmol = .HDL * 0.02586
TGmmol = .TG * 0.01129

'Source: Zhang Y, Shi R, Yu L, Ji L, Li M, Hu F. Establishment of a risk prediction model for non-alcoholic fatty liver disease in type 2 diabetes. Diabetes Therapy. 2020 Sep;11:2057-73.
ProbNASH = -7.35 + Abs(Abs(.Female) - 1) * 0.902 + .Age * 0.058 + (.Age - .DM_Diagnosis_Age) * -0.038 + .BMI * 0.216 + .WC * 0.041 + .DBP * 0.023 + TGmmol * 0.565 + HDLmmol * -1.283 + UAmicromol * 0.003


ProbNASH = Exp(ProbNASH)
ProbNASH = ProbNASH / (1 + ProbNASH)
ProbNASH = 1 - (1 - ProbNASH) ^ (1 / (73 - 18))

End With


'the equation estimates the risk of NAFLD
'Twenty percent of those affected by non-alcoholic fatty liver disease (NAFLD) eventually develop non-alcoholic steatohepatitis (NASH)
'soruce: Hughes AN, Oxford JT. A lipid-rich gestational diet predisposes offspring to nonalcoholic fatty liver disease: a potential sequence of events. Hepat Med. (2014) 6:15. 10.2147/HMER.S57500
ProbNASH = ProbNASH * 0.2

End Function
Function ProbDLP(patient As patient) As Double
'Source: Yang X, Xu C, Wang Y, Cao C, Qin T, Zhan S, et al. Risk prediction model of dyslipidaemia over a 5-year period based on the Taiwan MJ health check-up longitudinal database. Lipids in Health and Disease [Internet]. 2018 Nov 17 [cited 2023 Dec 26];17(1). Available from: https://lipidworld.biomedcentral.com/articles/10.1186/s12944-018-0906-2 Print
With patient
'this equation was not used as the lower part is used since it is more clinically defenisble and more clinically plausible
'If .Age < 35 Then
'
'      ProbDLP = 0
'
'Else
'
'      Dim Sex_coff As Byte
'
'      If .Female = True Then
'
'            Sex_coff = 2
'
'      Else
'
'            Sex_coff = 1
'
'      End If
'
'      ProbDLP = -5.2337 + Sex_coff * -0.229 + Abs(.family_history_DM) * 0.082 + .BMI * 0.0542 + .TG * 0.016 + .HDL * -0.00976 + .LDL * 0.0134
'
'      Transform the LogitP to probability over 5 years
'      ProbDLP = CDbl(Exp(ProbDLP) / (1 + Exp(ProbDLP)))
'      Calculate the annual probability
'
'      ProbDLP = 1 - (1 - ProbDLP) ^ (1 / 5)
'
'End If

'source:Combination between Clinical expert opinion and reference: Marzieh Nikparvar, Mohadeseh Khaladeh, Yousefi H, Akbar Etebarian, Behzad Moayedi, Masoumeh Kheirandish. Dyslipidemia and its associated factors in southern Iranian women, Bandare-Kong Cohort study, a cross-sectional survey. Scientific Reports [Internet]. 2021 Apr 28 [cited 2023 Dec 28];11(1). Available from: https://www.ncbi.nlm.nih.gov/pmc/articles/PMC8080669/ Print
If .ASCVD = True Then

      If .LDL >= 70 Or .HDL <= 40 Or .TG >= 200 Then

            ProbDLP = 1

      Else

            ProbDLP = 0

      End If

Else

      If .LDL >= 140 Or .HDL <= 40 Or .TG >= 200 Then

            ProbDLP = 1

      Else

            ProbDLP = 0

      End If

End If

End With
End Function
Function probKeto(patient As patient)

'check if patient is diabetic or not
'Source: Davis TME, Davis W. Incidence and associates of diabetic ketoacidosis in a community-based cohort: the Fremantle Diabetes Study Phase II. BMJ Open Diabetes Res Care. 2020;8(1):e000983. doi:10.1136/bmjdrc-2019-000983

With patient
' DM_type1= true means type 1 DM.
If .DM = True Then

      If .Keto_history = False Then
            'this ratio is for incidence rate of first incidence of DKA
            If .DM_type1 = True Then probKeto = 166.4 / 10000 Else probKeto = 8.3 / 10000
      Else
            'this ratio is for overall incidence, first and recurrent
            If .DM_type1 = True Then probKeto = 178.6 / 10000 Else probKeto = 13.3 / 10000
      End If
      
Else

      probKeto = 0

End If
End With

End Function
Function ProbMA(patient As patient) As Double
Dim DM_duration As Double
With patient

If .Retino = True And .DM = True Then
            
'      DM_duration = .Age - .DM_Diagnosis_Age
      
'Source: Brown JB, Russell A, Chan W, Pedula K, Aickin M. The global diabetes model: user friendly version 3.0. Diabetes research and clinical practice. 2000 Nov 1;50:S15-46.
'
'                        DM_duration = .Age - .DM_Diagnosis_Age
'                        If DM_duration >= 15 Then
'                            ProbMA = 0.08
'                         ElseIf DM_duration >= 10 Then
'                            ProbMA = 0.092
'                        ElseIf DM_duration <= 5 Then
'                            ProbMA = 0.095
'                        Else
'                            ProbMA = 0.047
'                        End If

'This is annual risk based on calculations from Martín-Merino E, Fortuny J, Rivero-Ferrer E, García-Rodríguez LA. Incidence of retinal complications in a cohort of newly diagnosed diabetic patients. PLoS One. 2014 Jun 25;9(6):e100283.
' The calculations are in the excel file named Macular Edema in the folder Macular Edema calculations
ProbMA = 0.00293
' to adjust for hba1c, Source:Brown JB, Russell A, Chan W, Pedula K, Aickin M. The global diabetes model: user friendly version 3.0. Diabetes research and clinical practice. 2000 Nov 1;50:S15-46.
ProbMA = ProbMA * (.HbA1C / 10) ^ 1.2

Else

    ProbMA = 0

End If

End With

End Function

Function ProbNephro(patient As patient)
Dim DM_duration As Double
Dim HbA1Cmmol As Double

'Source: Vu, V.N., Le Thi, K.A., Dinh, T.M.D., Nguyen, T.B.M., Le Thi, D.H. and Vu, T.T., 2021. Applying Logistic Regression to Predict Diabetic Nephropathy Based on Some Clinical and Paraclinical Characteristics of Type 2 Diabetic Patients. VNU Journal of Science: Medical and Pharmaceutical Sciences, 37(2).
'Source for extended equation: Grover G, Gadpayle AK, Sabharwal A. Identifying patients with diabetic nephropathy based on serum creatinine in the presence of covariates in type-2 diabetes: A retrospective study. Biomed Res India. 2012 Oct 1;23(4):1-1.
'Reference for HbA1Cmmol: https://ebmcalc.com/GlycemicAssessment.htm
'Reference for TCmmol: Haney EM, Huffman LH, Bougatsos C, et al. Screening for Lipid Disorders in Children and Adolescents [Internet]. Rockville (MD): Agency for Healthcare Research and Quality (US); 2007 Jul. (Evidence Syntheses, No. 47.) Appendix 2. Units of Measure Conversion Formulas.

With patient

'HbA1Cmmol = (28.7 * .HbA1C - 46.7) / 18.015

'TCmmol = .TC * 0.02586

DM_duration = .Age - .DM_Diagnosis_Age
      If .DM = True Then
             If .DM_type1 = False Then
            'Source: Adler, A. I., Stevens, R. J., Manley, S. E., Bilous, R. W., Cull, C. A., Holman, R. R., & UKPDS GROUP (2003). Development and progression of nephropathy in type 2 diabetes: the United Kingdom Prospective Diabetes Study (UKPDS 64). Kidney international, 63(1), 225–232. https://doi.org/10.1046/j.1523-1755.2003.00712.x
             'this value for type 2 DM patients to progress from no nephropathy to microalbumiuria which is the earliest marker of DN.
             
             ProbNephro = 0.02

            'ProbNephro = -56.333 + 0.704 * TCmmol + 0.469 * HbA1Cmmol
            
            'ProbNephro = Exp(ProbNephro)
            'ProbNephro = ProbNephro / (1 + ProbNephro)
            
            'ProbNephro = -15.567 + 0.0679 * .FBS + 0.0908 * .DBP + 0.0828 * .SBP + 0.0035 * .LDL + 0.2117 * DM_duration + -0.02493 * .DM_Diagnosis_Age
            'ProbNephro = Exp(ProbNephro)
            'ProbNephro = ProbNephro / (1 + ProbNephro)
            
            'ProbNephro = -21.935 + 0.0454 * .FBS + 0.0977 * .DBP + 0.0784 * .SBP + 0.0446 * .LDL + 0.1339 * DM_duration + 0.0127 * .DM_Diagnosis_Age
            'ProbNephro = Exp(ProbNephro)
            'ProbNephro = ProbNephro / (1 + ProbNephro)
            
            
            Else
     
     'Source: Huang CY, Ting WH, Lo FS, Tsai JD, Sun FJ, Chan CI, Chiang YT, Lin CH, Cheng BW, Wu YL, Hung CM, Lee YJ. Factors associated with diabetic nephropathy in children, adolescents, and adults with type 1 diabetes. J Formos Med Assoc. 2017 Dec;116(12):924-932. doi: 10.1016/j.jfma.2017.09.015. Epub 2017 Oct 23. PMID: 29070437.
            
            ProbNephro = -20.85 + 0.13 * .Age + 0.56 * .HbA1C + 0.06 * .SBP + 0.03 * .DBP
            ProbNephro = Exp(ProbNephro)
            ProbNephro = ProbNephro / (1 + ProbNephro)
            
            End If
      Else
      
            ProbNephro = 0
      
      End If
      
End With


End Function
Function ProbUlcer(patient As patient) As Double

    '-------------------------------------------------------------------------
    ' Purpose:
    '   Estimate the annual probability of developing a diabetic foot ulcer
    '   for a patient with diabetes.
    '
    ' Approach:
    '   1. Use the published logistic regression equation to estimate the
    '      patient's cumulative probability of diabetic foot ulcer.
    '   2. Convert the cumulative probability into an annual risk for the
    '      most recent year using a Weibull-based increasing-risk assumption.
    '
    ' Source for logistic regression model:
    '   Ahmadi, S. A. Y., Shirzadegan, R., Mousavi, N., Farokhi, E.,
    '   Soleimaninejad, M., & Jafarzadeh, M. (2021).
    '   Designing a Logistic Regression Model for a Dataset to Predict Diabetic
    '   Foot Ulcer in Diabetic Patients: High-Density Lipoprotein (HDL)
    '   Cholesterol Was the Negative Predictor.
    '   Journal of Diabetes Research, 2021, 5521493.
    '   https://doi.org/10.1155/2021/5521493
    '-------------------------------------------------------------------------

    Dim DiabetesDuration As Double
    Dim LinearPredictor As Double
    Dim CumulativeRisk As Double
    Dim WeibullShape As Double

    With patient

        '---------------------------------------------------------------------
        ' The ulcer risk is only estimated for patients with diabetes.
        ' If the patient does not have diabetes, the probability of diabetic
        ' foot ulcer is set to zero.
        '---------------------------------------------------------------------
        If .DM = True Then

            '-----------------------------------------------------------------
            ' Calculate diabetes duration as:
            '   Current age - age at diabetes diagnosis
            '
            ' This represents the number of years the patient has lived with
            ' diabetes.
            '-----------------------------------------------------------------
            DiabetesDuration = .Age - .DM_Diagnosis_Age

            '-----------------------------------------------------------------
            ' Avoid a diabetes duration of zero.
            '
            ' The Weibull-based annual risk equation requires division by
            ' DiabetesDuration. Therefore, if duration is zero, we assume a
            ' minimum duration of one year.
            '-----------------------------------------------------------------
            If DiabetesDuration = 0 Then DiabetesDuration = 1

            '-----------------------------------------------------------------
            ' Weibull shape parameter.
            '
            ' A value of 2 assumes that the instantaneous risk of diabetic foot
            ' ulcer increases over diabetes duration.
            '
            ' Interpretation:
            '   WeibullShape = 1   -> constant risk over time
            '   WeibullShape > 1   -> increasing risk over time
            '   WeibullShape < 1   -> decreasing risk over time
            '
            ' Here, WeibullShape = 2 is used as a simple increasing-risk
            ' assumption.
            '-----------------------------------------------------------------
            WeibullShape = 2

            '-----------------------------------------------------------------
            ' Step 1:
            ' Calculate the linear predictor from the logistic regression model.
            '
            ' The model estimates the log-odds of diabetic foot ulcer using:
            '   Age
            '   BMI
            '   Fasting blood sugar
            '   HDL cholesterol
            '   Insulin use/status
            '
            ' Note:
            '   HDL has a negative coefficient, meaning higher HDL is associated
            '   with lower predicted ulcer risk in the original model.
            '-----------------------------------------------------------------
            LinearPredictor = -12.98725 _
                              + 0.1331305 * .Age _
                              + 0.1944625 * .BMI _
                              + 0.0108864 * .FBS _
                              - 0.1184257 * .HDL _
                              + 0.9855977 * Abs(.Insulin)

            '-----------------------------------------------------------------
            ' Step 2:
            ' Convert the linear predictor from log-odds to probability using
            ' the logistic transformation:
            '
            '   Probability = exp(linear predictor) /
            '                 [1 + exp(linear predictor)]
            '
            ' This gives the predicted cumulative probability of diabetic foot
            ' ulcer based on the patient's current characteristics.
            '-----------------------------------------------------------------
            CumulativeRisk = Exp(LinearPredictor)
            CumulativeRisk = CumulativeRisk / (1 + CumulativeRisk)

            '-----------------------------------------------------------------
            ' Step 3:
            ' Convert the cumulative risk into an annual risk for the last year.
            '
            ' A Weibull distribution is assumed for time from diabetes diagnosis
            ' to diabetic foot ulcer. The Weibull curve is calibrated so that
            ' the cumulative risk at the current diabetes duration equals the
            ' cumulative risk predicted by the logistic regression model.
            '
            ' Formula:
            '
            '   Annual risk =
            '   1 - (1 - cumulative risk) ^
            '       [1 - ((D - 1) / D) ^ k]
            '
            ' Where:
            '   D = diabetes duration
            '   k = Weibull shape parameter
            '
            ' This estimates the conditional probability of developing diabetic
            ' foot ulcer during the most recent year, given that the patient
            ' was ulcer-free at the start of that year.
            '-----------------------------------------------------------------
            ProbUlcer = 1 - (1 - CumulativeRisk) ^ _
                        (1 - ((DiabetesDuration - 1) / DiabetesDuration) ^ WeibullShape)

        Else

            '-----------------------------------------------------------------
            ' If the patient does not have diabetes, diabetic foot ulcer risk is
            ' assumed to be zero in this model.
            '-----------------------------------------------------------------
            ProbUlcer = 0

        End If

    End With

End Function
Function ProbHypogly(patient As patient)
' these values are for the major hypoglycemic events. As per the interviews with the clinical experts, they mentioned the major hypoglycemic events are the cost drivers and minor events are low cost.
 With patient
 
      If IsNull(.Diabetes_treatment_ID) Or .Diabetes_treatment_ID = 0 Then
      
            ProbHypogly = 0
      
      Else
      
            'this function is intended to filter a multidimensional array based on upto 4 criteria. It needs the range, criteria and criteria columns
            ProbHypogly = .Diabetes_Drug.Major_Hypoglycemia
           
      End If

End With

End Function
Function ProbRetino(patient As patient) As Double
'Source: Fe'li SN, Emamian MH, Yaseri M, et al. Development and validation of prediction models for diabetic retinopathy in type 2 diabetes patients. PLoS One. 2025;20(7):e0325814. Published 2025 Jul 10. doi:10.1371/journal.pone.0325814
'Note that in the article BG is non-fasting blood glucose. we used estimated average as proxy.
Dim DM_duration As Double
Dim BG As Double
Dim MBP As Double
Dim LP5 As Double
Dim ProbRetino5 As Double

With patient

If .DM = True Then

      DM_duration = .Age - .DM_Diagnosis_Age

MBP = (2 * .DBP + .SBP) / 3
BG = (28.7 * .HbA1C) - 46.7

LP5 = -5.2855 + (DM_duration * 0.1412) + (0.0175 * MBP) + (0.007 * BG)
ProbRetino5 = Exp(LP5) / (1 + Exp(LP5))
ProbRetino = 1 - (1 - ProbRetino5) ^ (1 / 5)

Else

ProbRetino = 0

End If

End With

End Function

Function ProbNeuro(patient As patient)
'this is focused on diabetic neuropathy not neuropathy in general

Dim DiabetesDuration As Double
Dim WeibullShape As Byte
Dim LinearPredictor As Double
Dim CumulativeRisk As Double

With patient
      If .DM = True Then
        
            DiabetesDuration = .Age - .DM_Diagnosis_Age
            If DiabetesDuration = 0 Then DiabetesDuration = 1
            
            Dim RiskNow As Double, RiskLater As Double
            
            
            'Source:Young, M. J., Boulton, A. J., MacLeod, A. F., Williams, D. R., & Sonksen, P. H. (1993). A multicentre study of the prevalence of diabetic peripheral neuropathy in the United Kingdom hospital clinic population. Diabetologia, 36(2), 150–154. https://doi.org/10.1007/BF00400697
            ' the study provides the prevalence so we are calculating the prevalence now and after four year and calculating the difference to have an estimate of the four year cumulative incidence
            RiskNow = -4.56 + Abs(Abs(.DM_type1) - 1) * 0.085 + DiabetesDuration * 0.044 + .Age * 0.054 + Abs(.Female) * -0.095
            RiskNow = Exp(RiskNow)
            RiskNow = RiskNow / (1 + RiskNow)

            RiskLater = -4.56 + Abs(Abs(.DM_type1) - 1) * 0.085 + (DiabetesDuration + 4) * 0.044 + (.Age + 4) * 0.054 + Abs(.Female) * -0.095
            RiskLater = Exp(RiskLater)
            RiskLater = RiskLater / (1 + RiskLater)

            CumulativeRisk = RiskLater - RiskNow
            
        '-----------------------------------------------------------------
            ' Weibull shape parameter.
            '
            ' A value of 2 assumes that the instantaneous risk of diabetic foot
            ' ulcer increases over diabetes duration.
            '
            ' Interpretation:
            '   WeibullShape = 1   -> constant risk over time
            '   WeibullShape > 1   -> increasing risk over time
            '   WeibullShape < 1   -> decreasing risk over time
            '
            ' Here, WeibullShape = 2 is used as a simple increasing-risk
            ' assumption.
            '-----------------------------------------------------------------
            WeibullShape = 2

' then we need to adjust to the assumption that the risk is increasing over the duration of the four years and we capture the risk in the last year of the four
' and we are assuming increasing risk on a weibull distribution by assuming a shape of 2
            '-----------------------------------------------------------------
            ProbNeuro = 1 - (1 - CumulativeRisk) ^ _
                        (1 - ((4 - 1) / 4) ^ WeibullShape)
      
      Else
            
            ProbNeuro = 0
            
      End If
      
End With


End Function


Function ProbPVD(patient As patient)

  'Refernce:Murabito, J. M., D'Agostino, R. B., Silbershatz, H., & Wilson, W. F. (1997). Intermittent claudication. A risk profile from The Framingham Heart Study
  'https://www.ahajournals.org/doi/10.1161/01.CIR.96.1.44
    Dim BPCS As Double
    Dim BPCD As Double
    Dim BPC As Double
    Dim L As Double
    Dim P As Double

    With patient
   

'BPC = Blood Pressure Coefficient
'blood pressure was considered normal when the systolic blood pressure was <130 mm Hg and the diastolic blood pressure was <85 mm Hg.
'High-normal blood pressure was defined as a systolic blood pressure of 130 to 139 mm Hg or a diastolic blood pressure of 85 to 89 mm Hg.
'Stage 1 hypertension occurred when the systolic blood pressure was 140 to 159 mm Hg or the diastolic blood pressure was 90 to 99 mm Hg.
'Stage 2 or greater hypertension occurred when the systolic blood pressure was >=160 mm Hg or the diastolic blood pressure was >=100 mm Hg.

          ' calculate BPC based on SBP
          If .SBP < 130 Then
              BPCS = 0
          ElseIf .SBP < 140 Then
              BPCS = 0.2621
          ElseIf .SBP < 160 Then
              BPCS = 0.4067
          Else
              BPCS = 0.7977
          End If
          
          ' calculate BPC based on DBP
          If .DBP < 85 Then
              BPCD = 0
          ElseIf .DBP < 90 Then
              BPCD = 0.2621
          ElseIf .DBP < 100 Then
              BPCD = 0.4067
          Else
              BPCD = 0.7977
          End If
          
'If the systolic and diastolic pressures fell into different blood pressure stages, the higher stage was used to classify the blood pressure status.

          If BPCS >= BPCD Then BPC = BPCS Else BPC = BPCD
          
          ' calculate L
            L = -8.9152 + 0.5033 * (Abs(Abs(.Female) - 1)) + 0.0372 * .Age + BPC + 0.9503 * Abs(.DM) + 0.0314 * .Number_of_cigaretts + 0.0048 * .TC + 0.9939 * Abs(.CHD)

      End With

    ' calculate P
    P = (1 / (1 + Exp(-L)))
    
    ' calculate P1
    ProbPVD = P
    'The equation above gives 4-year risk. to calculate 1-year risk we use the following
    ProbPVD = 1 - (1 - P) ^ (1 / 4)
    
    
End Function

Function ProbCKD(patient As patient) As Double
'Source: Chien KL, Lin HJ, Lee BC, Hsu HC, Lee YT, Chen MF. A prediction model for the risk of incident chronic kidney disease. The American journal of medicine. 2010 Sep 1;123(9):836-46.
With patient

    ' Calculate the score for age
      Dim ageScore As Integer
      
      If .Age >= 65 Then
          ageScore = 8
      ElseIf .Age >= 55 Then
          ageScore = 5
      ElseIf .Age >= 45 Then
          ageScore = 3
      Else
          ageScore = 0
      End If
      
    
    ' Calculate the score for BMI
    'disabled BMI scoring and setted it to the reference BMI score in the other publication because in this study only non obese patients were considered
    Dim bmiScore As Integer
'    If .BMI >= 26 Then
'        bmiScore = 2
'    ElseIf .BMI >= 21 Then
'        bmiScore = 1
'    Else
'        bmiScore = 0
'    End If
    bmiScore = 1
    
    
    ' Calculate the score for diastolic blood pressure
    Dim dbpScore As Integer
    If .DBP >= 80 Then
        dbpScore = 2
    ElseIf .DBP >= 66 Then
        dbpScore = 1
    Else
        dbpScore = 0
    End If
    
    ' Calculate the score for diabetes
    Dim diabetesScore As Integer
    If .DM = True And .DM_type1 = False Then
        diabetesScore = 1
    Else
        diabetesScore = 0
    End If
    
    ' Calculate the score for history of stroke
    Dim strokeScore As Integer
    If .Stroke Then
        strokeScore = 4
    Else
        strokeScore = 0
    End If
    

    ' Calculate the total score
      Dim totalScore As Integer
      totalScore = ageScore + bmiScore + dbpScore + diabetesScore + strokeScore

    ' Match the total score to the estimated risk
      Dim EstimatedRisk As Single
    Select Case totalScore
        Case 0
            EstimatedRisk = 0.02
        Case 1
            EstimatedRisk = 0.02
        Case 2
            EstimatedRisk = 0.03
        Case 3
            EstimatedRisk = 0.04
        Case 4
            EstimatedRisk = 0.06
        Case 5
            EstimatedRisk = 0.07
        Case 6
            EstimatedRisk = 0.1
        Case 7
            EstimatedRisk = 0.13
        Case 8
            EstimatedRisk = 0.18
        Case 9
            EstimatedRisk = 0.23
        Case 10
            EstimatedRisk = 0.31
        Case 11
            EstimatedRisk = 0.39
        Case 12
            EstimatedRisk = 0.49
        Case 13
            EstimatedRisk = 0.6
        Case 14
            EstimatedRisk = 0.72
        Case 15
            EstimatedRisk = 0.82
        Case 16
            EstimatedRisk = 0.9
        Case Is = 17
            EstimatedRisk = 0.96
        
    End Select


    ProbCKD = 1 - (1 - EstimatedRisk) ^ (1 / 4)

'Adjust the risk of CKD by the BMI of bigger range because the original study only considered non obese patients
'Herrington, W. G., Smith, M., Bankhead, C., Matsushita, K., Stevens, S., Holt, T., ... & Woodward, M. (2017). Body-mass index and risk of advanced chronic kidney disease: prospective analyses from a primary care cohort of 1.4 million adults in England. PloS one, 12(3), e0173515.
'the paper reported Hazard Ratios for CKD stage 4-5 which was assumed to be as the RR
      If .BMI < 20 Then
            ProbCKD = ProbCKD * 0.98
      'reference group
      ElseIf .BMI < 25 Then
            ProbCKD = ProbCKD * 1
      ElseIf .BMI < 30 Then
            ProbCKD = ProbCKD * 1.2
      ElseIf .BMI < 35 Then
            ProbCKD = ProbCKD * 1.54
      Else
            ProbCKD = ProbCKD * 2.19
      End If

End With

End Function

Function ProbHF(patient As patient)
'Source: Khan, S. S., Ning, H., Shah, S. J., Yancy, C. W., Carnethon, M., Berry, J. D., Mentz, R. J., O'Brien, E., Correa, A., Suthahar, N., de Boer, R. A., Wilkins, J. T., & Lloyd-Jones, D. M. (2019). 10-Year Risk Equations for Incident Heart Failure in the General Population. Journal of the American College of Cardiology, 73(19), 2388–2397. https://doi.org/10.1016/j.jacc.2019.02.057

'ProbHF =
End Function
