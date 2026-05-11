Attribute VB_Name = "Utilities"
Option Explicit

Function Agg_Utility(Disutility_Arr() As Double, patient As patient, Optional Utility_Method As String = "ADE") As Double
'This function aggregates the impact of various disutilities in the model to provide the final utility index value of the patient after considering all the effects
'this function allows the choice of the aggregation methods

Dim k As Integer
Dim Temp_Utility As Double

Select Case Utility_Method     ' Choose the aggregation method based on the user selection.

      Case "Multiplicative"    ' Multiplicative disutility.
      'the multiplicative approach assumes that for each further condition a
      'patient population experiences, the disutility impact is proportional to the disutility
      'already experienced. Consequently, the absolute disutility for each additional
      'condition occurs at a diminishing rate. For higher numbers of morbidities, the
      'multiplicative approach would tend towards zero
      Agg_Utility = Product_NotZero(Disutility_Arr) * BaseUtilityProxy(patient)
      
      Case "Additive"
      'the additive model is assumed to be equivalent to
      'the sum of the disutilities measured in populations with only the individual conditions
      
            Temp_Utility = BaseUtilityProxy(patient)
            'convert multiplicative disutilities to subtractive disutilities assuming the base utility
            For k = LBound(Disutility_Arr) To UBound(Disutility_Arr)
            
                  Disutility_Arr(k) = Temp_Utility - Disutility_Arr(k) * Temp_Utility
            
            Next k
            
            'subtract all subtractive disutilities from the base utility
            For k = LBound(Disutility_Arr) To UBound(Disutility_Arr)
            
                  Temp_Utility = Temp_Utility - Disutility_Arr(k)
            
            Next k
            
            Agg_Utility = Temp_Utility
      
      Case "Minimum"
      'The minimum approach takes the utility value for the worst condition amongst a set of conditions as equivalent to the overall utility
            
            ' Sort the utility values in descending order.
            Call ShellSortDescending(Disutility_Arr)
            
            ' Remove trailing zeros: find the highest index with a nonzero value.
            For k = UBound(Disutility_Arr) To LBound(Disutility_Arr) Step -1
                
                If Disutility_Arr(k) <> 0 Then
                
                        ReDim Preserve Disutility_Arr(LBound(Disutility_Arr) To k)
                        Exit For
                
                End If
            
            Next k
            
            Agg_Utility = Application.WorksheetFunction.Min(Disutility_Arr) * BaseUtilityProxy(patient)
      
      Case "ADE"
      'The adjusted decrement estimator (ADE) assumes that the upper limit for the utility
      'value for any joint health condition is set at the minimum for any single health state
      'from the set of health state utility values166. Furthermore, each additional health state
      'utility value is a function of the health state utility values patients already have,
      'adjusted by the minimum value.
            Agg_Utility = ADE(Disutility_Arr, BaseUtilityProxy(patient))
      
      Case "Jia 2005 only"
      'Exclude the effect of all the extra disutilities and rely only on the conditions included in the initial regression equation
            Agg_Utility = BaseUtilityProxyAll(patient)

End Select


End Function

Function BaseUtilityProxy(patient As patient)

With patient

'All comorbidities turned off to be handled from submodels
      BaseUtilityProxy = UtilityIndex(.Age, .Female, "White", 3, .smoking, .physical_activity, False, _
      False, False, False, False, False, .BMI)

End With

End Function

Function BaseUtilityProxyAll(patient As patient)

With patient

      BaseUtilityProxyAll = UtilityIndex(.Age, .Female, "White", 3, .smoking, .physical_activity, False, _
      .Hypertension, .DM, .CHD, .Stroke_history, False, .BMI)

End With

End Function

Sub testutility()

Debug.Print UtilityIndex(64, 0, "White", 3, 0, 1, False, _
      False, False, False, False, False, 23)


End Sub

Function UtilityIndex(Age As Single, Female As Boolean, Race As String, Income As Double, SmokeStatus As Boolean, _
PhysicalActivity As Boolean, Asthma As Boolean, Hypertension As Boolean, Diabetes As Boolean, CHD As Boolean, Stroke As Boolean, _
Emphysema As Boolean, BMI As Single) As Single

'Source: Jia H, Lubetkin EI. The impact of obesity on health-related quality-of-life in the general adult US population. Journal of public health. 2005 Jun 1;27(2):156-64.

UtilityIndex = 1

' Determine age coefficient
If Age < 25 Then
      UtilityIndex = UtilityIndex + 0
ElseIf Age < 45 Then
      UtilityIndex = UtilityIndex + -0.026
ElseIf Age < 65 Then
      UtilityIndex = UtilityIndex + -0.077
Else
      UtilityIndex = UtilityIndex + 0.099
End If

' Determine sex coefficient
If Female = False Then UtilityIndex = UtilityIndex + 0.023

' Determine race coefficient
If Race = "Black" Then
      UtilityIndex = UtilityIndex + 0.025
ElseIf Race = "Asian" Then
      UtilityIndex = UtilityIndex + 0.046
ElseIf Race = "American Indian" Then
      UtilityIndex = UtilityIndex + 0.023
ElseIf Race = "Hispanic" Then
      UtilityIndex = UtilityIndex + 0.006
End If

' Determine income coefficient
If Income < 1 Then
      UtilityIndex = UtilityIndex + -0.124
ElseIf Income < 1.25 Then
      UtilityIndex = UtilityIndex + -0.093
ElseIf Income < 2 Then
      UtilityIndex = UtilityIndex + -0.061
ElseIf Income < 4 Then
      UtilityIndex = UtilityIndex + -0.027
Else
      UtilityIndex = UtilityIndex + 0
End If

' Determine smoke status coefficient
If SmokeStatus Then UtilityIndex = UtilityIndex + -0.046

' Determine physical activity coefficient
If PhysicalActivity Then UtilityIndex = UtilityIndex + 0.046

' Determine disease coefficients
If Asthma = True Then UtilityIndex = UtilityIndex + -0.045
If Hypertension = True Then UtilityIndex = UtilityIndex + -0.053
If Diabetes = True Then UtilityIndex = UtilityIndex + -0.042
If CHD = True Then UtilityIndex = UtilityIndex + -0.083
If Stroke = True Then UtilityIndex = UtilityIndex + -0.08
If Emphysema = True Then UtilityIndex = UtilityIndex + -0.012

' Determine obesity coefficient
If BMI < 18.5 Then
      UtilityIndex = UtilityIndex + -0.029
ElseIf BMI < 25 Then
      UtilityIndex = UtilityIndex + 0
ElseIf BMI < 30 Then
      UtilityIndex = UtilityIndex + -0.013
ElseIf BMI < 35 Then
      UtilityIndex = UtilityIndex + -0.033
Else
      UtilityIndex = UtilityIndex + -0.073
End If

End Function

Function Decrement_Duration_Adj(Utility_Decrement As Double, Duration As Double, Interval_Length As Double) As Double

Dim Prop_Disutility As Double
Dim Duration_Effective As Double

If Duration Mod (Interval_Length * 12) = 0 Then
      
      Duration_Effective = Interval_Length

Else
      
      Duration_Effective = Duration Mod (Interval_Length * 12)

End If

Prop_Disutility = Duration_Effective / (Interval_Length * 12)

Decrement_Duration_Adj = (1 - Prop_Disutility) + Utility_Decrement * Prop_Disutility

End Function

Function ADE(Disutility_Arr() As Double, Base_Utility As Double) As Double
'Thompson, A. J., Sutton, M., & Payne, K. (2019). Estimating joint health condition utility values. Value in Health, 22(4), 482-490.

    Dim k As Long
    Dim newUB As Long
    Dim MinUtility As Double

    ' Convert disutilities to utilities by multiplying each element by Base_Utility.
    For k = LBound(Disutility_Arr) To UBound(Disutility_Arr)
    
        Disutility_Arr(k) = Disutility_Arr(k) * Base_Utility
    
    Next k

    ' Sort the utility values in descending order.
    Call ShellSortDescending(Disutility_Arr)

    ' Remove trailing zeros: find the highest index with a nonzero value.
    For k = UBound(Disutility_Arr) To LBound(Disutility_Arr) Step -1
        If Disutility_Arr(k) <> 0 Then
            newUB = k
            Exit For
        End If
    Next k
    
    ReDim Preserve Disutility_Arr(LBound(Disutility_Arr) To newUB)

    ' Combine pairs using the ADE estimator:
    ' Loop from the end (lowest utility) toward the beginning.
    For k = UBound(Disutility_Arr) To LBound(Disutility_Arr) + 1 Step -1
        MinUtility = Application.WorksheetFunction.Min(Disutility_Arr(k), Disutility_Arr(k - 1))
        Disutility_Arr(k - 1) = MinUtility - MinUtility * ((Base_Utility - Disutility_Arr(k)) * (Base_Utility - Disutility_Arr(k - 1)))
    Next k

    ' Return the combined utility (now stored at the lower bound).
    ADE = Disutility_Arr(LBound(Disutility_Arr))
End Function
