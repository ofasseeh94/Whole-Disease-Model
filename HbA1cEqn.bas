Attribute VB_Name = "HbA1cEqn"
Function Hba1cProfileParametric(time_since_baseline As Single, hba1c_baseline As Single, treatment_effect As Single, kOvershoot As Single, kHba1cLimit As Double) As Double
'Source: McEwan P, Bennett H, Qin L, Bergenheim K, Gordon J and Evans M. An alternative approach to modelling HbA1c trajectories in patients with type 2 diabetes mellitus. Diabetes Obes Metab. 2017;19:628– 634. https://doi.org/10.1111/dom.12865

' Function to determine parametric HbA1c profile
    
    Dim hba1c_drop As Single
    Dim hba1c_minimum As Single
    Dim hba1c_out As Single
    Dim half_drop_rate As Single
    Dim Root1 As Single
    Dim root2 As Single
    Dim root3 As Single
    
      'adjust if drop is not realistic given the patient baseline
      'disabled to avoid masking other errors
      'If hba1c_baseline + treatment_effect < 1 Then treatment_effect = (hba1c_baseline - 4) * -1
      'If treatment_effect > 0 Then treatment_effect = 0
      
    hba1c_drop = hba1c_baseline + treatment_effect
    hba1c_minimum = kOvershoot * hba1c_drop
    
    If kHba1cLimit - hba1c_minimum < 0 Then
        'MsgBox "Parameters resulting in attempt to calculate square root of negative number." & vbNewLine & _
        '       "Either reduce the overshoot parameter, increase the magnitude of the treatment effect parameter, or increase the maximum HbA1c parameter."
        hba1c_minimum = kHba1cLimit - 2
        'Hba1cProfileParametric = 999
        'Exit Function
    End If
    
    Root1 = (kHba1cLimit - hba1c_minimum) ^ 0.5
    root2 = (hba1c_baseline - hba1c_minimum) ^ 0.5
    root3 = (hba1c_drop - hba1c_minimum) ^ 0.5
    half_drop_rate = Log(Root1 + root2)
    half_drop_rate = half_drop_rate - Log(Root1 + root3)
    hba1c_out = Exp(-half_drop_rate * time_since_baseline)
    hba1c_out = hba1c_out * (Root1 + root2)
    hba1c_out = hba1c_out - Root1
    hba1c_out = hba1c_out * hba1c_out
    hba1c_out = hba1c_out + hba1c_minimum
    
    Hba1cProfileParametric = hba1c_out

End Function


Function Hba1cReverse(hba1c_out As Single, time_since_baseline As Single, _
                      treatment_effect As Single, kOvershoot As Single, kHba1cLimit As Double) As Double
    
    Dim hba1c_baseline As Single
    Dim hba1c_minimum As Single
    Dim half_drop_rate As Single
    Dim Root1 As Single, root2 As Single, root3 As Single
    Dim delta As Double
    Dim max_iter As Integer
    Dim iter As Integer

    ' Initial guess for hba1c_baseline
    hba1c_baseline = hba1c_out - treatment_effect

    ' Iterative solver parameters
    delta = 1
    max_iter = 100
    iter = 0

    Do While Abs(delta) > 0.01 And iter < max_iter
        
        Dim hba1c_pred As Single

        hba1c_pred = Hba1cProfileParametric(time_since_baseline, hba1c_baseline, treatment_effect, kOvershoot, kHba1cLimit)

        ' Update hba1c_baseline using Newton-Raphson
        delta = hba1c_out - hba1c_pred
        hba1c_baseline = hba1c_baseline + delta / 10 ' Adjust step size for stability

        iter = iter + 1
    Loop


Hba1cReverse = hba1c_baseline

    
End Function
