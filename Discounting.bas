Attribute VB_Name = "Discounting"
Function DiscountedValue(Value As Double, DiscountRate As Double, TimeFrom As Double, TimeTo As Double, _
Optional TimeUnit As Double, Optional OneTimeEvent As Boolean) As Double
'timeunit in relation to years
'month should be 12
'days should be 365.25
'6month should be 2

Dim IntervalCount As Integer
Dim Interval As Integer
Dim FirstCutOff As Double

Dim IntervalTime As Double
Dim IntervalValue As Double
Dim Time As Double

'Time = TimeFrom

If IsEmpty(TimeUnit) Or TimeUnit = 0 Then TimeUnit = 1

    IntervalCount = Application.WorksheetFunction.RoundUp((TimeTo) / TimeUnit, 0) - Application.WorksheetFunction.RoundDown((TimeFrom) / TimeUnit, 0)
    
    If TimeTo / TimeUnit = Int(TimeTo / TimeUnit) And TimeFrom / TimeUnit = Int(TimeFrom / TimeUnit) Then IntervalCount = IntervalCount - 1
    
    IntervalCount = WorksheetFunction.Max(IntervalCount, 1)
    
      FirstCutOff = Int((TimeFrom + TimeUnit) / TimeUnit) * TimeUnit
      If FirstCutOff > TimeTo Then FirstCutOff = TimeTo
      Time = FirstCutOff
            Rem add Condition for one time event
            If Not OneTimeEvent Then
            For Interval = 1 To IntervalCount
            
                  IntervalTime = Application.WorksheetFunction.RoundUp(((TimeFrom + (Interval - 1) * TimeUnit) / TimeUnit), 0)
                  If TimeFrom = 0 Then IntervalTime = 1
                  
                  If Interval = 1 Then
                        IntervalValue = ((FirstCutOff - TimeFrom) / (TimeTo - TimeFrom)) * Value
                  
                  ElseIf Interval = IntervalCount Then
                  
                        IntervalValue = ((Time - TimeTo) / (TimeTo - TimeFrom)) * Value
                  
                  Else
                        
                        IntervalValue = ((TimeUnit) / (TimeTo - TimeFrom)) * Value
                  
                  End If
                  
                  DiscountedValue = DiscountedValue + IntervalValue / ((1 + DiscountRate) ^ (Time))
                  Time = Time + 1
            
            Next Interval
      
      Else
            
            Rem Calculate Discounted value at the begining of the period
            DiscountedValue = Value / ((1 + DiscountRate) ^ (Time))
            
      End If

End Function


