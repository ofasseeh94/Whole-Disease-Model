Attribute VB_Name = "MEA_Module"
Option Explicit

'Public Function AvgRebate(NPatients As Double, RebateRange As Range) As Double
'    '===============================
'    ' Computes a weighted average rebate rate
'    ' based on “tiers” of patient counts and rebate percentages.
'    ' - NPatients: total number of patients (can be fractional, but normally integer)
'    ' - RebateRange: an Excel Range (e.g., A1:B5) where:
'    '       Column1 = “tier capacity” (number of patients for that tier)
'    '       Column2 = rebate rate (e.g., 0.05 for 5 %)
'    ' Returns a Double = the average rebate (as a decimal, e.g., 0.03 for 3 %).
'    '===============================
'
'    Dim RebateMatrix As Variant
'    ' RebateMatrix will hold the contents of RebateRange as a 2D Variant array.
'    ' Reading from an array is faster than repeatedly reading from the Range object.
'    RebateMatrix = RebateRange
'
'    Dim i As Byte
'    Dim RebateTiers As Byte
'    Dim RemainingPatients As Double
'    Dim AvgDiscount As Double
'    '   i                 ? loop counter for each tier (1-based index)
'    '   RebateTiers       ? number of rows (tiers) in RebateMatrix
'    '   RemainingPatients ? how many patients still need to be “assigned” to a tier
'    '   AvgDiscount       ? holds the running total of (patients × rebate rate)
'    '                       before we divide by NPatients to get an average.
'
'    ' Determine how many tiers there are by finding the upper bound of the first array dimension.
'    ' Since RebateMatrix was created from a Range (m rows × 2 columns), UBound(RebateMatrix) = m.
'    RebateTiers = UBound(RebateMatrix)
'
'    ' Initialize RemainingPatients to the full count of patients we need to process.
'    RemainingPatients = NPatients
'
'    ' Loop through each tier from 1 to RebateTiers
'    For i = 1 To RebateTiers
'
'        ' ------------------------
'        ' Read this tier’s “capacity” and “rate” from the array:
'        '   RebateMatrix(i, 1) = number of patients that tier i can “cover”
'        '   RebateMatrix(i, 2) = rebate rate for tier i (e.g., 0.04 means 4 %)
'        ' ------------------------
'        Dim TierCapacity As Double, TierExtraAmount As Double
'        Dim TierRate As Double
'
'        TierCapacity = RebateMatrix(i, 1)
'
'        If i <> 1 Then
'
'            TierExtraAmount = RebateMatrix(i, 1) - RebateMatrix(i - 1, 1)
'
'        Else
'
'            TierExtraAmount = TierCapacity
'
'        End If
'
'        TierRate = RebateMatrix(i, 2)
'
'        ' If there are more patients remaining than this tier can handle,
'        ' we assign the full TierCapacity at TierRate.
'        If RemainingPatients > TierExtraAmount Then
'
'            ' Add (TierCapacity × TierRate) to our running total
'            AvgDiscount = AvgDiscount + (TierExtraAmount * TierRate)
'
'            ' Subtract TierCapacity from RemainingPatients,
'            ' because we’ve “used up” that many patients in this tier.
'            RemainingPatients = RemainingPatients - TierExtraAmount
'
'            ' Then move on to the next tier (i = i + 1)
'            ' (Loop continues automatically.)
'
'        Else
'            ' ------------------------
'            ' Otherwise, if RemainingPatients = TierCapacity:
'            ' We only need to apply this tier’s discount to all the leftover patients.
'            ' ------------------------
'
'            ' Add (RemainingPatients × TierRate) to the running total
'            AvgDiscount = AvgDiscount + (RemainingPatients * TierRate)
'
'            ' Now all patients have been “assigned,” so we exit the loop.
'            GoTo JumpOut
'        End If
'
'        ' After subtracting TierCapacity (above), the loop goes to the next i automatically.
'    Next i
'
'JumpOut:
'    ' ===============================
'    ' At this point, AvgDiscount holds the sum of (patients_in_each_tier × tier_rate)
'    ' To get the average rebate rate per patient, divide by the original NPatients.
'    ' ===============================
'    AvgDiscount = AvgDiscount / NPatients
'
'    ' Debug.Print is only for the Immediate Window during development/testing.
'    ' You can remove it if you don’t need to trace execution.
'    Debug.Print "test"
'
'    ' Return the final average rebate to whoever called AvgRebate(...)
'    AvgRebate = AvgDiscount
'
'End Function


'Sub CalculateMEAOnce()
'
'Sheets("MEA").EnableCalculation = True
'
'
'End Sub
