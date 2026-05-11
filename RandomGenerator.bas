Attribute VB_Name = "RandomGenerator"
Option Explicit

Function GenerateRandomArray(NPatients As Long, NCycles As Long, NParameters As Long, Optional Seed As Integer) As Single()

'redim the array of random numbers based on the number of patients and the cycle length
Dim TempRandArr() As Single
ReDim TempRandArr(NPatients, NCycles, NParameters)

'set a random seed
      If IsEmpty(Seed) Then Seed = 42
      Rnd (-1)
      Randomize Seed

Dim i As Long
Dim x As Long
Dim j As Long

'loop through the array to assign random numbers and consume as much memory as we can
For i = 0 To NPatients

      For x = 0 To NCycles
      
            For j = 0 To NParameters
            
                 TempRandArr(i, x, j) = Rnd
            
            Next j
            
      Next x

Next i

GenerateRandomArray = TempRandArr

End Function
