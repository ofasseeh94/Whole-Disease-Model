Attribute VB_Name = "GeneralMortality"

Public Function General_Mortality(Age As Integer, Gender_Female As Boolean, Optional Country As String = "OM", Optional Model_Year As Integer) As Double
'Country provided as an ISO3 code
'The function will provide the annual mortality rate based on age, year of birth and country
Dim TempMatrix As Variant

If Model_Year > 2100 Then Model_Year = 2100

General_Mortality = General_Mortality_Matrix(Abs(Gender_Female) * 7777 + (Model_Year - 2024) * 101 + Age + 1, 6)

End Function

