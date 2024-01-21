Attribute VB_Name = "array_module"
'Public countries(1 To 4) As String

Sub array_1d()

countries(1) = "Nepal"
countries(2) = "India"
countries(3) = "Germany"
countries(4) = "Netherlands"

Dim i As Integer
Range("A1").Value = "Country"

For i = 1 To 4

    Range("A" & i + 1).Value = countries(i)

Next i


End Sub


Sub array_2d()

Dim country_capital(4, 2) As String


For i = 1 To 4
    country_capital(i, 1) = countries(i)
Next i

country_capital(1, 2) = "Kathmandu"
country_capital(2, 2) = "New Delhi"
country_capital(3, 2) = "Berlin"
country_capital(4, 2) = "Amsterdam"

Range("B1").Value = "Capital"

For i = 1 To 4
    Range("A" & i + 1).Value = country_capital(i, 1)
    Range("B" & i + 1).Value = country_capital(i, 2)
    
Next i


End Sub

