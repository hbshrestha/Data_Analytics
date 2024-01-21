Attribute VB_Name = "dynamic_array"
Option Base 1

Sub array_dynamic()

Dim wb As Workbook
Dim ws2 As Worksheet
Set wb = ThisWorkbook
Set ws2 = wb.Worksheets("Sheet2")


ReDim countries_visited(4)
ReDim population(4)

countries_visited(1) = "France"
population(1) = 68

countries_visited(2) = "Spain"
population(2) = 48

countries_visited(3) = "Iran"
population(3) = 88

countries_visited(4) = "Indonesia"
population(4) = 274

ws2.Range("A1").Value = "Countries visited"
ws2.Range("B1").Value = "Population (million)"

ReDim Preserve countries_visited(5)
ReDim Preserve population(5)

countries_visited(5) = "Portugal"
population(5) = 10

Dim size As Integer
size = UBound(countries_visited) - LBound(countries_visited)
Debug.Print LBound(countries_visited)


Dim i As Integer
For i = 2 To 6
    Range("A" & i).Value = countries_visited(i - 1)
    Range("B" & i).Value = population(i - 1)

Next i



End Sub
