Attribute VB_Name = "GetHourlyAverageAdvanced"
Option Base 1
Sub GetHourlyAverage()
    
'defining an array for 4 strings
Dim columns(4) As String
columns(1) = "B"
columns(2) = "C"
columns(3) = "D"
columns(4) = "E"

'Definining mnth because Month is a function
Dim mnth As Integer
Dim row As Integer
    
Dim sum As Double
Dim num_hours As Double
        
Dim wb As Workbook
Dim ws As Worksheet
Dim datetime As Range
Dim last_row As Integer

Set wb = ThisWorkbook
Set ws = ThisWorkbook.Sheets("Sheet1")
Set datetime = ws.Range("datetime")
last_row = Cells(datetime.row, datetime.column).End(xlDown).row

Debug.Print datetime.Value
Debug.Print "Row: " & datetime.row & " Column: " & datetime.column
Debug.Print "Last row: " & last_row
       
        
'Use Option Base 1 before this subroutine if you want to start the list from 1 instead of 0.
'https://excelchamps.com/vba/arrays/

'Loop through column for each city
For Each column In columns
    
    'Loop through each hour of the day
    For hr = 0 To 23
        
        sum = 0
        num_hours = 0
        
        'Loop through each row
        For row = datetime.row + 1 To last_row
        
            If Hour(Cells(row, datetime.column).Value) = hr Then
                Range(column & row).Interior.Color = RGB(255, 255, 0)
                num_hours = num_hours + 1
                sum = sum + Range(column & row).Value
            
            End If
        
            
        Next row
               
        Range(column & hr + 2).Offset(0, 14).Value = sum / num_hours
          
    Next hr

Next column

End Sub




