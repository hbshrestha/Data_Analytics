Attribute VB_Name = "GetMonthlyAverage"
Option Base 1
Sub GetMonthlyAverage()
    
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
        
Dim datetime As Range
Set datetime = Range("datetime")
Debug.Print datetime.row, datetime.column

        
'Use Option Base 1 before this subroutine if you want to start the list from 1 instead of 0.
'https://excelchamps.com/vba/arrays/

'Loop through column for each city
For Each column In columns
    
    'Loop through each month of the year
    For mnth = 1 To 12
        
        sum = 0
        num_hours = 0
        
        'Loop through each row
        For row = 2 To 8785
        
            If month(Cells(row, datetime.column)) = mnth Then
                Range(column & row).Interior.Color = RGB(255, 255, 0)
                num_hours = num_hours + 1
                sum = sum + Range(column & row).Value
            
            End If
        
            
        Next row
               
        Range(column & mnth).Offset(1, 7).Value = sum / num_hours
          
    Next mnth

Next column

End Sub


