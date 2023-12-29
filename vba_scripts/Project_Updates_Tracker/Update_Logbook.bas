Attribute VB_Name = "Update_Logbook"
Sub Update_logbook()

Dim wb As Workbook
Dim ws1 As Worksheet 'Project tracker worksheet
Dim ws2 As Worksheet 'Logbook worksheet

Set wb = ThisWorkbook
Set ws1 = ThisWorkbook.Sheets("ProjectTasksTracker")
Set ws2 = ThisWorkbook.Sheets("Logbook")


Dim lr1, lc1 As Integer
lr1 = ws1.Cells(Rows.Count, "A").End(xlUp).Row 'Count the number of rows in the ProjectTasksTracker sheet
lc1 = ws1.Cells(1, Columns.Count).End(xlToLeft).Column 'Count number of columns in the ProjectTasksTrackersheet


Dim i, j, lr2 As Integer
Dim rg1, rg2 As Range
Dim k As Long

'Loop through each row of ProjectTasksTracker sheet except the header row
For i = 2 To lr1
    
    'Check if the update column is not empty in ProjectTasksTracker sheet.
    'Proceed if not empty
    If ws1.Cells(i, 6).Value <> "" Then
        
        'Count the number of rows in Logbook sheet
        lr2 = ws2.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Create a boolean datatype named valuesMatch and assign to False as default
        Dim valuesMatch As Boolean
        valuesMatch = False
        
        'Loop through each row in Logbook sheet except header row
        For j = 2 To lr2
            
            'Set rg1 as a row in ProjectTasksTracker sheet ("B"&i:"F"&i)
            Set rg1 = ws1.Range("B" & i & ":" & "F" & i)
            'Set rg2 as a row in Logbook sheet ("B"&j:"F"&j)
            Set rg2 = ws2.Range("B" & j & ":" & "F" & j)
            
            'Loop through each column in range rg1
            For k = 1 To rg1.Count
            
                'Check if value in each cell for given range rg1 matches with value in corresponding cell in range rg2.
                If rg1(k).Value <> rg2(k).Value Then
                    
                    'If there are no matches, valuesMatch remains False by default and the loop is exited.
                    valuesMatch = False
                    Exit For
                    
                Else
                    'If there are matches, this means that the update in ProjectTasksTracker sheet already exists in Logbook sheet.
                    'valuesMatch is set to True in this case.
                    valuesMatch = True
                End If
            
            Next k
            
            'Exit for loop if the duplicate exists.
            If valuesMatch = True Then
                Exit For
            End If

        Next j
        
        'If the update in ProjectTasksTracker doesn't exist in Logbook sheet, then
        If valuesMatch = False Then
            'Copy the update from ProjectTasksTracker sheet
            ws1.Range("A" & i & ":F" & i).Copy
            'Go to the end of the Logbook sheet and paste it
            ws2.Range("A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row + 1).PasteSpecial xlPasteValuesAndNumberFormats
        End If
    
    End If
    
Next i

End Sub

