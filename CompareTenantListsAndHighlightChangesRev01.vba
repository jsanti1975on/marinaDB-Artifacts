Sub CompareTenantListsAndHighlightChangesRev01()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentTenant As String
    Dim previousTenant As String
    Dim highlightedSlips As String
    
    ' Set the worksheet where the tenant data is
    Set ws = ThisWorkbook.Sheets("compare")
    
    ' Find the last row in Column A (current tenant list)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the highlightedSlips variable
    highlightedSlips = ""
    
    ' Loop through the current tenant list and compare with previous tenants in Column G
    For i = 2 To lastRow ' Start from row 2 to skip headers
        currentTenant = Trim(ws.Cells(i, 1).Value) ' Current tenant in Column A
        previousTenant = Trim(ws.Cells(i, 7).Value) ' Previous tenant in Column G
        
        ' Compare the current tenant with the previous tenant
        If currentTenant <> previousTenant Then
            ' If there is a difference, highlight the cell in Column A (current tenant)
            ws.Cells(i, 1).Interior.Color = RGB(255, 0, 0) ' Highlight with red
            
            ' Add the slip number (row - 1 for header offset) to the highlightedSlips list
            highlightedSlips = highlightedSlips & (i - 1) & ", "
        End If
    Next i
    
    ' Remove the trailing comma and space if any slips were highlighted
    If Len(highlightedSlips) > 0 Then
        highlightedSlips = Left(highlightedSlips, Len(highlightedSlips) - 2)
    Else
        highlightedSlips = "None" ' No changes found
    End If
    
    ' Display the highlighted slip numbers in a message box
    MsgBox "Comparison completed. Changes found in the following slips: " & highlightedSlips, vbInformation
End Sub
