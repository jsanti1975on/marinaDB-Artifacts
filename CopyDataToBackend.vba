Sub CopyDataToBackend()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim copyRange As Range

    ' Set the source and target sheets
    Set sourceSheet = ThisWorkbook.Sheets("ParsedData")
    Set targetSheet = ThisWorkbook.Sheets("compare")
    
    ' Define the range to copy
    Set copyRange = sourceSheet.Range("A1:B81")
    
    ' Copy the values from Sheet1 to Backend
    targetSheet.Range("A1:B81").Value = copyRange.Value
    
    MsgBox "Data copied successfully from Sheet1 to Backend.", vbInformation
End Sub
