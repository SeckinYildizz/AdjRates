Sub NJ_AdjRates()
'The macro utilize the visit maintenance report and creates a part of the adjustment report from it.
'The number of adjustments must be filled manually.
    
    'Set the variables
    Dim i, lastRow As Integer
    Dim cell As Range
    Dim targetSheet As Worksheet
    
    'Create a new sheet and copy the DSP column to it
    Sheets.Add after:=Sheets(Sheets.Count)
    Set targetSheet = Sheets(Sheets.Count)
    Sheets(1).Range("l:l").Copy Destination:=targetSheet.Range("a:a")
    
    'Remove the duplicate names and apply countif formula
    With targetSheet
        .Range("a:a").RemoveDuplicates Columns:=1, Header:=xlYes
        .Rows("1:2").EntireRow.Delete
        .Range("b1:d1").Value = Array("DSP Names", "Total Visits", "Adjustment Numbers")
        lastRow = .UsedRange.Rows.Count
        For i = 2 To lastRow
            .Cells(i, "B").Formula = Left(.Cells(i, "A").Value, InStrRev(.Cells(i, "A"), " ") - 1)
            .Cells(i, "C").Formula = WorksheetFunction.CountIf(Sheets(1).Range("L:L"), .Cells(i, "A").Value)
            .UsedRange.Sort Key1:=Range("b1"), order1:=xlAscending, Header:=xlYes
        Next i
    End With

End Sub
