
Sub MarkSettledContracts()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range
    Dim filterColumn As Integer
    Dim cell As Range

    ' Define which worksheet to use
    Set ws = ThisWorkbook.Sheets("Finance") ' Change "YourSheetName" to the name of your worksheet

    ' Find the last row with data on the sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define the range to apply the filter to, including the header row
    Set filterRange = ws.Range("A1:Z" & lastRow) ' Change "A1:Z" to match the range of your actual data

    ' Determine the column number where "Running - Dehired" is the header.
    ' It's important that the header name matches exactly.
    For i = 1 To filterRange.Columns.Count
        If filterRange.Cells(1, i).Value = "Running - Dehired" Then
            filterColumn = i
            Exit For
        End If
    Next i

    ' Check if the filterColumn was found
    If filterColumn > 0 Then
        ' Apply the filter
        filterRange.AutoFilter Field:=filterColumn, Criteria1:="Settled Contracts"

        ' Loop through each cell in the filterColumn and mark if visible (not filtered out)
        For Each cell In ws.Range(ws.Cells(2, filterColumn), ws.Cells(lastRow, filterColumn))
            If cell.EntireRow.Hidden = False Then
                ' Mark the cell or row as needed
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow background for the filtered cell
                ' If you want to mark the entire row, use:
                ' cell.EntireRow.Interior.Color = RGB(255, 255, 0)
            End If
        Next cell

        ' If you want to clear the filter after marking the rows, uncomment the next line
        ' ws.ShowAllData
    Else
        MsgBox "Column 'Running - Dehired' not found."
    End If
End Sub
