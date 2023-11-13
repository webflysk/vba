
   
Sub FindMatchingRegistrationNumbers()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim rng1 As Range, rng2 As Range
    Dim cell As Range, foundCell As Range

    ' Set references to the sheets
    Set ws1 = ThisWorkbook.Sheets("Slovenske")
    Set ws2 = ThisWorkbook.Sheets("Finance")

    ' Assuming the Registration Nr. is in Column A in both sheets
    Set rng1 = ws1.Range("B1:B" & ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row)
    Set rng2 = ws2.Range("C1:C" & ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row)

    ' Loop through each cell in rng1
    For Each cell In rng1
        ' Check if the cell value is present in rng2
        Set foundCell = rng2.Find(What:=cell.Value, LookIn:=xlValues, LookAt:=xlWhole)

        ' If found, do something with the row where the match is found
        If Not foundCell Is Nothing Then
            ' For example, highlight the found row in yellow
            foundCell.EntireRow.Interior.Color = vbYellow
            ' You can also perform other actions here
        End If
    Next cell
End Sub

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
