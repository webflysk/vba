Sub Nakup()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, lastRowDest As Long
    Dim currentRow As Long
    Dim dateThreshold As Date

    ' Set the date threshold to 10 days ago
    dateThreshold = Date

    ' Set references to the source sheet
    Set wsSource = ThisWorkbook.Sheets("Slovenske")

    ' Create a new sheet named "kupit" if it doesn't exist
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("kupit")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "kupit"
    End If
    On Error GoTo 0

    ' Find the last row in the "Slovenske" sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row

    ' Find the last row in the "kupit" sheet
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1

    Application.ScreenUpdating = False ' Disable screen updating for performance

    ' Loop through each row in the "Slovenske" sheet
    For currentRow = 2 To lastRowSource ' Assuming data starts at row 2
        ' Check if the date in column G is within the last 10 days and column L is "SK"
        If wsSource.Cells(currentRow, "G").Value >= dateThreshold And wsSource.Cells(currentRow, "L").Value = "SK" Then
            ' If there's a match, copy columns A, B, and C to the "kupit" sheet
            wsSource.Cells(currentRow, "A").Copy Destination:=wsDest.Cells(lastRowDest, "A")
            wsSource.Cells(currentRow, "B").Copy Destination:=wsDest.Cells(lastRowDest, "B")
            wsSource.Cells(currentRow, "C").Copy Destination:=wsDest.Cells(lastRowDest, "C")
            lastRowDest = lastRowDest + 1 ' Increment to next available row
        End If
    Next currentRow

    Application.ScreenUpdating = True ' Re-enable screen updating

    MsgBox "Rows from the last 10 days with 'SK' in column L have been copied to 'kupit' sheet."
End Sub

Sub RemoveYellowRows()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim cell As Range
    Dim yellowColor As Long
    Dim isYellow As Boolean

    ' Define the RGB color for yellow
    yellowColor = RGB(255, 255, 0) ' This is the RGB value for yellow

    ' Set a reference to the "Finance" sheet
    Set ws = ThisWorkbook.Sheets("Finance")

    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False ' Turn off screen updating to speed up the macro
    Application.Calculation = xlCalculationManual ' Turn off automatic calculations

    ' Loop from lastRow to the first row (reverse loop to avoid skipping rows when deleting)
    For i = lastRow To 1 Step -1
        isYellow = False
        ' Check each cell in the row to see if it's yellow
        For Each cell In ws.Rows(i).Cells
            If cell.Interior.Color = yellowColor Then
                isYellow = True
                Exit For
            End If
        Next cell

        ' If any cell in the row is yellow, delete the entire row
        If isYellow Then
            On Error Resume Next ' Ignore errors and continue
            ws.Rows(i).Delete
            On Error GoTo 0 ' Stop ignoring errors
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic ' Turn automatic calculations back on
    Application.ScreenUpdating = True ' Turn screen updating back on

    MsgBox "Yellow rows have been removed."
End Sub

