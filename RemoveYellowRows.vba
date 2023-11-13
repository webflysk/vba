
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

