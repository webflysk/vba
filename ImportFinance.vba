Sub ImportFinance()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim FilePath As String
    Dim ws As Worksheet
    Dim SheetExists As Boolean
    
    ' Define the file path of the external workbook
    FilePath = "C:\\0x\\Finance - Details at contract level.xlsx"
    
    ' Check if the external workbook is open
    On Error Resume Next
    Set SourceWorkbook = Workbooks("Finance - Details at contract level.xlsx")
    On Error GoTo 0
    
    ' If the external workbook is not open, open it
    If SourceWorkbook Is Nothing Then
        Set SourceWorkbook = Workbooks.Open(FilePath)
    End If
    
    ' Set the destination workbook (the current workbook)
    Set DestinationWorkbook = ThisWorkbook
    
    ' Check if "Finance" sheet already exists
    SheetExists = False
    For Each ws In DestinationWorkbook.Sheets
        If ws.Name = "Finance" Then
            SheetExists = True
            Exit For
        End If
    Next ws
    
    ' If "Finance" sheet exists, delete it
    If SheetExists Then
        Application.DisplayAlerts = False
        DestinationWorkbook.Sheets("Finance").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Set the source sheet (sheet to be imported)
    Set SourceSheet = SourceWorkbook.Sheets(1) ' Assuming you want to import the first sheet
    
    ' Copy the source sheet to the destination workbook after the last sheet
    SourceSheet.Copy After:=DestinationWorkbook.Sheets(DestinationWorkbook.Sheets.Count)
    
    ' Rename the newly added sheet to "Finance"
    Set NewSheet = DestinationWorkbook.Sheets(DestinationWorkbook.Sheets.Count)
    NewSheet.Name = "Finance"
    
    ' Close the external workbook without saving changes (if it was opened)
    If SourceWorkbook.Name <> DestinationWorkbook.Name Then
        SourceWorkbook.Close SaveChanges:=False
    End If
    
    ' Define the table range (assuming your data starts in cell A1)
    Dim TableRange As Range
    Set TableRange = NewSheet.UsedRange
    
    ' Add a table with the name "FinanceTable" and use the first row as column names
    DestinationWorkbook.Sheets("Finance").ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = "FinanceTable"
    
    ' Loop through the table columns
    Dim tbl As ListObject
    Dim col As ListColumn
    Set tbl = DestinationWorkbook.Sheets("Finance").ListObjects("FinanceTable")
    
    For Each col In tbl.ListColumns
        ' Remove double quotes from column name
        col.Name = Replace(col.Name, Chr(34), "")
        ' Set text color to automatic (black)
        col.DataBodyRange.Font.Color = -16777216 ' Automatic color (RGB 0,0,0)
    Next col
    
    ' Apply a blue table design style
    tbl.TableStyle = "TableStyleLight21" ' You can change this to other blue styles as needed
End Sub

