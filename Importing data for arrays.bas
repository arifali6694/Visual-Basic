Sub csv_formatting()
'
' csv_formatting Macro
' loading in the semicolon csv
'
Dim Import_csv As String
'
Import_csv = Application.GetOpenFilename()
Cells(1, 1).Value = Import_csv + " = import_csv"
Import_csv = "TEXT;" & Import_csv

'    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\Arif\Desktop\File1.csv", Destination:= _
        Range("$A$1"))'
     With ActiveSheet.QueryTables.Add(Connection:= _
         Import_csv, Destination:= _
         Range("$A$1"))
        .Name = "File1"
        .FieldNames = True=
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
Dim i As Integer, j As Integer, numofcol As Integer, numofrow As Integer
Dim filepath As String, arraypos As Long

Public array1() As String

'Private Sub copy_table()
'For i = 1 To 11
'    For j = 1 To 2
'        table1(i, j) = Cells(i, j).Value
'    Next j
'Next i
'
'For j = 1 To 2
'    For i = 1 To 11
'        Cells(i, j + 2) = table1(i, j)
'    Next i
'Next j
'
'i = 1
'Do While Cells(i, 1).Value <> ""
'    Cells(i, 3).Value = Cells(i, 1).Value
'    Cells(i, 4).Value = Cells(i, 1).Value
'    i = i + 1
'    Loop



Sub load_result_table()
filepath = Application.GetOpenFilename()
Open filepath For Input As #1
row_number = 0
Do Until EOF(1)
    Line Input #1, linefromfile
        lineitems = Split(linefromfile, "]:")
    
        row_number = row_number + 1
    numofrow = row_number
    numofcol = UBound(lineitems())
'numofrow = WorksheetFunction.CountA(Application.Workbooks("authors.csv").Worksheets("authors").Columns(1))
'numofcol = WorksheetFunction.CountA(Application.Workbooks("authors.csv").Worksheets("authors").Rows(1))
ReDim array1(numofrow - 1, numofcol - 1)
For i = 1 To numofrow
    For j = 1 To numofcol
        array1(i - 1, j - 1) = lineitems(j - 1)
     Next j
Next i
Loop
Close #1
'For arraypos = 0 To UBound(array1, 1)
'    MsgBox array1(arraypos, 0)
'Next arraypos
'Open filepath For Input As #1
'row_number = 0
'Do Until EOF(1)
'   Line Input #1, linefromfile
'    lineitems = Split(linefromfile, ",")
'    ActiveCell.Offset(row_number, 0).Value = lineitems(2)
'    ActiveCell.Offset(row_number, 1).Value = lineitems(1)
'    ActiveCell.Offset(row_number, 2).Value = lineitems(0)
'    i = 0
'    Do While lineitems(i) <> ""
'        ActiveCell.Offset(1, i).Value = lineitems(i)
'        ActiveCell.Offset(1, i).Value = lineitems(i)
'        ActiveCell.Offset(1, i).Value = lineitems(i)
'        i = i + 1
'    Loop
'
'Loop
'Close #1
'End Sub
'
'Sub Print_Array()
End Sub
