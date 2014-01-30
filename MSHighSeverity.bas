Attribute VB_Name = "Module1"
Sub AllVulnerabilityReportMS()
Cells.Select
Selection.RowHeight = 15
Range("A1").Select

ActiveSheet.Copy After:=Sheets(Sheets.Count)
ActiveSheet.Name = "5"
'===============================================
'Delete the first four rows in the sheet
'===============================================
Rows("1:7").Select
Range("A7").Activate
Selection.Delete Shift:=x1Up

'===============================================
'Sorts sheet by Severity, IP, QID
'===============================================
Dim rowCount As String 'Number of rows in sheet
Dim endRow As String
Dim firstColumn As String
Dim SortAll As String
Dim SortIP As String
Dim SortQID As String
Dim filterString As String
Dim filterCount As String

'Setup Variables
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select 'Select entire column
rowCount = Selection.Rows.Count                'Number of Rows in column
firstColumn = "A" & rowCount                   'Select column "A" Vulnerabilities only
endRow = "A2:AH" & rowCount                    'Select column
SortAll = "A1:AH" & rowCount                   'Select All Vuln Rows and Coln
SortSeverity = "K2:K" & rowCount               'Selecr Coln for sort by severity
SortIP = "A2:A" & rowCount                     'Select Coln for sort by IP
SortQID = "G2:G" & rowCount                    'Select Coln for sort by QID
Range(endRow).Select



'===============================================
'Start Sorting
'===============================================
'Reset cursor to first cell
Range("A1").Select

ActiveWorkbook.Worksheets("5").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("5").Sort.SortFields.Add Key:=Range(SortSeverity), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("5").Sort.SortFields.Add Key:=Range(SortIP), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("5").Sort.SortFields.Add Key:=Range(SortQID), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("5").Sort
    .SetRange Range(SortAll)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'Select all Cells and format row height to 15
Cells.Select
Selection.RowHeight = 15
Range("A1").Select

'===============================================
'Copy header to sheet 4,3
'===============================================
'Create sheets for 4,3  2,1 and "No Results For Host"
Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "4"
Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "3"
Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "No Results For Host"

'Copy header
Sheets("5").Select
Rows("1:1").Select
Selection.Copy

'Paste header to sheets created above
Sheets("4").Select
ActiveSheet.Paste
Sheets("3").Select
ActiveSheet.Paste


'Go back to Sheet 5
Sheets("5").Select
'deselect copied cells
Application.CutCopyMode = False

'===============================================
'Filtering and Copying to appropriate tabs
'===============================================
Range("A1").Select
ActiveWorkbook.Worksheets("5").Sort.SortFields. _
    Clear
ActiveWorkbook.Worksheets("5").Sort.SortFields. _
    Add Key:=Range(SortSeverity), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("5").Sort.SortFields. _
    Add Key:=Range(SortIP), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("5").Sort.SortFields. _
    Add Key:=Range(SortQID), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("5").Sort
    .SetRange Range(SortAll)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Format cell height
Cells.Select
Selection.RowHeight = 15

'===============================================
'filter No Results and No Vulnerabilities
'===============================================
Dim foundCell As String
Dim startRange As String
Dim newRange As String
Dim cell As Object

'Cut and paste "No results available for these hosts
Set cell = Cells.Find(What:="No results available for these hosts", After:=ActiveCell _
    , LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
If cell Is Nothing Then
    
Else
    cell.Select
    foundCell = ActiveCell.Address
    startRange = ActiveCell.Offset(0, -5).Address
    newRange = startRange & ":" & foundCell
    Range(newRange).Select
    Selection.Cut
    
    'Paste selection
    Sheets("No Results For Host").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.RowHeight = 15
    
    Sheets("5").Select
    Selection.EntireRow.Delete
End If

Range("A1").Select

'Cut and paste "No vulnerabilities match your filters for these hosts
Set cell = Cells.Find(What:="No vulnerabilities match your filters for these hosts", After:=ActiveCell _
    , LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
If cell Is Nothing Then
    
Else
    cell.Select
    foundCell = ActiveCell.Address
    startRange = ActiveCell.Offset(0, -5).Address
    newRange = startRange & ":" & foundCell
    Range(newRange).Select
    Selection.Cut
    
    'Paste selection
    Sheets("No Results For Host").Select
    Range("A2").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.RowHeight = 15
    
    Sheets("5").Select
    Selection.EntireRow.Delete
End If

'===============================================
'filter 3's
'===============================================
Range("A1").Select
Selection.AutoFilter
filterString = "$A$1:$AH$" & rowCount
ActiveSheet.Range(filterString).AutoFilter Field:=11, Criteria1:="3"

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
'Count rows in new filter selection
filterCount = Selection.Rows.Count + 1               'Number of Rows in column
filterString = "A2:AH" & filterCount
Range(filterString).Select

'Paste selection
Selection.Cut
Sheets("3").Select
Range("A2").Select
ActiveSheet.Paste
Cells.Select
Selection.RowHeight = 15
Range("A1").Select

Sheets("5").Select
Selection.EntireRow.Delete

'===============================================
'filter 4's
'===============================================
Range("A1").Select
Selection.AutoFilter
filterString = "$A$1:$AH$" & rowCount
ActiveSheet.Range(filterString).AutoFilter Field:=11, Criteria1:="4"

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
'Count rows in new filter selection
filterCount = Selection.Rows.Count + 1               'Number of Rows in column
filterString = "A2:AH" & filterCount
Range(filterString).Select

'Paste selection
Selection.Cut
Sheets("4").Select
Range("A2").Select
ActiveSheet.Paste
Cells.Select
Selection.RowHeight = 15
Range("A1").Select

Sheets("5").Select
Selection.EntireRow.Delete


'===============================================
'filter 5's
'===============================================
Range("A1").Select
Selection.AutoFilter
filterString = "$A$1:$AH$" & rowCount
ActiveSheet.Range(filterString).AutoFilter Field:=11, Criteria1:="5"
Selection.AutoFilter

End Sub



