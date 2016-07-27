Public saveToPath As String
Public monthStr As String
Public yearStr As Integer
Public mainwkbk As Workbook



Private Sub AllCutButton_Click()
Application.ScreenUpdating = False

Dim intChoice As Long
Dim strPath As String

If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

monthStr = InputBox(Prompt:="Please input the month in CAPITAL LETTERS (MMM)", _
          Title:="Please input the month (MMM)")
If monthStr <> "JAN" And monthStr <> "FEB" And monthStr <> "MAR" _
And monthStr <> "APR" And monthStr <> "MAY" And monthStr <> "JUN" _
And monthStr <> "JUL" And monthStr <> "AUG" And monthStr <> "SEP" _
And monthStr <> "OCT" And monthStr <> "NOV" And monthStr <> "DEC" Then
MsgBox ("PLEASE INPUT A CORRECT MONTH!")
Exit Sub
End If

yearStr = InputBox(Prompt:="Please input the year (YYYY)", _
          Title:="Please input the year (YYYY)")
If yearStr > 9999 Or yearStr < 2000 Then
MsgBox ("PLEASE INPUT A CORRECT YEAR!")
Exit Sub
End If




Application.FileDialog(msoFileDialogOpen).Title = "Select the G&B file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

Set wkbk = Workbooks.Open(strPath)
Set mainwkbk = wkbk
Set homesheet = ActiveSheet
    CreatePivotTableForAll
    BMDrillDown
    getSPC
    homesheet.Activate
    CreatePivotTableForAll_BBD
    BBDDrillDown
    homesheet.Activate
    CreatePivotTableForAll_ZT
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox ("AUM Cutting 3 in 1 is Done, Please view files in indicated folder")
End Sub

Private Sub BMCutButton_Click()
Application.ScreenUpdating = False

Dim intChoice As Long
Dim strPath As String

If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

monthStr = InputBox(Prompt:="Please input the month in CAPITAL LETTERS (MMM)", _
          Title:="Please input the month (MMM)")
If monthStr <> "JAN" And monthStr <> "FEB" And monthStr <> "MAR" _
And monthStr <> "APR" And monthStr <> "MAY" And monthStr <> "JUN" _
And monthStr <> "JUL" And monthStr <> "AUG" And monthStr <> "SEP" _
And monthStr <> "OCT" And monthStr <> "NOV" And monthStr <> "DEC" Then
MsgBox ("PLEASE INPUT A CORRECT MONTH!")
Exit Sub
End If

yearStr = InputBox(Prompt:="Please input the year (YYYY)", _
          Title:="Please input the year (YYYY)")
If yearStr > 9999 Or yearStr < 2000 Then
MsgBox ("PLEASE INPUT A CORRECT YEAR!")
Exit Sub
End If






Application.FileDialog(msoFileDialogOpen).Title = "Select the G&B file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
Else: Exit Sub
End If

    Set wkbk = Workbooks.Open(strPath)
    Set mainwkbk = wkbk
    CreatePivotTableForAll
    BMDrillDown
    getSPC
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
Application.ScreenUpdating = True



MsgBox ("AUM Cutting for BM is Done, Please view files in indicated folder")
End Sub


Private Sub BBDCutButton_Click()
Application.ScreenUpdating = False


Dim intChoice As Long
Dim strPath As String


If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

monthStr = InputBox(Prompt:="Please input the month in CAPITAL LETTERS (MMM)", _
          Title:="Please input the month (MMM)")
If monthStr <> "JAN" And monthStr <> "FEB" And monthStr <> "MAR" _
And monthStr <> "APR" And monthStr <> "MAY" And monthStr <> "JUN" _
And monthStr <> "JUL" And monthStr <> "AUG" And monthStr <> "SEP" _
And monthStr <> "OCT" And monthStr <> "NOV" And monthStr <> "DEC" Then
MsgBox ("PLEASE INPUT A CORRECT MONTH!")
Exit Sub
End If

yearStr = InputBox(Prompt:="Please input the year (YYYY)", _
          Title:="Please input the year (YYYY)")
If yearStr > 9999 Or yearStr < 2000 Then
MsgBox ("PLEASE INPUT A CORRECT YEAR!")
Exit Sub
End If





Application.FileDialog(msoFileDialogOpen).Title = "Select the G&B file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
Else: Exit Sub
End If

    Set wkbk = Workbooks.Open(strPath)
    CreatePivotTableForAll_BBD
    BBDDrillDown
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox ("AUM Cutting for BBD is Done, Please view files in indicated folder")

End Sub


' This function is to create the pivot table for a BM
' input dataset is the dataset for entire BM (By double click the pivot table of AUM)
Sub CreatePivotTable()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim numOfRow As Long
numOfRow = ActiveSheet.UsedRange.Rows.Count



 SrcData = ActiveSheet.Name & "!" & Range("A1:BF" & numOfRow).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    
  'Add item to the Row Labels
    pvt.PivotFields("Mgr").Orientation = xlRowField
    pvt.PivotFields("Mgr").Position = 1
    pvt.PivotFields("RM").Orientation = xlRowField
    pvt.PivotFields("RM").Position = 2
    
  'Turn on Automatic updates/calculations --like screenupdating to speed up code
    pvt.ManualUpdate = False

    Dim values() As String
    ReDim values(5)
    values(1) = "TOT_AUM"
    values(2) = "AUM_INCENTIVE"
    values(3) = "custnum"
    values(4) = "LIF_INSURANCE_PEN"
    values(5) = "GEN_INSURANCE_PEN"
    
    Dim values_name() As String
    ReDim values_name(5)
    values_name(1) = "Sum of TOT_AUM"
    values_name(2) = "Sum of AUM_INCENTIVE"
    values_name(3) = "Count of custnum"
    values_name(4) = "Sum of LIF_INSURANCE_PEN"
    values_name(5) = "Sum of GEN_INSURANCE_PEN"
    
    Dim i As Long
    For i = 1 To 5
        If i = 3 Then
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlCount
        Else
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlSum
        End If
    
    Next
    
   pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom

   sht.Columns("C:F").NumberFormat = "#,##0"
   


End Sub




' This function is to create the pivot table for all BMs together
' input dataset is the dataset for entire AUM table
Sub CreatePivotTableForAll()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim NumOfRows As Long
Dim RangeStr As String


NumOfRows = ActiveSheet.UsedRange.Rows.Count
RangeStr = "A1:BF" & NumOfRows


 

 SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add
 sht.Name = "PivotTableAll"

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    
  'Add item to the Report Filter
  
  'Add item to the Column Labels
    'pvt.PivotFields("Month").Orientation = xlColumnField
    
  'Add item to the Row Labels
    pvt.PivotFields("Mgr").Orientation = xlRowField
    pvt.PivotFields("Mgr").Position = 1
    pvt.PivotFields("RM").Orientation = xlRowField
    pvt.PivotFields("RM").Position = 2
    
  'Position Item in list
    'pvt.PivotFields("Year").Position = 1
    
 
    
  'Turn on Automatic updates/calculations --like screenupdating to speed up code
    pvt.ManualUpdate = False


    
    
    Dim values() As String
    ReDim values(5)
    values(1) = "TOT_AUM"
    values(2) = "AUM_INCENTIVE"
    values(3) = "custnum"
    values(4) = "LIF_INSURANCE_PEN"
    values(5) = "GEN_INSURANCE_PEN"
    
    Dim values_name() As String
    ReDim values_name(5)
    values_name(1) = "Sum of TOT_AUM"
    values_name(2) = "Sum of AUM_INCENTIVE"
    values_name(3) = "Count of custnum"
    values_name(4) = "Sum of LIF_INSURANCE_PEN"
    values_name(5) = "Sum of GEN_INSURANCE_PEN"
    
    Dim i As Long
    For i = 1 To 5
        If i = 3 Then
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlCount
        Else
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlSum
        End If
    
    Next
    
   pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom


   sht.Columns("C").NumberFormat = "#,##0"


End Sub

Sub BMDrillDown()

    'Set nwSheet = Worksheets.Add
    'nwSheet.Activate
    Set sht = Worksheets("PivotTableAll")
    Set pvttable = Worksheets("PivotTableAll").Range("A1").PivotTable
    Dim numOfBM As Long
    Dim BMList() As String
    ReDim BMList(100)
    
    Dim BMInputList() As String
    BMInputList = getBMList()
    
    For i = 1 To sht.UsedRange.Rows.Count - 1
        Dim cell As String
        Dim BNamevalue As String
        Dim FName  As String
        
        
        If Not IsEmpty(sht.Cells(i, 1)) Then
        cellvalue = sht.Cells(i, 1).Value
        For Each j In BMInputList
            If cellvalue = j + " Total" And j <> "" And j <> " " Then
                'MsgBox (cellvalue)
                sht.Activate
                sht.Cells(i, 3).Select
                Selection.ShowDetail = True
                ActiveSheet.Name = j & "_RawData"
                CreatePivotTable
                ' PROBLEM OF ENDIF
                If getBranchCode(CStr(j)) = "" Or getBranchCode(CStr(j)) = "#" Then
                BName = j
                Else:
                BName = getBranchCode(CStr(j))
                End If
                ActiveSheet.Name = "AUM_" + BName
                ActiveSheet.Copy
                FName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - " & BName & ".xlsx"
                
                Application.DisplayAlerts = False
                With ActiveWorkbook
                    .SaveAs Filename:=FName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
                    .Close 0
                End With
                
                
                ActiveSheet.Delete
                Worksheets(j & "_RawData").Delete
                Application.DisplayAlerts = True
               
            End If
            
            
        Next
        End If
    Next
            
End Sub

Function getBMList() As Variant
    Dim nameList() As String
    FName = Application.ThisWorkbook.Path & "\" & "BMBBDcontact.xlsx"
    Set wkbk = Workbooks.Open(FName)
    Set sht = wkbk.Worksheets("BMcontact")
    ReDim nameList(sht.UsedRange.Rows.Count)
    For i = 2 To sht.UsedRange.Rows.Count
        nameList(i - 1) = sht.Cells(i, 1).Value
    Next
    wkbk.Close
    getBMList = nameList
End Function

Function getBranchCode(BMname As String) As String
    Dim ans As String
    ans = "#"
    FName = Application.ThisWorkbook.Path & "\" & "BMBBDcontact.xlsx"
    Set wkbk = Workbooks.Open(FName)
    Set sht = wkbk.Worksheets("BMcontact")
    For i = 2 To sht.UsedRange.Rows.Count
        If sht.Cells(i, 1) = BMname Then
            ans = sht.Cells(i, 2)
        End If
    Next
    wkbk.Close
    getBranchCode = ans
End Function


Private Sub SelectSavingDirectoryButton_Click()
    saveToPath = GetFolder
End Sub

Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = Application.ActiveWorkbook.Path
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function


Private Sub CheckPathButton_Click()
If saveToPath = "" Then
MsgBox ("YOU HAVE NOT SELECT ANY PATH")
Else
MsgBox (saveToPath)
End If
End Sub


Sub CreatePivotTableForAll_BBD()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim NumOfRows As Long
Dim RangeStr As String


NumOfRows = ActiveSheet.UsedRange.Rows.Count
RangeStr = "A1:BF" & NumOfRows


 

 SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add
 sht.Name = "PivotTableAllBBD"

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    
  'Add item to the Report Filter
  
  'Add item to the Column Labels
    'pvt.PivotFields("Month").Orientation = xlColumnField
    
  'Add item to the Row Labels
    pvt.PivotFields("BBD").Orientation = xlRowField
    pvt.PivotFields("BBD").Position = 1
    pvt.PivotFields("Mgr").Orientation = xlRowField
    pvt.PivotFields("Mgr").Position = 2
    pvt.PivotFields("RM").Orientation = xlRowField
    pvt.PivotFields("RM").Position = 3
    
  'Position Item in list
    'pvt.PivotFields("Year").Position = 1
    
 
    
  'Turn on Automatic updates/calculations --like screenupdating to speed up code
    pvt.ManualUpdate = False


    
    
    Dim values() As String
    ReDim values(5)
    values(1) = "TOT_AUM"
    values(2) = "AUM_INCENTIVE"
    values(3) = "custnum"
    values(4) = "LIF_INSURANCE_PEN"
    values(5) = "GEN_INSURANCE_PEN"
    
    Dim values_name() As String
    ReDim values_name(5)
    values_name(1) = "Sum of TOT_AUM"
    values_name(2) = "Sum of AUM_INCENTIVE"
    values_name(3) = "Count of custnum"
    values_name(4) = "Sum of LIF_INSURANCE_PEN"
    values_name(5) = "Sum of GEN_INSURANCE_PEN"
    
    Dim i As Long
    For i = 1 To 5
        If i = 3 Then
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlCount
        Else
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlSum
        End If
    
    Next
    
   pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom


   sht.Columns("C").NumberFormat = "#,##0"


End Sub

Sub BBDDrillDown()

    'Set nwSheet = Worksheets.Add
    'nwSheet.Activate
    Set sht = Worksheets("PivotTableAllBBD")
    Set pvttable = Worksheets("PivotTableAllBBD").Range("A1").PivotTable
    Dim numOfBBD As Integer
    Dim BBDList() As String
    ReDim BMList(100)
    
    Dim BBDInputList() As String
    BBDInputList = getBBDList()
    

    
    
    For i = 1 To sht.UsedRange.Rows.Count - 1
        Dim cell As String
        Dim BBDNamevalue As String
        Dim FName  As String
        
        
        If Not IsEmpty(sht.Cells(i, 1)) Then
        cellvalue = sht.Cells(i, 1).Value
        For Each j In BBDInputList
            If cellvalue = j + " Total" And j <> "" Then
                'MsgBox (cellvalue)
                sht.Activate
                sht.Cells(i, 4).Select
                Selection.ShowDetail = True
                ActiveSheet.Name = j & "_RawData"
                CreatePivotTable_BBD
                ' PROBLEM OF ENDIF
                If getBBDFirstName(CStr(j)) = "" Then
                BName = j
                Else:
                BName = getBBDFirstName(CStr(j))
                End If
                ActiveSheet.Name = "AUM_" + BName
                ActiveSheet.Copy
                FName = saveToPath & "\" & "ZONE AUM " & monthStr & " " & yearStr & " - " & BName
                
                Application.DisplayAlerts = False
                
                With ActiveWorkbook
                    .SaveAs Filename:=FName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
                    .Close 0
                End With
                Application.DisplayAlerts = True
            End If
            
            
        Next
        End If
    Next
            
End Sub

' This function is to create the pivot table for a BM
' input dataset is the dataset for entire BM (By double click the pivot table of AUM)
Sub CreatePivotTable_BBD()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim numOfRow As Long
numOfRow = ActiveSheet.UsedRange.Rows.Count



 SrcData = ActiveSheet.Name & "!" & Range("A1:BF" & numOfRow).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    
  'Add item to the Report Filter
  
  'Add item to the Column Labels
    'pvt.PivotFields("Month").Orientation = xlColumnField
    
  'Add item to the Row Labels
    pvt.PivotFields("BBD").Orientation = xlRowField
    pvt.PivotFields("BBD").Position = 1
    pvt.PivotFields("Mgr").Orientation = xlRowField
    pvt.PivotFields("Mgr").Position = 2
    pvt.PivotFields("RM").Orientation = xlRowField
    pvt.PivotFields("RM").Position = 3
    
  'Position Item in list
    'pvt.PivotFields("Year").Position = 1
    
 
    
  'Turn on Automatic updates/calculations --like screenupdating to speed up code
    pvt.ManualUpdate = False

    
    Dim values() As String
    ReDim values(5)
    values(1) = "TOT_AUM"
    values(2) = "AUM_INCENTIVE"
    values(3) = "custnum"
    values(4) = "LIF_INSURANCE_PEN"
    values(5) = "GEN_INSURANCE_PEN"
    
    Dim values_name() As String
    ReDim values_name(5)
    values_name(1) = "Sum of TOT_AUM"
    values_name(2) = "Sum of AUM_INCENTIVE"
    values_name(3) = "Count of custnum"
    values_name(4) = "Sum of LIF_INSURANCE_PEN"
    values_name(5) = "Sum of GEN_INSURANCE_PEN"
    
    Dim i As Long
    For i = 1 To 5
        If i = 3 Then
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlCount
        Else
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlSum
        End If
    
    Next
    
   pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom

   sht.Columns("D:G").NumberFormat = "#,##0"
   


End Sub

Function getBBDList() As Variant
    Dim nameList() As String
    FName = Application.ThisWorkbook.Path & "\" & "BMBBDcontact.xlsx"
    Set wkbk = Workbooks.Open(FName)
    Set sht = wkbk.Worksheets("BBDcontact")
    ReDim nameList(sht.UsedRange.Rows.Count)
    For i = 2 To sht.UsedRange.Rows.Count
        nameList(i - 1) = sht.Cells(i, 1).Value
    Next
    wkbk.Close
    getBBDList = nameList
End Function

Function getBBDFirstName(BBDName As String) As String
    Dim ans As String
    FName = Application.ThisWorkbook.Path & "\" & "BMBBDcontact.xlsx"
    Set wkbk = Workbooks.Open(FName)
    Set sht = wkbk.Worksheets("BBDcontact")
    For i = 2 To sht.UsedRange.Rows.Count
        If sht.Cells(i, 1) = BBDName Then
            ans = sht.Cells(i, 2)
        End If
    Next
    wkbk.Close
    getBBDFirstName = ans
End Function

Sub CreatePivotTableForAll_ZT()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim NumOfRows As Long
Dim RangeStr As String


NumOfRows = ActiveSheet.UsedRange.Rows.Count
RangeStr = "A1:BF" & NumOfRows

Dim BBDName() As String
BBDName = getBBDList()
 

 SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add
 sht.Name = "PivotTableAllZT"

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
   
  pvt.AddFields RowFields:=Array("BBD", "Mgr", "RM"), ColumnFields:=Array("Values", "Tier")

  pvt.ManualUpdate = False

   
    Dim values() As String
    ReDim values(5)
    values(1) = "TOT_AUM"
    values(2) = "AUM_INCENTIVE"
    values(3) = "custnum"
    values(4) = "LIF_INSURANCE_PEN"
    values(5) = "GEN_INSURANCE_PEN"
    
    Dim values_name() As String
    ReDim values_name(5)
    values_name(1) = "Sum of TOT_AUM"
    values_name(2) = "Sum of AUM_INCENTIVE"
    values_name(3) = "Count of custnum"
    values_name(4) = "Sum of LIF_INSURANCE_PEN"
    values_name(5) = "Sum of GEN_INSURANCE_PEN"
    
    Dim i As Long
    For i = 1 To 5
        If i = 3 Then
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlCount
        Else
          pvt.AddDataField pvt.PivotFields(values(i)), values_name(i), xlSum
        End If
    
    Next
    
   pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom


    For Each itm In pvt.PivotFields("BBD").PivotItems
        If itm.Name = "" Or itm.Name = "(blank)" Then
            itm.Visible = False
        End If
    Next
    
   Dim flag As Integer
   For Each itm In pvt.PivotFields("BBD").PivotItems
        flag = 0
        For Each nme In BBDName
        If CStr(itm.Name) = nme Then
            flag = 1
        End If
        Next
        If flag = 0 Then
        itm.Visible = False
        End If
   Next

   sht.Range("D:BF").NumberFormat = "#,##0"
   Dim numOfRow As Long
   numOfRow = sht.UsedRange.Rows.Count

   sht.Range("A1:BF3").Interior.Color = RGB(220, 230, 241)
   sht.Range("A1:BF3").Font.Bold = True
   sht.Range("D2:BA" + CStr(numOfRow - 2)).Interior.Color = RGB(217, 217, 217) 'data color
   sht.Range("A" + CStr(numOfRow) + ":BF" + CStr(numOfRow)).Interior.Color = RGB(220, 230, 241) 'grandtotal color
   sht.Range("A" + CStr(numOfRow) + ":BF" + CStr(numOfRow)).Font.Bold = True
   sht.Range("D3:BA3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   sht.Range("D3:BA3").Borders(xlEdgeBottom).Color = RGB(83, 141, 213)
   
   For i = 1 To numOfRow
    For Each nme In BBDName
        If sht.Cells(i, 1).Value = nme + " Total" Then
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Interior.Color = RGB(252, 213, 180) ' bbd color
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Font.Bold = True
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Borders(xlTop).LineStyle = xlContinuous
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Borders(xlEdgeBottom).Color = RGB(83, 141, 213)
        sht.Range("A" + CStr(i) + ":BF" + CStr(i)).Borders(xlTop).Color = RGB(83, 141, 213)
        End If
    Next
   Next
 
    ActiveSheet.Range("A:BF").Copy
    ActiveSheet.Range("A:BF").PasteSpecial xlPasteValues
    'sht.Range("A" + CStr(numOfRow) + ":AU" + CStr(numOfRow)).Interior.Color = RGB(252, 213, 180) ' bbd color
    ActiveSheet.Name = "AUM_Zone and Tiers"
   
    ActiveSheet.Copy
    FName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - Breakdown by Zone and Tiers"
                
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=FName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
        .Close 0
    End With
    Application.DisplayAlerts = True

End Sub

Private Sub ZipMailButton_Click()


If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If


Dim oApp As Object
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVED FILES")
Exit Sub
End If
Set oApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFSO.GetFolder(saveToPath)
Dim relfile As String
For Each objfile In objfolder.Files
    MsgBox objfile.Path
    'MsgBox objfile.Name & "---------" & objfile.Path
    
    If objfile.Name Like "*AUM ??? ####*.xlsx" Then
        zipName = Left(objfile.Path, InStrRev(objfile.Path, ".")) + "zip"
        NewZip (zipName)
        relfile = Left(objfile.Path, InStrRev(objfile.Path, "\")) & "RELNUM_RCAO_MAPPING " & Mid(objfile.Name, InStrRev(objfile.Name, "-"), InStrRev(objfile.Name, ".") - InStrRev(objfile.Name, "-")) & ".xlsx"
        MsgBox relfile
        'Workbooks.Open relfile
       
        oApp.Namespace(zipName).copyhere objfile.Path
        'oApp.Namespace(zipName).copyhere objfile.Path
        'Keep script waiting until Compressing is done
        On Error Resume Next
        Application.Wait (Now + TimeValue("0:00:05"))

    End If

Next

msg = MsgBox("ZIP files are generated. Please check files before email them." & vbNewLine & vbNewLine & "Are you sure to send the emails?", vbYesNo, "Email Confirmation")
If msg = vbYes Then
    sentMail
Else
    Exit Sub
End If

End Sub
Sub sentMail()

Application.ScreenUpdating = False

Dim oApp As Object
Dim outApp As Outlook.Application
Set oApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFSO.GetFolder(saveToPath)
Set thiswkbk = ActiveWorkbook
Workbooks.Open (ActiveWorkbook.Path & "\BMBBDcontact.xlsx")
Set cwkbk = ActiveWorkbook
Dim toStr As String
Dim ccStr As String
Dim subjectStr As String
Dim bodyStr As String
Dim monStr As String
Dim yearStr As String
Dim nameStr As String
Dim BMname As String
Dim BRcode As String

Dim zipCount As Integer
Dim mailCount As Integer
zipCount = 0
mailCount = 0



For Each objfile In objfolder.Files
    
    'MsgBox objfile.Name & "---------" & objfile.Path
    
    If objfile.Name Like "*AUM ??? ####*.zip" Then
        toStr = ""
        ccStr = ""
        subjectStr = ""
        bodyStr = ""
        BMname = ""
        nameStr = ""
        zipCount = zipCount + 1
        'AUM CUtting for for zone and tiers
        If objfile.Name Like "*Breakdown by Zone and Tiers*" Then
            monStr = Mid(objfile.Name, 5, 3)
            yearStr = Mid(objfile.Name, 9, 4)
            Set sht = cwkbk.Worksheets("Zone & Tiers")
            toStr = sht.Cells(2, 1).Value
            ccStr = sht.Cells(2, 2).Value
            subjectStr = "Branch AUM " & monStr & " " & yearStr & " - Breakdown by Zone and Tiers"
    
            bodyStr = "Hi All," & vbNewLine & vbNewLine & _
                          "Please see attached for branch AUM (Breakdown by Zone and Tiers) for " & monStr & " " & yearStr & "." & vbNewLine & vbNewLine & _
                          "Password : Uni******" & Right(yearStr, 2) & vbNewLine & vbNewLine & _
                          "Thank you!" & vbNewLine & vbNewLine & _
                          "Best Regards," & vbNewLine & "Ding Ming"
         'AUM Cutting file for BBD
         ElseIf objfile.Name Like "ZONE AUM *" Then
            monStr = Mid(objfile.Name, 10, 3)
            yearStr = Mid(objfile.Name, 14, 4)
            Set sht = cwkbk.Worksheets("BBDcontact")
            
            nameStr = Mid(objfile.Name, InStr(1, objfile.Name, "- ") + 2, InStr(1, objfile.Name, ".") - InStr(1, objfile.Name, "- ") - 2)
    
            For i = 2 To sht.UsedRange.Rows.Count
                If sht.Cells(i, 2).Value = nameStr Then
                    toStr = sht.Cells(i, 3).Value
                    ccStr = sht.Cells(i, 4).Value
                End If
            Next
            
            subjectStr = "ZONE AUM " & monStr & " " & yearStr & " (" & nameStr & ")"
            
            bodyStr = "Hi " & nameStr & "," & vbNewLine & vbNewLine & _
                          "Please see attached for zone AUM for " & monStr & " " & yearStr & "." & vbNewLine & vbNewLine & _
                          "Password : Uni******" & Right(yearStr, 2) & vbNewLine & vbNewLine & _
                          "Thank you!" & vbNewLine & vbNewLine & _
                          "Best Regards," & vbNewLine & "Ding Ming"
            
            'AUM Cutting file for BM
            ElseIf objfile.Name Like "AUM ??? ####* - *" Then
            monStr = Mid(objfile.Name, 5, 3)
            yearStr = Mid(objfile.Name, 9, 4)
            Set sht = cwkbk.Worksheets("BMcontact")
            
            BRcode = Mid(objfile.Name, InStr(1, objfile.Name, "- ") + 2, InStr(1, objfile.Name, ".") - InStr(1, objfile.Name, "- ") - 2)
    
            For i = 2 To sht.UsedRange.Rows.Count
                If sht.Cells(i, 2).Value = BRcode Then
                    toStr = sht.Cells(i, 3).Value
                    ccStr = sht.Cells(i, 4).Value
                    BMname = sht.Cells(i, 5).Value
                    
                End If
            Next
            
            subjectStr = "AUM " & monStr & " " & yearStr & " (" & BRcode & ")"
            
            bodyStr = "Hi " & BMname & "," & vbNewLine & vbNewLine & _
                          "Please see attached for branch AUM for " & monStr & " " & yearStr & "." & vbNewLine & vbNewLine & _
                          "Please reference The Relationship Number - RCAO mapping file if needed." & vbNewLine & vbNewLine & _
                          "Password : Uni******" & Right(yearStr, 2) & vbNewLine & vbNewLine & _
                          "Thank you!" & vbNewLine & vbNewLine & _
                          "Best Regards," & vbNewLine & "Ding Ming"
         End If
            If toStr <> "" Then
                Set outApp = New Outlook.Application
                Set OutMail = outApp.CreateItem(olMailItem)
            
                On Error Resume Next
                With OutMail
                    .To = toStr
                    .cc = ccStr
                    .BCC = ""
                    .Subject = subjectStr
                    .Body = bodyStr
                    .Attachments.Add objfile.Path
                    ssstr = Right(objfile.Name, Len(objfile.Name) - InStrRev(objfile.Name, "-"))
                    ssstr1 = Left(ssstr, InStr(ssstr, ".") - 1)
             
                    ssstr2 = Right(ssstr1, Len(ssstr1) - 1)
                  
             
                    '.Attachments.Add "G:\Plus\(SK) AUM Cutting\7. Jul15\RELNUM_RCAO_MAPPING - " & ssstr2 & ".xlsx"
                    .Display
                End With
                mailCount = mailCount + 1
            End If
        
    End If
Next
Application.DisplayAlerts = False
cwkbk.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox ("Zip & Mail has been done!" & vbNewLine & zipCount & " zips files are generated." & vbNewLine & mailCount & " mails are sent.")
End Sub
Sub Zip_All_Files_in_Folder_Browse()
    Dim FileNameZip, FolderName, oFolder
    Dim strDate As String, DefPath As String
    Dim oApp As Object

    DefPath = "I:\CAP_Profile_PRD65\Desktop"
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"

    Set oApp = CreateObject("Shell.Application")

    'Browse to the folder
    Set oFolder = oApp.BrowseForFolder(0, "Select folder to Zip", 512)
    If Not oFolder Is Nothing Then
        'Create empty Zip File
        NewZip (FileNameZip)

        FolderName = oFolder.Self.Path
        If Right(FolderName, 1) <> "\" Then
            FolderName = FolderName & "\"
        End If

        'Copy the files to the compressed folder
        oApp.Namespace(FileNameZip).copyhere oApp.Namespace(FolderName).items
        

        'Keep script waiting until Compressing is done
        On Error Resume Next
        Do Until oApp.Namespace(FileNameZip).items.Count = _
        oApp.Namespace(FolderName).items.Count
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0

        MsgBox "You find the zipfile here: " & FileNameZip

    End If
End Sub

Private Sub ZipMailButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub ZTCutButton_Click()
    Application.ScreenUpdating = False

Dim intChoice As Long
Dim strPath As String


If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

monthStr = InputBox(Prompt:="Please input the month in CAPITAL LETTERS (MMM)", _
          Title:="Please input the month (MMM)")
If monthStr <> "JAN" And monthStr <> "FEB" And monthStr <> "MAR" _
And monthStr <> "APR" And monthStr <> "MAY" And monthStr <> "JUN" _
And monthStr <> "JUL" And monthStr <> "AUG" And monthStr <> "SEP" _
And monthStr <> "OCT" And monthStr <> "NOV" And monthStr <> "DEC" Then
MsgBox ("PLEASE INPUT A CORRECT MONTH!")
Exit Sub
End If

yearStr = InputBox(Prompt:="Please input the year (YYYY)", _
          Title:="Please input the year (YYYY)")
If yearStr > 9999 Or yearStr < 2000 Then
MsgBox ("PLEASE INPUT A CORRECT YEAR!")
Exit Sub
End If






Application.FileDialog(msoFileDialogOpen).Title = "Select the G&B file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
Else: Exit Sub
End If

    Set wkbk = Workbooks.Open(strPath)
    CreatePivotTableForAll_ZT
    
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
MsgBox ("AUM Cutting for Zone and Tiers is Done, Please view files in indicated folder")

End Sub




' Special Case 1: Data of one BM entirely belong to another BM

Function SPC_1(NameOfAttach As String, NameToAttach As String) As Integer

    Application.DisplayAlerts = False
    For Each sht In mainwkbk.Worksheets
        If sht.Name = "Temp" Or sht.Name = "ToAppend" Then
        sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
    
    SPC_1 = 0
    Dim flag As Integer
    Dim numOfRow As Long
    Dim check As Integer
    flag = 0
    check = 0
    Dim NameOfBranch As String
    NameOfBranch = getBranchCode(NameToAttach)
    If NameOfBranch = "#" Then
        Exit Function
    End If
    Set sht = mainwkbk.Worksheets("PivotTableAll")
    sht.Activate
    numOfRow = sht.UsedRange.Rows.Count
    Dim i As Long
    For i = 2 To numOfRow - 1
        If check = 2 Then
            Exit For
        End If
        sht.Activate
        If sht.Cells(i, 1).Value = NameOfAttach & " Total" Then
            sht.Cells(i, 3).Select
            Selection.ShowDetail = True
            ActiveSheet.Name = "Temp"
            flag = 1
            check = check + 1
        End If
        If sht.Cells(i, 1).Value = NameToAttach & " Total" Then
            sht.Cells(i, 3).Select
            Selection.ShowDetail = True
            ActiveSheet.Name = "ToAppend"
            check = check + 1
        End If
    Next
    
    If flag = 0 Or check <> 2 Then
        Exit Function
    End If
    
    numOfRow = Worksheets("Temp").UsedRange.Rows.Count
    Worksheets("Temp").Rows("2:" + CStr(numOfRow)).Copy
    
    numOfRow = Worksheets("ToAppend").UsedRange.Rows.Count
    Worksheets("ToAppend").Activate
    ActiveSheet.Cells(numOfRow + 1, 1).Select
    Selection.PasteSpecial xlPasteAll

    CreatePivotTable
    ActiveSheet.Name = "AUM_" + NameOfBranch
    Application.DisplayAlerts = False
    mainwkbk.Worksheets("Temp").Delete
    mainwkbk.Worksheets("ToAppend").Delete
    Application.DisplayAlerts = True

    ActiveSheet.Copy
    Dim FName As String
    FName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - " & NameOfBranch
                
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=FName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
        .Close 0
    End With
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    SPC_1 = 1
End Function


'Special Case 2: One user name of BM is not belong to him. Filter out that data.
'E.G. MDH 2 remote from Constance Choong should be removed

Function SPC_2(NameOfBM As String, User_NM As String) As Integer
    Application.DisplayAlerts = False
    For Each sht In mainwkbk.Worksheets
        If sht.Name = "Temp" Or sht.Name = "ToAppend" Then
        sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
    SPC_2 = 0
    On Error GoTo lb1:
    Dim FName As String
    Dim NameOfBranch As String
    NameOfBranch = getBranchCode(NameOfBM)
    If NameOfBranch = "#" Then
    Exit Function
    End If
    FName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - " & NameOfBranch & ".xlsx"
    
    Set wkbk = Workbooks.Open(FName, Password:="Universal" & Right(yearStr, 2))
    Set pvt = ActiveSheet.PivotTables(1)
    'For Each itm In pvt.PivotFields("USER_NM").PivotItems
    '    If itm.Name = User_NM Then
    '    itm.Visible = False
    '    End If
    'Next
    pvt.PivotFields("USER_NM").PivotItems(User_NM).Visible = False
    Application.DisplayAlerts = False
    wkbk.Save Password:="Universal" & Right(yearStr, 2)
    wkbk.Close
    Application.DisplayAlerts = True
    SPC_2 = 1
    Exit Function
lb1:
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
End Function

'Special Case 3: One user name of BM should be counted to other BM
'E.G. MDH 2 Remote from Constance Choong should be counted to Evelyn Fong

Function SPC_3(USERNM As String, NameToAttach As String) As Integer
    Application.DisplayAlerts = False
    For Each sht In mainwkbk.Worksheets
        If sht.Name = "Temp" Or sht.Name = "ToAppend" Then
        sht.Delete
        End If
    Next

    Application.DisplayAlerts = True
    SPC_3 = 0
    Dim flag As Integer
    Dim numOfRow As Long
    Dim check As Integer
    flag = 0
    check = 0
    Dim NameOfBranch As String
    NameOfBranch = getBranchCode(NameToAttach)
    If NameOfBranch = "#" Then
        Exit Function
    End If
    Set sht = mainwkbk.Worksheets("PivotTableAll")
    sht.Activate
    numOfRow = sht.UsedRange.Rows.Count
    Dim i As Long
    
    For i = 2 To numOfRow - 1
        If check = 2 Then
            Exit For
        End If
        sht.Activate
        If sht.Cells(i, 2).Value = USERNM Then
            sht.Cells(i, 3).Select
            Selection.ShowDetail = True
            ActiveSheet.Name = "Temp"
            flag = 1
            check = check + 1
        End If
        If sht.Cells(i, 1).Value = NameToAttach & " Total" Then
            sht.Cells(i, 3).Select
            Selection.ShowDetail = True
            ActiveSheet.Name = "ToAppend"
            check = check + 1
        End If
    Next
    
    If flag = 0 Or check <> 2 Then
        Exit Function
    End If
    
    numOfRow = Worksheets("Temp").UsedRange.Rows.Count
    Worksheets("Temp").Rows("2:" + CStr(numOfRow)).Copy
    
    numOfRow = Worksheets("ToAppend").UsedRange.Rows.Count
    Worksheets("ToAppend").Activate
    ActiveSheet.Cells(numOfRow + 1, 1).Select
    Selection.PasteSpecial xlPasteAll

    CreatePivotTable
    ActiveSheet.Name = "AUM_" + NameOfBranch
    Application.DisplayAlerts = False
    mainwkbk.Worksheets("Temp").Delete
    mainwkbk.Worksheets("ToAppend").Delete
    Application.DisplayAlerts = True

    ActiveSheet.Copy
    Dim FName As String
    FName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - " & NameOfBranch
                
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=FName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
        .Close 0
    End With
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    SPC_3 = 1
End Function



Sub getSPC()
    Dim result As Integer
    FName = Application.ThisWorkbook.Path & "\" & "SpecialCase.xlsx"
    Set wkbk = Workbooks.Open(FName)
    Set sht = wkbk.Worksheets("Data")
    For i = 2 To sht.UsedRange.Rows.Count
    result = -1
        Select Case sht.Cells(i, 1).Value
        Case "SPC_1"
            result = SPC_1(sht.Cells(i, 2).Value, sht.Cells(i, 3).Value)
        Case "SPC_2"
            result = SPC_2(sht.Cells(i, 2).Value, sht.Cells(i, 3).Value)
        Case "SPC_3"
            result = SPC_3(sht.Cells(i, 2).Value, sht.Cells(i, 3).Value)
        Case Else
            result = 0
        End Select
        sht.Rows(i).Interior.Color = xlNone
        If result = 0 Then
            sht.Rows(i).Interior.Color = vbRed
        Else
            sht.Rows(i).Interior.Color = vbGreen
        End If
    Next
    Application.DisplayAlerts = False
    With wkbk
        .Save
        .Close 0
    End With
    Application.DisplayAlerts = True
End Sub

Sub CombineButton_click()

Dim strPath1 As String
Dim strPath2 As String
Dim numOfRow1 As Long
Dim numOfRow2 As Long

 Application.FileDialog(msoFileDialogOpen).Title = "Select the Blue and Gold file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = Application.ActiveWorkbook.Path
intChoice = Application.FileDialog(msoFileDialogOpen).Show

Application.ScreenUpdating = False
If intChoice <> 0 Then
    'get the file path selected by the user
    If Application.FileDialog(msoFileDialogOpen).SelectedItems.Count <> 2 Then
        MsgBox "Please select both Gold and Blue files!"
        Exit Sub
    Else
        strPath1 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        strPath2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(2)
        If Not (strPath1 Like "*\Gold*" Or strPath1 Like "*\Blue*") Then
            MsgBox "Please Select only Gold and Blue files!"
            Exit Sub
        End If
        If Not (strPath2 Like "*\Gold*" Or strPath2 Like "*\Blue*") Then
            MsgBox "Please Select only Gold and Blue files!"
            Exit Sub
        End If
        
        Workbooks.Open (strPath1)
        Set wkbk1 = ActiveWorkbook
        Workbooks.Open (strPath2)
        Set wkbk2 = ActiveWorkbook
        numOfRow1 = wkbk1.Worksheets(1).UsedRange.Rows.Count
        numOfRow2 = wkbk2.Worksheets(1).UsedRange.Rows.Count
        wkbk1.Worksheets(1).Rows("2:" + CStr(numOfRow1)).Copy
        wkbk2.Worksheets(1).Rows(CStr(numOfRow2 + 1) + ":" + CStr(numOfRow1 + numOfRow2 - 1)).PasteSpecial xlPasteValues
        ' swap col B & C if they are reverted
        If wkbk1.Worksheets(1).Cells(1, 2).Value = wkbk2.Worksheets(1).Cells(1, 3).Value And wkbk1.Worksheets(1).Cells(1, 3).Value = wkbk2.Worksheets(1).Cells(1, 2).Value Then
            For i = numOfRow2 + 1 To numOfRow1 + numOfRow2 - 1
                Dim swap As Long
                swap = wkbk2.Worksheets(1).Cells(i, 2).Value
                wkbk2.Worksheets(1).Cells(i, 2).Value = wkbk2.Worksheets(1).Cells(i, 3).Value
                wkbk2.Worksheets(1).Cells(i, 3).Value = swap
            Next
        End If
        
        Dim saveToPath As String
        saveToPath = Replace(strPath1, "Gold", "G&B")
        saveToPath = Replace(saveToPath, "Blue", "G&B")
        Application.DisplayAlerts = False
        With wkbk2
            .SaveAs Filename:=saveToPath, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
            .Close 0
        End With
        wkbk1.Close
        Application.DisplayAlerts = True
    End If
    'print the file path to sheet 1
Else: Exit Sub
End If
Application.ScreenUpdating = True
MsgBox "G&B File generation completed!"
End Sub

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function









