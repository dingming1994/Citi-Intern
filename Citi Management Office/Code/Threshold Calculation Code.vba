Public sopbk, prevbk, mvmtbk, outbk As Workbook
Public monthstr As String
Public isCPCLoad, isRMLoad, isPBLoad As Integer





Sub CreatePT()

Set outbk = ActiveWorkbook

prevmon = getPrevMonth()
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim numOfRow As Long

On Error Resume Next

Set sht = outbk.Worksheets.Add(after:=outbk.Worksheets(3))
sht.Name = "Pivot Summary"
'CPC TABLE

Set shet = outbk.Worksheets("CPC")
shet.Activate
numOfRow = shet.Range("A2").End(xlDown).Row
 
 'SrcData = shet.Name & "!" & Range("A2:N" & numOfRow).Address(ReferenceStyle:=xlR1C1)
SrcData = "'" & shet.Name & "'!" & Range("A2:N" & numOfRow).Address(ReferenceStyle:=xlR1C1)
 StartPvt = "'" & sht.Name & "'!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)
  Set pvtCache = outbk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    If isCPCLoad = 1 Then
        pvt.AddFields RowFields:=Array("Branch", "CPC")
        pvt.PivotFields("Indicator").Orientation = xlPageField
        pvt.PivotFields("Indicator").CurrentPage = "LOAD"
    Else
        pvt.AddFields RowFields:=Array("Branch")
        pvt.PivotFields("Indicator").Orientation = xlPageField
    End If
    pvt.AddDataField pvt.PivotFields(prevmon & " Cr Threshold"), "Sum of " & prevmon & " Cr Threshold", xlSum
    pvt.AddDataField pvt.PivotFields(prevmon & " AUM Movement"), "Sum of " & prevmon & " AUM Movement", xlSum
    pvt.AddDataField pvt.PivotFields(monthstr & " Cr Threshold"), "Sum of " & monthstr & " Cr Threshold", xlSum
    pvt.InGridDropZones = True

    pvt.ColumnGrand = True
    pvt.SubtotalLocation xlAtBottom
    pvt.PivotFields("Branch").Subtotals(1) = False
    If isCPCLoad = 1 Then
        pvt.PivotFields("CPC").Subtotals(1) = False
    End If
    
'RM TABLE
Set shet = outbk.Worksheets("RM")
shet.Activate
numOfRow = shet.Range("A2").End(xlDown).Row

SrcData = "'" & shet.Name & "'!" & Range("A2:O" & numOfRow).Address(ReferenceStyle:=xlR1C1)



startrow = sht.PivotTables(1).TableRange1.Rows(sht.PivotTables(1).TableRange1.Rows.Count).Row + 6
 StartPvt = "'" & sht.Name & "'!" & sht.Range("A" & CStr(startrow)).Address(ReferenceStyle:=xlR1C1)
  Set pvtCache = outbk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")
    If isRMLoad = 1 Then
        pvt.AddFields RowFields:=Array("Branch", "RM")
        pvt.PivotFields("Indicator").Orientation = xlPageField
        pvt.PivotFields("Indicator").CurrentPage = "LOAD"
    Else
        pvt.AddFields RowFields:=Array("Branch")
        pvt.PivotFields("Indicator").Orientation = xlPageField
    End If
    pvt.AddDataField pvt.PivotFields(prevmon & " Cr Threshold"), "Sum of " & prevmon & " Cr Threshold", xlSum
    pvt.AddDataField pvt.PivotFields(prevmon & " AUM Movement"), "Sum of " & prevmon & " AUM Movement", xlSum
    pvt.AddDataField pvt.PivotFields(monthstr & " Cr Threshold"), "Sum of " & monthstr & " Cr Threshold", xlSum
    
    pvt.PivotFields("Branch").Subtotals(1) = False
    If isRMLoad = 1 Then
        pvt.PivotFields("RM").Subtotals(1) = False
    End If
    pvt.InGridDropZones = True
    pvt.ColumnGrand = True
    pvt.SubtotalLocation xlAtBottom

'PB TABLE
Set shet = outbk.Worksheets("PB")
shet.Activate
numOfRow = shet.Range("A2").End(xlDown).Row

SrcData = "'" & shet.Name & "'!" & Range("A2:N" & numOfRow).Address(ReferenceStyle:=xlR1C1)

startrow = 0
startrow = sht.PivotTables("PivotTable2").TableRange1.Rows(sht.PivotTables("PivotTable2").TableRange1.Rows.Count).Row + 6
 StartPvt = "'" & sht.Name & "'!" & sht.Range("A" & CStr(startrow)).Address(ReferenceStyle:=xlR1C1)
  Set pvtCache = outbk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable3")
    If isPBLoad = 1 Then
        pvt.AddFields RowFields:=Array("Branch", "PB")
        pvt.PivotFields("Indicator").Orientation = xlPageField
        pvt.PivotFields("Indicator").CurrentPage = "LOAD"
    Else
        pvt.AddFields RowFields:=Array("Branch")
        pvt.PivotFields("Indicator").Orientation = xlPageField
    End If
   pvt.AddDataField pvt.PivotFields(prevmon & " Cr Threshold"), "Sum of " & prevmon & " Cr Threshold", xlSum
    pvt.AddDataField pvt.PivotFields(prevmon & " AUM Movement"), "Sum of " & prevmon & " AUM Movement", xlSum
    pvt.AddDataField pvt.PivotFields(monthstr & " Cr Threshold"), "Sum of " & monthstr & " Cr Threshold", xlSum
    
    pvt.PivotFields("Branch").Subtotals(1) = False
    If isPBLoad = 1 Then
        pvt.PivotFields("PB").Subtotals(1) = False
    End If
    pvt.InGridDropZones = True
    pvt.ColumnGrand = True
    pvt.SubtotalLocation xlAtBottom
    
    
sht.Range("B:D").NumberFormat = "0,0"
sht.Activate

sht.Range("F1") = "Role"
sht.Range("G1") = "Branch"
sht.Range("H1") = "Name"
sht.Range("I1") = prevmon & " Threshold"
sht.Range("J1") = prevmon & " AUM Movement"
sht.Range("K1") = monthstr & " Threshold"
sht.Range("L1") = "Variance"
idx = 1
If isCPCLoad = 1 Then
    For i = 1 To sht.PivotTables("PivotTable1").TableRange1.Rows(sht.PivotTables("PivotTable1").TableRange1.Rows.Count).Row
        If sht.Cells(i, 2) > 0 And sht.Cells(i, 2) < 10000000 And sht.Cells(i, 1) <> "Grand Total" Then
            idx = idx + 1
            sht.Cells(idx, 6) = "CPC"
            sht.Range(sht.Cells(i, 1), sht.Cells(i, 4)).Copy
            sht.Cells(idx, 8).PasteSpecial xlPasteValues
            sht.Cells(idx, 12) = sht.Cells(idx, 11) - sht.Cells(idx, 9)
            'find bbd
            p = i
            While sht.Cells(p, 2) <> ""
                p = p - 1
            Wend
            sht.Cells(idx, 7) = sht.Cells(p, 1)
        End If
    Next
End If

endi = i
If isRMLoad = 1 Then
    For i = endi To sht.PivotTables("PivotTable2").TableRange1.Rows(sht.PivotTables("PivotTable2").TableRange1.Rows.Count).Row
        If sht.Cells(i, 2) > 0 And sht.Cells(i, 2) < 10000000 And sht.Cells(i, 1) <> "Grand Total" Then
            idx = idx + 1
            sht.Cells(idx, 6) = "RM"
            sht.Range(sht.Cells(i, 1), sht.Cells(i, 4)).Copy
            sht.Cells(idx, 8).PasteSpecial xlPasteValues
            sht.Cells(idx, 12) = sht.Cells(idx, 11) - sht.Cells(idx, 9)
            'find bbd
            p = i
            While sht.Cells(p, 2) <> ""
                p = p - 1
            Wend
            sht.Cells(idx, 7) = sht.Cells(p, 1)
        End If
    Next
End If

endi = i
If isPBLoad = 1 Then
    For i = endi To sht.PivotTables("PivotTable3").TableRange1.Rows(sht.PivotTables("PivotTable3").TableRange1.Rows.Count).Row
        If sht.Cells(i, 2) > 0 And sht.Cells(i, 2) < 10000000 And sht.Cells(i, 1) <> "Grand Total" Then
            idx = idx + 1
            sht.Cells(idx, 6) = "PB"
            sht.Range(sht.Cells(i, 1), sht.Cells(i, 4)).Copy
            sht.Cells(idx, 8).PasteSpecial xlPasteValues
            sht.Cells(idx, 12) = sht.Cells(idx, 11) - sht.Cells(idx, 9)
            'find bbd
            p = i
            While sht.Cells(p, 2) <> ""
                p = p - 1
            Wend
            sht.Cells(idx, 7) = sht.Cells(p, 1)
        End If
    Next
End If

sht.Columns("F:L").AutoFit
sht.Columns("F:G").Font.Bold = True
sht.Columns("I:L").NumberFormat = "0,0"
With sht.Range("F1").CurrentRegion
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Rows(1).Interior.ColorIndex = 11
    .Rows(1).Font.Bold = True
    .Rows(1).Font.Color = vbWhite
    For i = 2 To .Rows.Count
        If sht.Cells(i, 6) = "CPC" Then
            sht.Cells(i, 6).Interior.ColorIndex = 40
        End If
        If sht.Cells(i, 6) = "RM" Then
            sht.Cells(i, 6).Interior.ColorIndex = 19
        End If
        If sht.Cells(i, 6) = "PB" Then
            sht.Cells(i, 6).Interior.ColorIndex = 44
        End If
        If sht.Cells(i, 12) >= 500 Then
            sht.Cells(i, 12).Interior.Color = vbYellow
        End If
    Next
End With
End Sub

Function getPrevMonth() As String
    Select Case Left(monthstr, 3)
        Case "Jan": ans = "Dec" & CStr(CInt(Right(monthstr, 2)) - 1)
        Case "Feb": ans = "Jan" & Right(monthstr, 2)
        Case "Mar": ans = "Feb" & Right(monthstr, 2)
        Case "Apr": ans = "Mar" & Right(monthstr, 2)
        Case "May": ans = "Apr" & Right(monthstr, 2)
        Case "Jun": ans = "May" & Right(monthstr, 2)
        Case "Jul": ans = "Jun" & Right(monthstr, 2)
        Case "Aug": ans = "Jul" & Right(monthstr, 2)
        Case "Sep": ans = "Aug" & Right(monthstr, 2)
        Case "Oct": ans = "Sep" & Right(monthstr, 2)
        Case "Nov": ans = "Oct" & Right(monthstr, 2)
        Case "Dec": ans = "Nov" & Right(monthstr, 2)
     End Select
     
          
     getPrevMonth = ans
        
End Function

Sub BankerPages()

Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP masterlist file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\apacdfs\SG\GCG\USERS\md34851\CAP_Profile_PRD65\Desktop\threshold calculation"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    If Not strpath Like "*SOP Masterlist ####*" Then
        MsgBox "Please select the correct SOP file."
        Exit Sub
    End If
Else: Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select previous month threshold file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\apacdfs\SG\GCG\USERS\md34851\CAP_Profile_PRD65\Desktop\threshold calculation"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the AUM movement file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\(SK) AUM Movement Report for 2015"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath3 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

monthstr = InputBox(Prompt:="Please input the month and year (MmmYY)", _
          Title:="Please input the month and year")
Set prevbk = Workbooks.Open(strpath2, False)
Set mvmtbk = Workbooks.Open(strpath3, False)
Set sopbk = Workbooks.Open(strpath, False)
Set outbk = Workbooks.Add
UpdateMonthTitle
PBPage
RMPage
CPCPage

Application.DisplayAlerts = False
sopbk.Close
prevbk.Close
mvmtbk.Close
Application.DisplayAlerts = True
CreatePT

End Sub


Sub PBPage()

isPBLoad = 0
For Each sht In sopbk.Worksheets
  
    If sht.Name Like "*PB" Then
        Set soppbsht = sht
    End If
Next

Set pbsht = outbk.Worksheets(1)
pbsht.Name = "PB"

ThisWorkbook.Worksheets("title").Activate
ThisWorkbook.Worksheets("title").Rows(1).Copy

pbsht.Activate
pbsht.Rows(2).PasteSpecial xlPasteAll

idx = 3
For i = 4 To soppbsht.Range("G4").End(xlDown).Row
    pbsht.Cells(idx, 1) = soppbsht.Cells(i, 4)
    pbsht.Cells(idx, 2) = soppbsht.Cells(i, 8)
    pbsht.Cells(idx, 3) = soppbsht.Cells(i, 7)
    pbsht.Cells(idx, 4) = "Q"
    pbsht.Cells(idx, 5) = soppbsht.Cells(i, 13)
    pbsht.Cells(idx, 6) = 5600
    idx = idx + 1
Next

pbsht.Columns("E:E").NumberFormat = "MM-DD-YYYY"

For Each sht In prevbk.Worksheets
    If sht.Name Like "PB*" Then
        Set tmpsht = sht
    End If
Next


Dim c As Variant
idx = 3
For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    
    On Error Resume Next
    
    c = Application.VLookup(pbsht.Cells(i, 2).Value, tmpsht.Range("B:O"), 13, False)
    If Not (c Is Nothing) Then
        pbsht.Cells(i, 7).Value = c
        
    End If
Next

Set tmpsht = mvmtbk.Worksheets("PB")

Set pvt = tmpsht.PivotTables("PivotTable1")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(pbsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 3), tmpsht.Cells(tmpsht.Cells(tstartrow, 3).End(xlDown).Row, 4)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           pbsht.Cells(i, 8) = c
        End If
    End If
    
 Next
 
 Set pvt = tmpsht.PivotTables("PivotTable2")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(pbsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 9), tmpsht.Cells(tmpsht.Cells(tstartrow, 9).End(xlDown).Row, 10)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           pbsht.Cells(i, 9) = c
        End If
    End If
    
 Next

Set pvt = tmpsht.PivotTables("PivotTable3")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(pbsht.Cells(i, 3), tmpsht.Range(tmpsht.Cells(tstartrow, 15), tmpsht.Cells(tmpsht.Cells(tstartrow, 15).End(xlDown).Row, 16)), 2, False)
    If Not (c Is Nothing) Then
        If IsError(c) = False Then
           pbsht.Cells(i, 10) = c
        End If
    End If
Next



'vlookup AUM data

For Each sht In prevbk.Worksheets
    If sht.Name Like "PB*" Then
        Set tmpsht = sht
    End If
Next


'USE FOR LOOP. put in monthly movement data from prev bk

monthCt = Month(DateValue("01 " & Left(monthstr, 3) & " 2015"))
For j = 1 To monthCt - 2
    For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
        
        On Error Resume Next
        
        c = Application.VLookup(pbsht.Cells(i, 2).Value, tmpsht.Range("B:AB"), 15 + j, False)
        If Not (c Is Nothing) Then
            pbsht.Cells(i, 16 + j).Value = c
            
        End If
    Next
Next


pbsht.Range("K1").Value = 367
For i = 3 To pbsht.Range("A4").End(xlDown).Row
    pbsht.Range("K" & CStr(i)).Formula = "=sum(H" & CStr(i) & ":I" & CStr(i) & ") -J" & CStr(i)
Next

'Paste current AUM movement
For i = 3 To pbsht.Range("A4").End(xlDown).Row
    pbsht.Cells(i, 15 + monthCt) = pbsht.Cells(i, 11)
Next

For i = 3 To pbsht.Range("A4").End(xlDown).Row
    pbsht.Cells(i, 29).FormulaR1C1 = "=SUM(RC[-11] : RC[-1])"
Next

For i = 3 To pbsht.Range("A4").End(xlDown).Row
    pbsht.Cells(i, 30).FormulaR1C1 = "=sum(RC[-1], RC[-13])"
Next

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    On Error Resume Next
    
    c = Application.VLookup(pbsht.Cells(i, 2).Value, tmpsht.Range("B:AE"), 30, False)
    If Not (c Is Nothing) Then
        pbsht.Cells(i, 31).Value = c
        
    End If
Next

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    pbsht.Range("AF" & CStr(i)).Formula = "=IF((AD" & CStr(i) & "<AE" & CStr(i) & "),""xxx"",""Yes"")"
    pbsht.Range("AF" & CStr(i)).Interior.Color = vbYellow
Next

For i = 3 To pbsht.Cells(3, 3).End(xlDown).Row
    pbsht.Range("L" & CStr(i)).Formula = "=IF(AND(K" & CStr(i) & ">1000000, AF" & CStr(i) & "=""YES""), ""LOAD"","""")"
    If IsError(pbsht.Range("L" & CStr(i))) = False Then
    If pbsht.Range("L" & CStr(i)) = "LOAD" Then

        pbsht.Rows(i).Font.Color = vbRed
        isPBLoad = 1
    End If
    End If
    pbsht.Range("M" & CStr(i)).Formula = "=IF(K" & CStr(i) & "<0,(K" & CStr(i) & "/1000000)*$K$1,IF(L" & CStr(i) & "=""LOAD"",(K" & CStr(i) & "/1000000)*$K$1,0))"
    pbsht.Range("N" & CStr(i)).Formula = "=IF((G" & CStr(i) & "+M" & CStr(i) & ")<F" & CStr(i) & ",F" & CStr(i) & ",(G" & CStr(i) & "+M" & CStr(i) & "))"
Next

pbsht.Columns("A:C").AutoFit

End Sub


Sub RMPage()

isRMLoad = 0

For Each sht In sopbk.Worksheets
  
    If sht.Name Like "*RM" Then
        Set soprmsht = sht
    End If
Next

Set rmsht = outbk.Worksheets(2)
rmsht.Name = "RM"

ThisWorkbook.Worksheets("title").Activate
ThisWorkbook.Worksheets("title").Rows(4).Copy

rmsht.Activate
rmsht.Rows(2).PasteSpecial xlPasteAll

idx = 3
For i = 4 To soprmsht.Range("G4").End(xlDown).Row
    If soprmsht.Cells(i, 7).Font.Color = vbBlack Then
        rmsht.Cells(idx, 1) = soprmsht.Cells(i, 4)
        rmsht.Cells(idx, 2) = soprmsht.Cells(i, 8)
        rmsht.Cells(idx, 3) = soprmsht.Cells(i, 7)
        rmsht.Cells(idx, 4) = soprmsht.Cells(i, 11)
        rmsht.Cells(idx, 5) = soprmsht.Cells(i, 5)
        rmsht.Cells(idx, 6) = soprmsht.Cells(i, 13)
        'rmsht.Cells(idx, 6) = 5600
        idx = idx + 1
    End If
Next

rmsht.Columns("F:F").NumberFormat = "MM-DD-YYYY"

For Each sht In prevbk.Worksheets
    If sht.Name Like "RM*" Then
        Set tmpsht = sht
    End If
Next


Dim c As Variant
idx = 3
For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    
    On Error Resume Next
    
    c = Application.VLookup(rmsht.Cells(i, 2).Value, tmpsht.Range("B:O"), 6, False)
    If Not (c Is Nothing) Then
        rmsht.Cells(i, 7).Value = c
    End If
    c = Application.VLookup(rmsht.Cells(i, 2).Value, tmpsht.Range("B:O"), 14, False)
    If Not (c Is Nothing) Then
        rmsht.Cells(i, 8).Value = c
    End If
Next

Set tmpsht = mvmtbk.Worksheets("RM")

Set pvt = tmpsht.PivotTables("PivotTable1")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(rmsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 3), tmpsht.Cells(tmpsht.Cells(tstartrow, 3).End(xlDown).Row, 4)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           rmsht.Cells(i, 9) = c
        End If
    End If
    
 Next
 
 Set pvt = tmpsht.PivotTables("PivotTable2")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(rmsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 9), tmpsht.Cells(tmpsht.Cells(tstartrow, 9).End(xlDown).Row, 10)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           rmsht.Cells(i, 10) = c
        End If
    End If
    
 Next

Set pvt = tmpsht.PivotTables("PivotTable3")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(rmsht.Cells(i, 3), tmpsht.Range(tmpsht.Cells(tstartrow, 15), tmpsht.Cells(tmpsht.Cells(tstartrow, 15).End(xlDown).Row, 16)), 2, False)
    If Not (c Is Nothing) Then
        If IsError(c) = False Then
           rmsht.Cells(i, 11) = c
        End If
    End If
Next



'vlookup AUM data

For Each sht In prevbk.Worksheets
    If sht.Name Like "RM*" Then
        Set tmpsht = sht
    End If
Next


'USE FOR LOOP. put in monthly movement data from prev bk

monthCt = Month(DateValue("01 " & Left(monthstr, 3) & " 2015"))
For j = 1 To monthCt - 2
    For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
        
        On Error Resume Next
        
        c = Application.VLookup(rmsht.Cells(i, 2).Value, tmpsht.Range("B:AB"), 16 + j, False)
        If Not (c Is Nothing) Then
            rmsht.Cells(i, 17 + j).Value = c
            
        End If
    Next
Next


rmsht.Range("L1").Value = 367
For i = 3 To rmsht.Range("A4").End(xlDown).Row
    rmsht.Range("L" & CStr(i)).Formula = "=sum(I" & CStr(i) & ":J" & CStr(i) & ") -K" & CStr(i)
Next

'Paste current AUM movement
For i = 3 To rmsht.Range("A4").End(xlDown).Row
    rmsht.Cells(i, 16 + monthCt) = rmsht.Cells(i, 12)
Next

For i = 3 To rmsht.Range("A4").End(xlDown).Row
    rmsht.Cells(i, 30).FormulaR1C1 = "=SUM(RC[-11] : RC[-1])"
Next

For i = 3 To rmsht.Range("A4").End(xlDown).Row
    rmsht.Cells(i, 31).FormulaR1C1 = "=sum(RC[-1], RC[-13])"
Next

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    On Error Resume Next
    
    c = Application.VLookup(rmsht.Cells(i, 2).Value, tmpsht.Range("B:AF"), 31, False)
    If Not (c Is Nothing) Then
        rmsht.Cells(i, 32).Value = c
        
    End If
Next

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    rmsht.Range("AG" & CStr(i)).Formula = "=IF((AE" & CStr(i) & "<AF" & CStr(i) & "),""xxx"",""Yes"")"
    rmsht.Range("AG" & CStr(i)).Interior.Color = vbYellow
Next

For i = 3 To rmsht.Cells(3, 3).End(xlDown).Row
    rmsht.Range("M" & CStr(i)).Formula = "=IF(AND(L" & CStr(i) & ">1000000, AG" & CStr(i) & "=""YES""), ""LOAD"","""")"
    If rmsht.Range("M" & CStr(i)) = "LOAD" Then
        rmsht.Rows(i).Font.Color = vbRed
        isRMLoad = 1
    End If
    rmsht.Range("N" & CStr(i)).Formula = "=IF(L" & CStr(i) & "<0,(L" & CStr(i) & "/1000000)*$L$1,IF(M" & CStr(i) & "=""LOAD"",(L" & CStr(i) & "/1000000)*$L$1,0))"
    rmsht.Range("O" & CStr(i)).Formula = "=IF((H" & CStr(i) & "+N" & CStr(i) & ")<G" & CStr(i) & ",G" & CStr(i) & ",(H" & CStr(i) & "+N" & CStr(i) & "))"
Next

rmsht.Columns("A:C").AutoFit

End Sub


Sub CPCPage()

isCPCLoad = 0
For Each sht In sopbk.Worksheets
  
    If sht.Name Like "*RM" Then
        Set sopcpcsht = sht
    End If
Next

Set cpcsht = outbk.Worksheets(3)
cpcsht.Name = "CPC"

ThisWorkbook.Worksheets("title").Activate
ThisWorkbook.Worksheets("title").Rows(7).Copy

cpcsht.Activate
cpcsht.Rows(2).PasteSpecial xlPasteAll

idx = 3
For i = 4 To sopcpcsht.Range("G4").End(xlDown).Row
    If sopcpcsht.Cells(i, 7).Font.Color <> vbBlack Then
    cpcsht.Cells(idx, 1) = sopcpcsht.Cells(i, 4)
    cpcsht.Cells(idx, 2) = sopcpcsht.Cells(i, 8)
    cpcsht.Cells(idx, 3) = sopcpcsht.Cells(i, 7)
    cpcsht.Cells(idx, 4) = sopcpcsht.Cells(i, 11)
    idx = idx + 1
    End If
Next

cpcsht.Columns("E:E").NumberFormat = "MM-DD-YYYY"

For Each sht In prevbk.Worksheets
    If sht.Name Like "CPC*" Then
        Set tmpsht = sht
    End If
Next


Dim c As Variant
idx = 3
For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    
    On Error Resume Next
    
    c = Application.VLookup(cpcsht.Cells(i, 2).Value, tmpsht.Range("B:O"), 13, False)
    If Not (c Is Nothing) Then
        cpcsht.Cells(i, 7).Value = c
    End If
    c = Application.VLookup(cpcsht.Cells(i, 2).Value, tmpsht.Range("B:O"), 5, False)
    If Not (c Is Nothing) Then
        cpcsht.Cells(i, 6).Value = c
    End If
Next

Set tmpsht = mvmtbk.Worksheets("CPC")

Set pvt = tmpsht.PivotTables("PivotTable1")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(cpcsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 3), tmpsht.Cells(tmpsht.Cells(tstartrow, 3).End(xlDown).Row, 4)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           cpcsht.Cells(i, 8) = c
        End If
    End If
    
 Next
 
 Set pvt = tmpsht.PivotTables("PivotTable2")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(cpcsht.Cells(i, 3).Value, tmpsht.Range(tmpsht.Cells(tstartrow, 9), tmpsht.Cells(tmpsht.Cells(tstartrow, 9).End(xlDown).Row, 10)), 2, False)
   ' MsgBox IsError(c)
    If (Not (c Is Nothing)) Then
        If IsError(c) = False Then
           cpcsht.Cells(i, 9) = c
        End If
    End If
    
 Next

Set pvt = tmpsht.PivotTables("PivotTable3")

tstartrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.Count).Row + 3

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    c = Nothing
    c = Application.VLookup(cpcsht.Cells(i, 3), tmpsht.Range(tmpsht.Cells(tstartrow, 15), tmpsht.Cells(tmpsht.Cells(tstartrow, 15).End(xlDown).Row, 16)), 2, False)
    If Not (c Is Nothing) Then
        If IsError(c) = False Then
           cpcsht.Cells(i, 10) = c
        End If
    End If
Next



'vlookup AUM data

For Each sht In prevbk.Worksheets
    If sht.Name Like "CPC*" Then
        Set tmpsht = sht
    End If
Next


'USE FOR LOOP. put in monthly movement data from prev bk

monthCt = Month(DateValue("01 " & Left(monthstr, 3) & " 2015"))
For j = 1 To monthCt - 2
    For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
        
        On Error Resume Next
        
        c = Application.VLookup(cpcsht.Cells(i, 2).Value, tmpsht.Range("B:AB"), 15 + j, False)
        If Not (c Is Nothing) Then
            cpcsht.Cells(i, 16 + j).Value = c
            
        End If
    Next
Next


cpcsht.Range("K1").Value = 367
For i = 3 To cpcsht.Range("A4").End(xlDown).Row
    cpcsht.Range("K" & CStr(i)).Formula = "=sum(H" & CStr(i) & ":I" & CStr(i) & ") -J" & CStr(i)
Next

'Paste current AUM movement
For i = 3 To cpcsht.Range("A4").End(xlDown).Row
    cpcsht.Cells(i, 15 + monthCt) = cpcsht.Cells(i, 11)
Next

For i = 3 To cpcsht.Range("A4").End(xlDown).Row
    cpcsht.Cells(i, 29).FormulaR1C1 = "=SUM(RC[-11] : RC[-1])"
Next

For i = 3 To cpcsht.Range("A4").End(xlDown).Row
    cpcsht.Cells(i, 30).FormulaR1C1 = "=sum(RC[-1], RC[-13])"
Next

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    On Error Resume Next
    
    c = Application.VLookup(cpcsht.Cells(i, 2).Value, tmpsht.Range("B:AE"), 30, False)
    If Not (c Is Nothing) Then
        cpcsht.Cells(i, 31).Value = c
        cpcsht.Cells(i, 5).Value = c
    End If
Next

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    cpcsht.Range("AF" & CStr(i)).Formula = "=IF((AD" & CStr(i) & "<AE" & CStr(i) & "),""xxx"",""Yes"")"
    cpcsht.Range("AF" & CStr(i)).Interior.Color = vbYellow
Next

For i = 3 To cpcsht.Cells(3, 3).End(xlDown).Row
    cpcsht.Range("L" & CStr(i)).Formula = "=IF(AND(K" & CStr(i) & ">1000000, AF" & CStr(i) & "=""YES""), ""LOAD"","""")"
    If cpcsht.Cells(i, 7).Value = 999999 Then
      
        cpcsht.Range("L" & CStr(i)) = ""
    End If
    If cpcsht.Range("L" & CStr(i)) = "LOAD" Then
        cpcsht.Rows(i).Font.Color = vbRed
        isCPCLoad = 1
    End If
    cpcsht.Range("M" & CStr(i)).Formula = "=IF(K" & CStr(i) & "<0,(K" & CStr(i) & "/1000000)*$K$1,IF(L" & CStr(i) & "=""LOAD"",(K" & CStr(i) & "/1000000)*$K$1,0))"
    cpcsht.Range("N" & CStr(i)).Formula = "=IF((G" & CStr(i) & "+M" & CStr(i) & ")<F" & CStr(i) & ",F" & CStr(i) & ",(G" & CStr(i) & "+M" & CStr(i) & "))"
Next

cpcsht.Columns("A:C").AutoFit

End Sub

Sub UpdateMonthTitle()
    Dim monthArr(12) As String
    
    Set sht = ThisWorkbook.Worksheets("title")
    sht.Range("N1") = monthstr + Right(sht.Range("G1"), Len(sht.Range("G1")) - 5)
    sht.Range("G1") = getPrevMonth + Right(sht.Range("G1"), Len(sht.Range("G1")) - 5)
    sht.Range("K1") = getPrevMonth + Right(sht.Range("K1"), Len(sht.Range("K1")) - 5)
    
    sht.Range("O4") = monthstr + Right(sht.Range("O4"), Len(sht.Range("O4")) - 5)
    sht.Range("H4") = getPrevMonth + Right(sht.Range("H4"), Len(sht.Range("H4")) - 5)
    sht.Range("L4") = getPrevMonth + Right(sht.Range("L4"), Len(sht.Range("L4")) - 5)
    
    sht.Range("N7") = monthstr + Right(sht.Range("G7"), Len(sht.Range("G7")) - 5)
    sht.Range("G7") = getPrevMonth + Right(sht.Range("G7"), Len(sht.Range("G7")) - 5)
    sht.Range("K7") = getPrevMonth + Right(sht.Range("K7"), Len(sht.Range("K7")) - 5)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
