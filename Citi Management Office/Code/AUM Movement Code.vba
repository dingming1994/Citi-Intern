Public saveToFolder As String



Function getReport_Mul(tbook As Workbook, RMname As String, mStr As String) As Integer
    getReport_Mul = 0
    Set wkbk = ActiveWorkbook
    'create new sheet for preparing data
    
    Set readysht = wkbk.Worksheets.Add
    readysht.Name = "AUM_" + RMname + "_" + mStr
    
    'Get the category of user. assume RM here
    Dim rank As String
    rank = getRank(RMname)
    If rank = "#" Then
        'MsgBox ("User Not Found")
        getReport_Mul = -1
        readysht.Cells(1, 1).Value = "No movement for this month"
        Application.DisplayAlerts = False
        readysht.Copy after:=tbook.Worksheets(tbook.Worksheets.count)
        Application.DisplayAlerts = True
        Exit Function
    End If
    If rank <> "RM" And rank <> "PB" And rank <> "G3" Then
        'MsgBox ("Undefined User Rank")
        getReport_Mul = -2
        readysht.Cells(1, 1).Value = "No movement for this month"
        Application.DisplayAlerts = False
        readysht.Copy after:=tbook.Worksheets(tbook.Worksheets.count)
        Application.DisplayAlerts = True
        Exit Function
    End If
    If rank = "G3" Then
        rank = "CPC"
    End If
    Set sht = wkbk.Worksheets(rank)
    sht.Activate
    Dim pvt As PivotTable
    Dim numOfRow As Long
    Dim resultRowCount As Long
    resultRowCount = 0
    Dim tot As Double
    Dim rowIdx(4) As Long ' array to denote the end row number of each category of record
    Dim fieldName As String ' RM Name RCAO_C or _P depends on the table name
    Dim flag As Integer   ' flag for whether found user name
    For tableidx = 1 To 3
        Set pvt = sht.PivotTables("PivotTable" + CStr(tableidx))
        pvt.RowRange.Select
        If tableidx <> 3 Then
            fieldName = "RM name_RCAO_C"
        Else
            fieldName = "RM name_RCAO_P"
        End If
        flag = 0
        
        pvt.PivotFields(fieldName).DataRange.Select
        For Each cell In Selection
            If cell.Value = RMname Then
            cell.Offset(0, 1).Select
            flag = 1
            End If
        Next
        
     
        If flag = 1 Then

        Selection.ShowDetail = True
        ActiveSheet.Name = "TEMP"
        numOfRow = ActiveSheet.UsedRange.Rows.count
        ActiveSheet.Rows("1:" & CStr(numOfRow)).Copy
        readysht.Activate
        readysht.Cells(resultRowCount + 1, 1).Select
        readysht.Paste
        resultRowCount = ActiveSheet.UsedRange.Rows.count
        End If
        rowIdx(tableidx + 1) = resultRowCount
        Application.DisplayAlerts = False
        For Each st In wkbk.Worksheets
            If st.Name = "TEMP" Then
                st.Delete
            End If
        Next
        Application.DisplayAlerts = True
        sht.Activate
    Next
    
    
    'insert column
    readysht.Activate
    readysht.Range("A1").EntireColumn.Insert
    Dim str As String
    Dim mul As Integer
    For i = 1 To 3
        Select Case i:
        Case 1:
            str = "inflow"
            mul = 1
        Case 2:
            str = "upgrade"
            mul = 1
        Case 3:
            str = "outflow"
            mul = -1
        End Select
        For j = rowIdx(i) + 2 To rowIdx(i + 1)
            readysht.Cells(j, 1).Value = str
            tot = tot + mul * readysht.Cells(j, 10).Value
        Next
    Next
    If tot <> 0 Then
    readysht.Cells(rowIdx(4) + 1, 10).Formula = "=" + CStr(tot)
    readysht.Columns("J:J").NumberFormat = "#,##0"
    Else
    readysht.Cells(1, 1).Value = "No movement for this month"
    End If
    readysht.Activate
    totsht = tbook.Worksheets.count
    Application.DisplayAlerts = False
    readysht.Copy after:=tbook.Worksheets(totsht)
    Application.DisplayAlerts = True
    
    'Dim FName As String
    
    'Application.DisplayAlerts = False
    'With ActiveWorkbook
    '    .SaveAs Filename:="", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    '    .Close 0
    'End With
    'ActiveSheet.Delete
    'Application.DisplayAlerts = True
    
    End Function


Function getReport(RMname As String, path As String, mStr As String) As Integer
    getReport = 0
    Set wkbk = ActiveWorkbook
    'create new sheet for preparing data
    
    Set readysht = wkbk.Worksheets.Add
    readysht.Name = "AUM_" + RMname
    
    'Get the category of user. assume RM here
    Dim rank As String
    rank = getRank(RMname)
    If rank = "#" Then
        MsgBox ("User Not Found")
        getReport = -1
        Exit Function
    End If
    If rank <> "RM" And rank <> "PB" And rank <> "G3" Then
        MsgBox ("Undefined User Rank")
        getReport = -2
        Exit Function
    End If
    If rank = "G3" Then
        rank = "CPC"
    End If
    Set sht = wkbk.Worksheets(rank)
    sht.Activate
    Dim pvt As PivotTable
    Dim numOfRow As Long
    Dim resultRowCount As Long
    resultRowCount = 0
    Dim tot As Double
    Dim rowIdx(4) As Long ' array to denote the end row number of each category of record
    Dim fieldName As String ' RM Name RCAO_C or _P depends on the table name
    Dim flag As Integer   ' flag for whether found user name
    For tableidx = 1 To 3
        Set pvt = sht.PivotTables("PivotTable" + CStr(tableidx))
        pvt.RowRange.Select
        If tableidx <> 3 Then
            fieldName = "RM name_RCAO_C"
        Else
            fieldName = "RM name_RCAO_P"
        End If
        flag = 0
        
        pvt.PivotFields(fieldName).DataRange.Select
        For Each cell In Selection
            If cell.Value = RMname Then
            cell.Offset(0, 1).Select
            flag = 1
            End If
        Next
        
     
        If flag = 1 Then

        Selection.ShowDetail = True
        ActiveSheet.Name = "TEMP"
        numOfRow = ActiveSheet.UsedRange.Rows.count
        ActiveSheet.Rows("1:" & CStr(numOfRow)).Copy
        readysht.Activate
        readysht.Cells(resultRowCount + 1, 1).Select
        readysht.Paste
        resultRowCount = ActiveSheet.UsedRange.Rows.count
        End If
        rowIdx(tableidx + 1) = resultRowCount
        Application.DisplayAlerts = False
        For Each st In wkbk.Worksheets
            If st.Name = "TEMP" Then
                st.Delete
            End If
        Next
        Application.DisplayAlerts = True
        sht.Activate
    Next
    
    
    'insert column
    readysht.Activate
    readysht.Range("A1").EntireColumn.Insert
    Dim str As String
    Dim mul As Integer
    For i = 1 To 3
        Select Case i:
        Case 1:
            str = "inflow"
            mul = 1
        Case 2:
            str = "upgrade"
            mul = 1
        Case 3:
            str = "outflow"
            mul = -1
        End Select
        For j = rowIdx(i) + 2 To rowIdx(i + 1)
            readysht.Cells(j, 1).Value = str
            tot = tot + mul * readysht.Cells(j, 10).Value
        Next
    Next
    If tot <> 0 Then
    readysht.Cells(rowIdx(4) + 1, 10).Formula = "=" + CStr(tot)
    readysht.Columns("J:J").NumberFormat = "#,##0"
    Else
    readysht.Cells(1, 1).Value = "No movement for this month"
    End If
    readysht.Activate
    readysht.Copy
    Dim Fname As String
    Fname = path + "\" + "AUM Movement(" + mStr + ") - " + RMname + " " + getTotStr(CLng(tot)) + ".xlsx"
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=Fname, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        .Close 0
    End With
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
    
End Function

Function getRank(nm As String) As String
    Set sht = ActiveWorkbook.Sheets("MOVEMENT 1&2")
    Dim numOfRow As Long
    numOfRow = sht.UsedRange.Rows.count
    For i = 2 To numOfRow
        If CStr(sht.Cells(i, 14).Value) = nm Then
            getRank = CStr(sht.Cells(i, 15).Value)
            Exit Function
        End If
        If CStr(sht.Cells(i, 18).Value) = nm Then
            getRank = CStr(sht.Cells(i, 19).Value)
            Exit Function
        End If
    Next
    getRank = "#"
        
End Function

Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

With fldr
    .Title = "Select a Folder to Save Files"
    .AllowMultiSelect = False
    .InitialFileName = Application.ActiveWorkbook.path
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function


Private Sub CombineButton_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CombineButton_click()


Dim strpath1 As String
Dim strpath2 As String
Dim numOfRow1 As Long
Dim numOfRow2 As Long

If saveToFolder = "" Then
MsgBox "Please select the saving directory first."
Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the monthly AUM 1 & 2 files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\"
intChoice = Application.FileDialog(msoFileDialogOpen).Show

'Application.ScreenUpdating = False
If intChoice <> 0 Then
    'get the file path selected by the user
    If Application.FileDialog(msoFileDialogOpen).SelectedItems.count <> 2 Then
        MsgBox "Please select both monthly AUM 1 & 2 files!"
        Exit Sub
    Else
        strpath1 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        strpath2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(2)
        If Not (strpath1 Like "*Monthly_AUM_####_Movement_1.*" Or strpath1 Like "*Monthly_AUM_####_Movement_2.*") Then
            MsgBox "Please Select only monthly AUM 1 & 2 files!"
            Exit Sub
        End If
        If Not (strpath2 Like "*Monthly_AUM_####_Movement_1.*" Or strpath2 Like "*Monthly_AUM_####_Movement_2.*") Then
            MsgBox "Please Select only monthly AUM 1 & 2 files!"
            Exit Sub
        End If
        
        '--
        Dim strpath_1 As String
        Dim strpath_2 As String
        
        MsgBox "Please select the SOP Masterlist for current month"
        
        Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP masterlist for current month"
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
        Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\(SK) SOP clean up files\"
        intChoice = Application.FileDialog(msoFileDialogOpen).Show
        
        If intChoice <> 0 Then
            strpath_1 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
            If Not strpath_1 Like "*SOP Masterlist*" Then
                MsgBox "Please Select The Correct SOP Masterlist File!"
                Exit Sub
            End If
        Else: Exit Sub
        End If
        MsgBox "Please select the SOP Masterlist for previous month"
        Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP masterlist for previous month"
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
        Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\(SK) SOP clean up files\"
        intChoice = Application.FileDialog(msoFileDialogOpen).Show
        
        If intChoice <> 0 Then
            strpath_2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
            If Not strpath_2 Like "*SOP Masterlist*" Then
                MsgBox "Please Select The Correct SOP Masterlist File!"
                Exit Sub
            End If
        Else: Exit Sub
        End If
        
        Dim strpath_3 As String
        
        '--
        
        Dim saveToPath As String
        saveToPath = left(strpath1, InStrRev(strpath1, ".") - 2) & " 1&2 " & Format(Date, "yyyy.mm.dd") & " (MTM)" & right(strpath1, Len(strpath1) - InStrRev(strpath1, ".") + 1)
        saveToPath = saveToFolder & "\" & right(saveToPath, Len(saveToPath) - InStrRev(saveToPath, "\"))
        strpath_3 = saveToPath
        Workbooks.Open (strpath1)
        Set wkbk1 = ActiveWorkbook
        Workbooks.Open (strpath2)
        Set wkbk2 = ActiveWorkbook
        numOfRow1 = wkbk1.Worksheets(1).UsedRange.Rows.count
        numOfRow2 = wkbk2.Worksheets(1).UsedRange.Rows.count
        wkbk1.Worksheets(1).Rows("2:" + CStr(numOfRow1)).Copy
        wkbk2.Worksheets(1).Rows(CStr(numOfRow2 + 1) + ":" + CStr(numOfRow1 + numOfRow2 - 1)).PasteSpecial xlPasteValues
        wkbk2.Worksheets(1).Name = "Movement 1&2"

        Application.DisplayAlerts = False
        With wkbk2
            .SaveAs Filename:=saveToPath, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
            .Close
        End With
        wkbk1.Close
        
        Application.DisplayAlerts = True
        FormattingSheet saveToPath
        SOP_In strpath_1, strpath_2, strpath_3
        CPCpage
        RMpage
        PBpage
        makeSecondTable ("RM")
        makeSecondTable ("PB")
        makeSecondTable ("CPC")
        Application.DisplayAlerts = False
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
    End If
Else: Exit Sub
End If
'Application.ScreenUpdating = True
'FormattingSheet saveToPath
MsgBox "AUM Movement 1&2 File generation completed!"
End Sub
Sub FormattingSheet(path As String)
    Workbooks.Open (path)
    Set sht = ActiveWorkbook.Worksheets(1)
    'MsgBox sht.Name
    Dim usedCol As Integer
    Dim usedRow As Long
    usedCol = sht.UsedRange.Columns.count
    usedRow = sht.UsedRange.Rows.count
    sht.Cells(1, usedCol + 1).Value = "Transfer_Internal bet RM"
    sht.Cells(1, usedCol + 2).Value = "BBD_RCAO_C"
    sht.Cells(1, usedCol + 3).Value = "BM_RCAO_C"
    sht.Cells(1, usedCol + 4).Value = "RM name_RCAO_C"
    sht.Cells(1, usedCol + 5).Value = "RM Rank_C"
    sht.Cells(1, usedCol + 6).Value = "BBD_RCAO_P"
    sht.Cells(1, usedCol + 7).Value = "BM_RCAO_P"
    sht.Cells(1, usedCol + 8).Value = "RM name_RCAO_P"
    sht.Cells(1, usedCol + 9).Value = "RM Rank_P"
    sht.Cells(1, usedCol + 10).Value = "Transfer_Internal (RM_C = RM_P)"
    sht.Cells(1, usedCol + 11).Value = "BBM transfer internal (BM_C =BM_P)"
    For i = 2 To usedRow
        If sht.Cells(i, 6).Value = 0 And sht.Cells(i, 5).Value = 0 Then
            sht.Cells(i, 11).Value = 1
        ElseIf sht.Cells(i, 6).Value = 1 Or sht.Cells(i, 5).Value = 1 Then
            sht.Cells(i, 11).Value = 0
        Else
            sht.Cells(i, 11).Value = "ERROR"
        End If
    
    Next
   Application.DisplayAlerts = False
   ActiveWorkbook.Save
   ActiveWorkbook.Close
   Application.DisplayAlerts = True
    
End Sub


Private Sub MulMonReport_Click()

Dim saveToPath As String
Dim strpath() As String
Dim fileCount As Integer
Dim RMname As String
ReDim strpath(1)

saveToPath = GetFolder()
If saveToPath = "" Then
 Exit Sub
 End If


Application.FileDialog(msoFileDialogOpen).Title = "Select the AUM Movement 1&2 file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
intChoice = Application.FileDialog(msoFileDialogOpen).Show



If intChoice <> 0 Then
    'get the file path selected by the user
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.count
    ReDim strpath(fileCount)
    For i = 1 To fileCount
        strpath(i) = Application.FileDialog(msoFileDialogOpen).SelectedItems(i)
    Next
    If checkDulFiles(strpath, fileCount) = 1 Then
        MsgBox ("Please Don't Select Files in Same Month")
        Exit Sub
    ElseIf checkDulFiles(strpath, fileCount) = -1 Then
        MsgBox ("Please Select the Correct Files")
        Exit Sub
    End If
Else: Exit Sub
End If

Set myfrm = UserForm1
'myfrm.Caption = "Select Name"
myfrm.Show
RMname = myfrm.searchName
Unload myfrm

If RMname = "NONE" Then
    Exit Sub
End If
'RMName = InputBox(Prompt:="Please input the FULL name of sales person")

Application.ScreenUpdating = False
'create a new workbook to save info
 Dim tbook As Workbook
 Set targetBook = Workbooks.Add
 Set tbook = targetBook
 Dim MonStr As String
For i = 1 To fileCount

    Application.DisplayAlerts = False
    Workbooks.Open strpath(i)
    Set wkbk = ActiveWorkbook
    Application.DisplayAlerts = True
    
    MonStr = MonStr + " " + getMonth(strpath(i))
    getReport_Mul tbook, RMname, getMonth(strpath(i))
    Application.DisplayAlerts = False
    wkbk.Close
    Application.DisplayAlerts = True
Next

tbook.Activate
Application.DisplayAlerts = False
'delete dummy sheet
Dim status As Integer
status = 1
If tbook.Worksheets.count = 3 Then
    MsgBox ("User not found in all files")
    tbook.Close
    status = 0
    Exit Sub
End If
For i = 1 To 3
    tbook.Worksheets(1).Delete
Next
Dim Fname As String
Fname = saveToPath + "\" + "AUM Movement(" + MonStr + ") - " + RMname + ".xlsx"
    
    With tbook
        .SaveAs Filename:=Fname, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        .Close 0
    End With
    

Workbooks.Open Fname, corruptload:=xlRepairFile
 
    With ActiveWorkbook
        .SaveAs Filename:=Fname, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        
        .Close 0
    End With

Application.DisplayAlerts = True

Application.ScreenUpdating = True
If status = 1 Then
    MsgBox "AUM Movement Report is Done"
End If

End Sub



Private Sub PBRMReport_Click()
Dim saveToPath As String
Dim strpath As String
Dim RMname As String
saveToPath = GetFolder()
If saveToPath = "" Then
 Exit Sub
 End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the AUM Movement 1&2 file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

Set myfrm = UserForm1
'myfrm.Caption = "Select Name"
myfrm.Show

RMname = myfrm.searchName
Unload myfrm

If RMname = "NONE" Then
    Exit Sub
End If
'RMName = InputBox(Prompt:="Please input the FULL name of sales person")

Application.ScreenUpdating = False

Application.DisplayAlerts = False
Workbooks.Open strpath
Set wkbk = ActiveWorkbook
Application.DisplayAlerts = True

Dim status As Integer
status = getReport(RMname, saveToPath, getMonth(strpath))
Application.DisplayAlerts = False
wkbk.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
If status <> -1 And status <> -2 Then
    MsgBox "AUM Movement Report is Done"
End If

End Sub


' to get the MMM expresion of a month
' input is the whole path of file (string)
Function getMonth(strpath As String) As String
    Dim tmpstr As String
    getMonth = "###"
    tmpstr = mid(strpath, InStrRev(strpath, "\") + 1, 18)
    tmpstr = left(tmpstr, 16)
    tmpstr = right(tmpstr, 2)  ' this is the 2 digit string for month.
    Select Case tmpstr:
    Case "01":
        getMonth = "Jan"
    Case "02":
        getMonth = "Feb"
    Case "03":
        getMonth = "Mar"
    Case "04":
        getMonth = "Apr"
    Case "05":
        getMonth = "May"
    Case "06":
        getMonth = "Jun"
    Case "07":
        getMonth = "Jul"
    Case "08":
        getMonth = "Aug"
    Case "09":
        getMonth = "Sep"
    Case "10":
        getMonth = "Oct"
    Case "11":
        getMonth = "Nov"
    Case "12":
        getMonth = "Dec"
    Case Else:
        getMonth = "###"
    End Select

End Function

Function getTotStr(tot As Long) As String
    If tot = 0 Then
        getTotStr = "0"
        Exit Function
    End If
    Dim temp As Double
    Dim leng As Integer
    leng = Int(Log(Abs(tot)) / Log(10))
    temp = tot / Application.WorksheetFunction.Power(10, leng)
    temp = Round(temp, 2)
    temp = temp * Application.WorksheetFunction.Power(10, leng)
    ' by now we hv temp is the first 3 digit of total
    If Abs(temp) > 1000000 Then
        getTotStr = CStr(temp / 1000000) + "MM"
    ElseIf Abs(temp) > 1000 Then
        getTotStr = CStr(temp / 1000) + "M"
    Else
        getTotStr = CStr(temp)
    End If
    
End Function

'check if files from same month are selected, 1 if yes, 0 if no
Function checkDulFiles(strp() As String, cnt As Integer) As Integer
    checkDulFiles = 0
    Dim iMon As String
    Dim jMon As String

    For i = 1 To cnt
        For j = 1 To cnt
            iMon = getMonth(CStr(strp(i)))
            jMon = getMonth(CStr(strp(j)))
            If iMon = "###" Or jMon = "###" Then
                checkDulFiles = -1
                Exit Function
            End If
            If i <> j And iMon = jMon Then
                checkDulFiles = 1
                Exit Function
            End If
        Next
    Next
    
End Function


Sub test112()
ActiveWorkbook.Sheets(1).Delete
ActiveWorkbook.Sheets(1).Delete
End Sub

Private Sub SOP_In(strpath1 As String, strpath2 As String, strpath3 As String)

Workbooks.Open strpath3
Set tarbk = ActiveWorkbook
Workbooks.Open strpath1, UpdateLinks:=False
Set cpbk = ActiveWorkbook
Set tsht = cpbk.Worksheets("SOP")

Dim ColIdxArr(5) As Integer
Dim ColTitle As Variant
ColTitle = Array("RCAO", "RM", "Mgr", "BBD", "Type")
For i = 1 To 5
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 4
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 5
    If ColIdxArr(k) <> 0 Then
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
    End If
Next
tsht.Activate
Y.Select
Selection.Copy
tarbk.Activate
tarbk.Worksheets.Add
Set tarsht = ActiveSheet
tarsht.Name = "SOP_C"


tarsht.Cells(1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set tsht = cpbk.Worksheets("ROE & others")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 3
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
Next
tsht.Activate
Y.Select
Selection.Copy
tarsht.Activate

tarsht.Cells(tarsht.UsedRange.Rows.count + 1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set tsht = cpbk.Worksheets("GEB & Remote codes")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 3
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
Next
tsht.Activate
Y.Select
Selection.Copy
tarsht.Activate

tarsht.Cells(tarsht.UsedRange.Rows.count + 1, 1).PasteSpecial xlPasteValues
Dim ct As Integer
ct = 2
Dim num As Integer
num = tarsht.UsedRange.Rows.count
For i = 2 To num
    If IsError(tarsht.Cells(ct, 1).Value) Or IsError(tarsht.Cells(ct, 2).Value) Or IsError(tarsht.Cells(ct, 3).Value) Or IsError(tarsht.Cells(ct, 4).Value) Then
        tarsht.Rows(ct).Delete
    ElseIf tarsht.Cells(ct, 1).Value = "RM" Or tarsht.Cells(ct, 1).Value = "" Or tarsht.Cells(ct, 2).Value = "" Or tarsht.Cells(ct, 3).Value = "" Or tarsht.Cells(ct, 4).Value = "" Then
        tarsht.Rows(ct).Delete
    Else
        If tarsht.Cells(ct, 5).Value = "" Then tarsht.Cells(ct, 5).Value = "0"
        ct = ct + 1
    End If
Next

Application.DisplayAlerts = False
cpbk.Close
Application.DisplayAlerts = True
tarsht.Activate
sortRCAO

'+++++++++++++++++++++++++++++++++++
Workbooks.Open strpath2, UpdateLinks:=False
Set cpbk = ActiveWorkbook
Set tsht = cpbk.Worksheets("SOP")



For i = 1 To 5
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 4
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 5
    If ColIdxArr(k) <> 0 Then
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
    End If
Next
tsht.Activate
Y.Select
Selection.Copy
tarbk.Activate
tarbk.Worksheets.Add
Set tarsht = ActiveSheet
tarsht.Name = "SOP_P"


tarsht.Cells(1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set tsht = cpbk.Worksheets("ROE & others")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 3
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
Next
tsht.Activate
Y.Select
Selection.Copy
tarsht.Activate

tarsht.Cells(tarsht.UsedRange.Rows.count + 1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set tsht = cpbk.Worksheets("GEB & Remote codes")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To tsht.UsedRange.Columns.count
    For j = 0 To 3
        If tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, tsht.UsedRange.Columns(ColIdxArr(k)))
Next
tsht.Activate
Y.Select
Selection.Copy
tarsht.Activate

tarsht.Cells(tarsht.UsedRange.Rows.count + 1, 1).PasteSpecial xlPasteValues
ct = 2
num = tarsht.UsedRange.Rows.count
For i = 2 To num
    If IsError(tarsht.Cells(ct, 1).Value) Or IsError(tarsht.Cells(ct, 2).Value) Or IsError(tarsht.Cells(ct, 3).Value) Or IsError(tarsht.Cells(ct, 4).Value) Then
        tarsht.Rows(ct).Delete
    ElseIf tarsht.Cells(ct, 1).Value = "RM" Or tarsht.Cells(ct, 1).Value = "" Or tarsht.Cells(ct, 2).Value = "" Or tarsht.Cells(ct, 3).Value = "" Or tarsht.Cells(ct, 4).Value = "" Then
        tarsht.Rows(ct).Delete
    Else
        If tarsht.Cells(ct, 5).Value = "" Then tarsht.Cells(ct, 5).Value = "0"
        ct = ct + 1
    End If
Next

Application.DisplayAlerts = False
cpbk.Close
Application.DisplayAlerts = True

tarsht.Activate
sortRCAO
tarbk.Activate

SOPFillIn



End Sub
Sub sortRCAO()
ActiveSheet.Columns("A:E").sort key1:=ActiveSheet.Range("B2"), order1:=xlAscending, Header:=xlYes
End Sub

Sub SOPFillIn()
    Set sht = ActiveWorkbook.Worksheets("Movement 1&2")
    Dim csht As Worksheet
    Dim psht As Worksheet
    Set csht = ActiveWorkbook.Worksheets("SOP_C")
    Set psht = ActiveWorkbook.Worksheets("SOP_P")
    'csht.Range("A:E").sort key1:=csht.Range("B2"), order1:=xlAscending, Header:=xlYes
    'psht.Range("A:E").sort key1:=psht.Range("B2"), order1:=xlAscending, Header:=xlYes

    Dim flag As Integer
    With sht
        For i = 2 To .UsedRange.Rows.count
            flag = binarySearch(csht, 2, csht.UsedRange.Rows.count, CLng(.Cells(i, 2).Value))
            
            If flag <> 0 Then
                    .Cells(i, 12).Value = csht.Cells(flag, 4).Value
                    .Cells(i, 13).Value = csht.Cells(flag, 3).Value
                    .Cells(i, 14).Value = csht.Cells(flag, 1).Value
                    .Cells(i, 15).Value = csht.Cells(flag, 5).Value
                End If
            
            If flag = 0 Then
                .Cells(i, 12).Value = "#N/A"
                .Cells(i, 13).Value = "#N/A"
                .Cells(i, 14).Value = "#N/A"
                .Cells(i, 15).Value = "#N/A"
            End If
        Next
        
        For i = 2 To .UsedRange.Rows.count
            flag = binarySearch(psht, 2, psht.UsedRange.Rows.count, CLng(.Cells(i, 3).Value))
            
            If flag <> 0 Then
                    .Cells(i, 16).Value = psht.Cells(flag, 4).Value
                    .Cells(i, 17).Value = psht.Cells(flag, 3).Value
                    .Cells(i, 18).Value = psht.Cells(flag, 1).Value
                    .Cells(i, 19).Value = psht.Cells(flag, 5).Value
                End If
           
            If flag = 0 Then
                .Cells(i, 16).Value = "#N/A"
                .Cells(i, 17).Value = "#N/A"
                .Cells(i, 18).Value = "#N/A"
                .Cells(i, 19).Value = "#N/A"
            End If
        Next
        
        For i = 2 To .UsedRange.Rows.count
            If IsError(.Cells(i, 14).Value) Or IsError(.Cells(i, 18).Value) Then
                .Cells(i, 20).Value = "FALSE"
            ElseIf .Cells(i, 17).Value = "CPP" Then
                .Cells(i, 20).Value = "CPP"
            ElseIf .Cells(i, 14).Value = .Cells(i, 18).Value Then
                .Cells(i, 20).Value = "TRUE"
            Else
                .Cells(i, 20).Value = "FALSE"
            End If
            
            If IsError(.Cells(i, 13).Value) Or IsError(.Cells(i, 17).Value) Then
                .Cells(i, 21).Value = "#N/A"
            ElseIf .Cells(i, 13).Value = .Cells(i, 17).Value Then
                .Cells(i, 21).Value = "TRUE"
            Else
                .Cells(i, 21).Value = "FALSE"
            End If
        Next
        
    End With
    
    sht.Activate
    CreatePT
End Sub

Private Function binarySearch(sht As Worksheet, col As Long, lastrow As Long, number As Long) As Long
Dim left As Long
Dim right As Long
Dim mid As Long
left = 2
right = lastrow
binarySearch = 0
If number = CLng(sht.Cells(left, col).Value) Then
    binarySearch = left
    Exit Function
End If
If number = CLng(sht.Cells(right, col).Value) Then
    binarySearch = right
    Exit Function
End If
While right - left > 1
    mid = Int((right + left) / 2)
    If number = CLng(sht.Cells(mid, col).Value) Then
        binarySearch = mid
        Exit Function
    End If
    If number > CLng(sht.Cells(mid, col).Value) Then
        left = mid
    Else
        right = mid
    End If
    
Wend
End Function


Sub CreatePT()
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As String

NumOfRows = ActiveSheet.UsedRange.Rows.count

RangeStr = "A1:U" & NumOfRows

SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

Set sht = Sheets.Add
sht.Name = "Step 3"
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

 'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
   
  pvt.AddFields RowFields:=Array("BM_RCAO_C", "BM_RCAO_P", "Segment")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.PivotFields("BBM transfer internal (BM_C =BM_P)").Orientation = xlPageField
  pvt.PivotFields("BBM transfer internal (BM_C =BM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM name_RCAO_C").Orientation = xlPageField
  With pvt.PivotFields("RM name_RCAO_C")
  .ClearAllFilters
  '.PivotFilters.Add Type:=xlCaptionContains, Value1:="CPP" ', Value2:="ROE", Value3:="Remote"
  For Each iitm In .PivotItems
    If iitm.Name Like "*CPP*" Or iitm.Name Like "*ROE*" Or iitm.Name Like "Remote" Then
        iitm.Visible = False
    Else
        iitm.Visible = True
    End If
  Next
  End With
  pvt.PivotFields("RM name_RCAO_P").Orientation = xlPageField
  With pvt.PivotFields("RM name_RCAO_P")
  .ClearAllFilters
   For Each iitm In .PivotItems
    If iitm.Name Like "*CPP*" Or iitm.Name Like "*ROE*" Or iitm.Name Like "Remote" Then
        iitm.Visible = False
    Else
        iitm.Visible = True
    End If
  Next
  End With
  For Each pivfld In pvt.PivotFields
      pivfld.Subtotals(1) = False
  Next
  
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = False
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = True
    pvt.EnableDrilldown = True
    
    sht.Columns(4).NumberFormat = "#,###"
    
    sht.Cells(5, 5).Value = "Check if Recalibration needed."
    sht.Cells(6, 5).Value = "Recalibrate"
    For i = 7 To sht.UsedRange.Rows.count
        If sht.Cells(i, 4).Value > 5000000 Then
            sht.Cells(i, 5).Value = "Recalibrate"
        Else
            sht.Cells(i, 5).Value = 0
        End If
    Next
    
    recalibrateData
End Sub

Sub recalibrateData()
Set sht = ActiveSheet
Dim recsht As Worksheet
Dim nameStr As String
Dim count As Long
Dim totAUM As Long
Dim RMname As String
Dim recidx As Integer
Dim BM_P As String
Dim BM_C As String
Dim seg As String
Dim tmp As Integer
Set recsht = Worksheets.Add
recsht.Name = "step 4"
recidx = 0
sht.Activate
For i = 5 To sht.UsedRange.Rows.count
    If sht.Cells(i, 5).Value = "Recalibrate" And sht.Cells(i, 3).Value = "CTG" Then
        sht.Activate
        sht.Cells(i, 4).Select
        seg = sht.Cells(i, 3).Value
        tmp = i
        While sht.Cells(tmp, 2).Value = ""
            tmp = tmp - 1
        Wend
        BM_P = sht.Cells(tmp, 2).Value
        tmp = i
        While sht.Cells(tmp, 1).Value = ""
            tmp = tmp - 1
        Wend
        BM_C = sht.Cells(tmp, 1).Value
        sht.Cells(i, 4).Select
        Selection.ShowDetail = True
        totAUM = 0
        nameStr = ""
        count = ActiveSheet.UsedRange.Rows.count - 1
        For j = 2 To count + 1
            totAUM = totAUM + CLng(ActiveSheet.Cells(j, 9).Value)
            RMname = ActiveSheet.Cells(j, 18).Value
            RMname = Replace(RMname, "(SEC)", "")
            RMname = Trim(RMname)
            If InStr(nameStr, RMname) = 0 Then
                If nameStr <> "" Then nameStr = nameStr + ", "
                nameStr = nameStr & RMname
            End If
        Next
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
        
        ' Output things to setp 4 sheet.
        
        If BM_C <> "CPP" And BM_P <> "CPP" Then
            recidx = recidx + 1
            recsht.Cells(recidx * 2 + 5, 2).Value = recidx
            recsht.Cells(recidx * 2 + 5, 3).Value = seg
            recsht.Cells(recidx * 2 + 5, 4).Value = BM_C
            recsht.Cells(recidx * 2 + 5, 5).Value = BM_P
            recsht.Cells(recidx * 2 + 5, 6).Value = totAUM
            recsht.Cells(recidx * 2 + 5, 7).Value = count
            recsht.Cells(recidx * 2 + 5, 9).Value = "Cms from " & nameStr
        End If
    End If
    
Next


For i = 5 To sht.UsedRange.Rows.count
    If sht.Cells(i, 5).Value = "Recalibrate" And sht.Cells(i, 3).Value = "CTB" Then
        sht.Activate
        sht.Cells(i, 4).Select
        seg = sht.Cells(i, 3).Value
        tmp = i
        While sht.Cells(tmp, 2).Value = ""
            tmp = tmp - 1
        Wend
        BM_P = sht.Cells(tmp, 2).Value
        tmp = i
        While sht.Cells(tmp, 1).Value = ""
            tmp = tmp - 1
        Wend
        BM_C = sht.Cells(tmp, 1).Value
        sht.Cells(i, 4).Select
        Selection.ShowDetail = True
        totAUM = 0
        nameStr = ""
        count = ActiveSheet.UsedRange.Rows.count - 1
        For j = 2 To count + 1
            totAUM = totAUM + CLng(ActiveSheet.Cells(j, 9).Value)
            RMname = ActiveSheet.Cells(j, 18).Value
            RMname = Replace(RMname, "(SEC)", "")
            RMname = Trim(RMname)
            If InStr(nameStr, RMname) = 0 Then
                If nameStr <> "" Then nameStr = nameStr + ", "
                nameStr = nameStr & RMname
            End If
        Next
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
        
        ' Output things to setp 4 sheet.
        If BM_C <> "CPP" And BM_P <> "CPP" Then
            recidx = recidx + 1
            recsht.Cells(recidx * 2 + 5, 2).Value = recidx
            recsht.Cells(recidx * 2 + 5, 3).Value = seg
            recsht.Cells(recidx * 2 + 5, 4).Value = BM_C
            recsht.Cells(recidx * 2 + 5, 5).Value = BM_P
            recsht.Cells(recidx * 2 + 5, 6).Value = totAUM
            recsht.Cells(recidx * 2 + 5, 7).Value = count
            recsht.Cells(recidx * 2 + 5, 9).Value = "Cms from " & nameStr
        End If
    End If
    
Next

recsht.Activate

formatOutput recidx
End Sub


Sub formatOutput(count As Integer)
    Dim sht As Worksheet
    Dim monthstr As String
    monthstr = getMonth(ActiveWorkbook.FullName)
    monthstr = monthstr & right(Format(Date, "YYYY"), 2)
    Set sht = ActiveSheet
    With sht
        .Range("A1:Z1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("C3:H3").Borders(xlEdgeBottom).LineStyle = xlcontinous
        .Range("C3:H3").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C3:H3").Font.Bold = True
        .Range("C3:H3").Interior.ColorIndex = 23
        .Range("C3:H3").Font.Color = vbWhite

        .Range("C4").Font.Color = vbRed
        .Range("C5").Font.Bold = True
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeBottom).LineStyle = xlcontinous
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeTop).LineStyle = xlcontinous
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeTop).Weight = xlMedium
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeRight).LineStyle = xlcontinous
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeRight).Weight = xlMedium
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeLeft).LineStyle = xlcontinous
        .Range("C3:H" & CStr(2 * count + 7)).Borders(xlEdgeLeft).Weight = xlMedium
        
        .Cells(1, 1).Value = "TO BBD"
        .Cells(2, 3).Value = "AUM move movement between RC effective T+1"
        .Cells(3, 3).Value = "Segment"
        .Cells(3, 4).Value = "AUM To"
        .Cells(3, 5).Value = "AUM From"
        .Cells(3, 6).Value = "Total AUM"
        .Cells(3, 7).Value = "No of" & vbNewLine & "Customers"
        .Cells(3, 8).Value = "Month of" & vbNewLine & "move"
        .Cells(3, 9).Value = "Remarks"
        .Cells(4, 3).Value = "Recalibrate " & monthstr & " SoV"
        .Cells(5, 3).Value = "All Zones"
        .Cells(2 * count + 8, 3).Value = "* Note (Rule of Thumb) : For BBM recalibration it must be more than or less than 5MM for calibration."
      
        .Columns("D:H").AutoFit
        .Columns("B:B").ColumnWidth = 3
        .Columns("C:C").ColumnWidth = 10
        .Rows("3:3").RowHeight = 30
        .Columns("H:H").NumberFormat = "@"
        .Columns("F:F").NumberFormat = "#,###"
        
        For i = 1 To count
            'MsgBox .Cells(2 * i + 5, 3).Value
            If .Cells(2 * i + 5, 3).Value = "CTG" Then
                .Range("B" & CStr(2 * i + 5) & ":H" & CStr(2 * i + 5)).Interior.Color = 14281213
            End If
            If .Cells(2 * i + 5, 3).Value = "CTB" Then
                .Range("B" & CStr(2 * i + 5) & ":H" & CStr(2 * i + 5)).Interior.ColorIndex = 20
            End If
            .Cells(2 * i + 5, 8).Value = monthstr
        Next
    End With
End Sub



Private Sub SOPButton_Click()

End Sub



Private Sub UpdateNameButton_Click()
    Dim strpath As String
    Application.FileDialog(msoFileDialogOpen).Title = "Select the most recent SOP masterlist file"

    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        'get the file path selected by the user
        strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Else: Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Workbooks.Open strpath, UpdateLinks = xlUpdateLinksNever
    
    Set wkbk = ActiveWorkbook
    ' check whether there is a sheet named "SOP"
    Dim flag As Integer
    flag = 0
    For Each st In wkbk.Worksheets
        If st.Name = "SOP" Then
            flag = 1
            Exit For
        End If
    Next
    If flag = 0 Then
        MsgBox "No SOP Tab Detected. Please select the correct SOP Masterlist"
        ActiveWorkbook.Close
        Exit Sub
    End If
    
    Set sht = wkbk.Worksheets("SOP")
    sht.Activate
    sht.Range("C:C,G:G").Select
    Selection.Copy
    
    Dim nlPath As String
    nlPath = Application.ThisWorkbook.path + "\NameList.xls"
     
    Set nlwkbk = Workbooks.Add
    
    Set nlsht = nlwkbk.Worksheets(1)
    nlsht.Name = "NameList"

    nlsht.Activate
    nlsht.Range("A:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Dim str As String
    nlsht.Range("A:B").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    For i = 1 To nlsht.UsedRange.Rows.count
        str = CStr(nlsht.Cells(i, 1).Value)
        If str Like "*(SEC)*" Or str Like "*untag*" Or str Like "*not in use*" Then
            nlsht.Rows(i).Delete
            i = i - 1
        End If
    Next
    
    Application.DisplayAlerts = False
    nlwkbk.SaveAs Filename:=nlPath, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    nlwkbk.Close
    wkbk.Close
    Application.ScreenUpdating = True
    MsgBox "Update Name List is Done."
End Sub

Sub CPCpage()
Dim cpcsht As Worksheet
Set cpcsht = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = ActiveWorkbook.Worksheets("Movement 1&2")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As String


On Error Resume Next

NumOfRows = datasht.UsedRange.Rows.count

RangeStr = "A1:U" & NumOfRows

SrcData = datasht.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

cpcsht.Name = "CPC"
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "G3"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
  
'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("G1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Upgrade").Orientation = xlPageField
  pvt.PivotFields("Upgrade").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "G3"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 

'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("M1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable3")
 
    pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
  pvt.AddFields RowFields:=Array("BBD_RCAO_P", "BM_RCAO_P", "RM name_RCAO_P")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_P").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_P").CurrentPage = "G3"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 cpcsht.Range("D:D,J:J,P:P").NumberFormat = "#,###"
 
 'append bottom subtable without sec port
 

End Sub

Sub makeSecondTable(nm As String)
Set cpcsht = ActiveWorkbook.Worksheets(nm)
For k = 1 To 3
    Sum = 0
    Set pvt = cpcsht.PivotTables(k)
    tendrow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.count).Row - 1
    tstartrow = 7
    tstartcol = pvt.TableRange1.Columns(pvt.TableRange1.Columns.count).Column - 2
    
    idx = tendrow + 4
    For i = tstartrow To tendrow
        With cpcsht
            If Not .Cells(i, tstartcol).Value Like "*(SEC)" Then
                .Cells(idx, tstartcol) = .Cells(i, tstartcol)
                .Cells(idx, tstartcol + 1) = .Cells(i, tstartcol + 1)
                Sum = Sum + .Cells(idx, tstartcol + 1)
                idx = idx + 1
            End If
        End With
    Next
    
    cpcsht.Cells(idx, tstartcol) = "Total"
    cpcsht.Cells(idx, tstartcol).Font.Bold = True
    cpcsht.Cells(idx, tstartcol + 1) = Sum
  
 Next
End Sub


Sub RMpage()
Dim cpcsht As Worksheet
Set cpcsht = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = ActiveWorkbook.Worksheets("Movement 1&2")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As String


On Error Resume Next

NumOfRows = datasht.UsedRange.Rows.count

RangeStr = "A1:U" & NumOfRows

SrcData = datasht.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

cpcsht.Name = "RM"
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "RM"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
  
'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("G1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Upgrade").Orientation = xlPageField
  pvt.PivotFields("Upgrade").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "RM"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 

'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("M1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable3")
 
    pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
  pvt.AddFields RowFields:=Array("BBD_RCAO_P", "BM_RCAO_P", "RM name_RCAO_P")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_P").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_P").CurrentPage = "RM"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 cpcsht.Range("D:D,J:J,P:P").NumberFormat = "#,###"

End Sub

Sub PBpage()
Dim cpcsht As Worksheet
Set cpcsht = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = ActiveWorkbook.Worksheets("Movement 1&2")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As String


On Error Resume Next

NumOfRows = datasht.UsedRange.Rows.count

RangeStr = "A1:U" & NumOfRows

SrcData = datasht.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

cpcsht.Name = "PB"
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "PB"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
  
'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("G1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")
 
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  pvt.AddFields RowFields:=Array("BBD_RCAO_C", "BM_RCAO_C", "RM name_RCAO_C")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_C").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Upgrade").Orientation = xlPageField
  pvt.PivotFields("Upgrade").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_C").CurrentPage = "PB"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 

'--------
StartPvt = "'" & cpcsht.Name & "'" & "!" & cpcsht.Range("M1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable3")
 
    pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
  pvt.AddFields RowFields:=Array("BBD_RCAO_P", "BM_RCAO_P", "RM name_RCAO_P")
  pvt.AddDataField pvt.PivotFields("AUM"), "Sum of AUM", xlSum
  pvt.AddDataField pvt.PivotFields("custnum"), "Count of custnum", xlCount
  pvt.PivotFields("RM Rank_P").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").Orientation = xlPageField
  
  pvt.PivotFields("Transfer_Internal bet RM").Orientation = xlPageField
  pvt.PivotFields("Transfer_Internal bet RM").CurrentPage = 1
  pvt.PivotFields("Transfer_Internal (RM_C = RM_P)").CurrentPage = "FALSE"
  pvt.PivotFields("RM Rank_P").CurrentPage = "PB"
  For Each pivfld In pvt.PivotFields
    If pivfld.Name <> "Values" Then
      pivfld.Subtotals(1) = False
    End If
  Next
  
 cpcsht.Range("D:D,J:J,P:P").NumberFormat = "#,###"

End Sub


Private Sub SelectPathButton_Click()
    saveToFolder = GetFolder2
End Sub

Function GetFolder2() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = Application.ActiveWorkbook.path
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder2 = sItem
Set fldr = Nothing
End Function
