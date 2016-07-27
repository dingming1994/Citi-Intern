Public saveToPath As String
Public monthstr As String



Private Sub CreateUploadButton_Click()
If saveToPath = "" Then
    MsgBox "You haven't select a saving path!"
    Exit Sub
End If
Application.FileDialog(msoFileDialogOpen).Title = "Select the Approved Deals file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If
Application.ScreenUpdating = False

monthstr = Left(Right(strpath, Len(strpath) - InStrRev(strpath, "\")), 5)
Workbooks.Open strpath
CreateUploadFile
ThisWorkbook.Worksheets("main").Activate
Application.ScreenUpdating = True
MsgBox "Upload file has been generated successfully."
End Sub

Private Sub GetARDealButton_Click()

If saveToPath = "" Then
    MsgBox "You haven't select a saving path!"
    Exit Sub
End If
Application.FileDialog(msoFileDialogOpen).Title = "Select the AdjustmentT file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

monthstr = InputBox("Please Input the month/year in MMMYY (Jun15)", Title:="Input Month")
If Not (monthstr Like "???##") Then
    MsgBox "Please input valid time!"
    Exit Sub
End If
Application.ScreenUpdating = False
Set wkbk = Workbooks.Open(strpath)
getARDeal
Application.ScreenUpdating = True
MsgBox "Approved & Rejected Deals have been generated successfully"
End Sub

Private Sub GetPDAdjustButton1_Click() 'NOT DONE!!!
If saveToPath = "" Then
    MsgBox "You haven't select a saving path!"
    Exit Sub
End If
Application.FileDialog(msoFileDialogOpen).Title = "Select the Approved Deals file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the Contracted Credit (v1.0) file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\RM\SIPs Adjustment Request Portal\Admin\Working files\"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If


Application.ScreenUpdating = False

monthstr = Left(Right(strpath, Len(strpath) - InStrRev(strpath, "\")), 5)
Workbooks.Open strpath



ThisWorkbook.Worksheets("main").Activate
Application.ScreenUpdating = True
MsgBox "Upload file has been generated successfully."
End Sub

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
    .InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"
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


Private Sub UploadCheckButton_Click()
Dim strpath, strpath1, strpath2 As String


Application.FileDialog(msoFileDialogOpen).Title = "Select the generated Adjustment file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If


Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP masterlist file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\(SK) SOP clean up files\"

intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath1 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the Contracted Credit (v1.0) file"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\Sabrina\2015 SIPs Adjustment\"

intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath2 = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If
Application.ScreenUpdating = False
Set bk = Workbooks.Open(strpath)
Dim sht As Worksheet

SOP_In (strpath1)
bk.Activate
Dim flag As Integer
flag = 0
For Each st In ThisWorkbook.Worksheets
If st.Name = "Checking Log" Then
    Set sht = st
    sht.Cells.ClearContents
    flag = 1
End If
Next

If flag = 0 Then
Set sht = ThisWorkbook.Worksheets.Add
sht.Name = "Checking Log"
End If

bk.Worksheets("Sheet1").Activate
bk.Worksheets("Sheet1").Rows("3:" & CStr(bk.Worksheets("Sheet1").UsedRange.Rows.count)).Interior.ColorIndex = 0
UploadFileChecking sht
UploadFileChecking_APR sht, strpath2
UploadFileChecking_RCAO sht
Application.DisplayAlerts = False
bk.Worksheets("SOP").Delete
bk.Worksheets("All_Products_Raw").Delete
bk.Save
bk.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Upload file checking is done." & vbNewLine & "Please review the file."
End Sub

Private Sub UploadFileChecking_RCAO(logsht As Worksheet)
Dim bk As Workbook
Set bk = ActiveWorkbook
Dim sht, rsht As Worksheet
Set sht = bk.Worksheets("Sheet1")
Set rsht = bk.Worksheets("SOP")
Dim rno As Integer
'for each creit RCAO, find in SOP
Dim rcao As String
For i = 1 To sht.UsedRange.Rows.count
    rcao = CStr(sht.Cells(i, 14))
    If Len(rcao) = 6 Then
        Set c = rsht.Range("B:B").Find(rcao)
        If Not c Is Nothing Then
            If CStr(sht.Cells(i, 15)) <> CStr(rsht.Cells(c.Row, 1)) Then
                sht.Cells(i, 14).Interior.Color = vbRed
                sht.Cells(i, 15).Interior.Color = vbRed
                logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "Credit RCAO-Banker mapping error in row " & CStr(i)
            End If
            
            If CStr(sht.Cells(i, 16)) <> CStr(rsht.Cells(c.Row, 3)) Then
                sht.Cells(i, 14).Interior.Color = vbRed
                sht.Cells(i, 16).Interior.Color = vbRed
                logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "Credit RCAO-BM mapping error in row " & CStr(i)
            End If
         
        Else
            sht.Cells(i, 14).Interior.Color = vbRed
            logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "Credit RCAO not found error in row " & CStr(i)
        End If
    End If

Next
End Sub

Private Sub UploadFileChecking_APR(logsht As Worksheet, strpath2 As String)
Dim bk As Workbook
Set bk = ActiveWorkbook
Dim sht, asht As Worksheet
Set sht = bk.Worksheets("Sheet1")
Dim cbk As Workbook
Set cbk = Workbooks.Open(strpath2)
cbk.Worksheets("All_Products_Raw").Copy before:=bk.Worksheets(1)
Application.DisplayAlerts = False
cbk.Close
Application.DisplayAlerts = True

Set asht = bk.Worksheets("All_Products_Raw")
Dim rno As Integer

' For each ref no, find in All_PR
Dim refno As String
For i = 3 To sht.UsedRange.Rows.count
    refno = sht.Cells(i, 9)
    Set c = asht.Range("E:E").Find(refno)
    If Not c Is Nothing Then
        'c.Select
        rno = c.Row
       
        If Not CDbl(sht.Cells(i, 10)) = CDbl(asht.Cells(rno, 2)) Then
            sht.Cells(i, 10).Interior.Color = vbBlue
            logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "FAdjVol is not consistent to ALL_PRODUCT_RAW in row " & CStr(i)
        End If
        If Not Round(CDbl(sht.Cells(i, 11)), 1) = Round(CDbl(asht.Cells(rno, 3)), 1) Then
            sht.Cells(i, 11).Interior.Color = vbBlue
            logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "FAdjRev is not consistent to ALL_PRODUCT_RAW in row " & CStr(i)
        End If
        If Not CDbl(sht.Cells(i, 13)) = CDbl(asht.Cells(rno, 16)) Then
            sht.Cells(i, 13).Interior.Color = vbBlue
            logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "FAdjCredits is not consistent to ALL_PRODUCT_RAW in row " & CStr(i)
        End If
        If Not CStr(sht.Cells(i, 2)) = CStr(asht.Cells(rno, 4)) Then
            sht.Cells(i, 2).Interior.Color = vbBlue
            logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "Debit RCAO is not consistent to ALL_PRODUCT_RAW in row " & CStr(i)
        End If
    Else
        sht.Cells(i, 9).Interior.Color = vbBlue
        logsht.Cells(logsht.UsedRange.Rows.count + 1, 1) = "Ref No not found in ALL_PRODUCT_RAW in row " & CStr(i)
    End If
Next
bk.Activate
End Sub

Private Sub UploadFileChecking(logsht As Worksheet)
Dim sht As Worksheet
Set sht = ActiveSheet
Dim numOfRow As Long
numOfRow = sht.UsedRange.Rows.count

If numOfRow = 2 Then Exit Sub
Dim count As Integer
count = 0

'Col 1: ID
' Non-Empty
For i = 3 To numOfRow
    If sht.Cells(i, 1) = "" Then
        sht.Cells(i, 1).Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Empty ID on cell(" & CStr(i) & ", 1)"
    End If
Next

'Col 2: Debit RCAO
' Six digit number or empty(in yellow color)
For i = 3 To numOfRow
    With sht.Cells(i, 2)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty Debit RCAO on cell(" & CStr(i) & ", 2)"
    ElseIf Len(CStr(.Value)) <> 6 Or IsNumeric(.Value) = False Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Wrong RCAO on cell(" & CStr(i) & ", 2)"
    End If
    End With
Next

'Col 3: Debit RM
'SKIP

'Col 4: Debit BBM
' yellow if empty
For i = 3 To numOfRow
    With sht.Cells(i, 4)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty Debit BBM on cell(" & CStr(i) & ", 4)"
    End If
    End With
Next

'Col 6: Product ID
' red if empty
For i = 3 To numOfRow
    With sht.Cells(i, 6)
    If .Value = "" Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Empty product ID on cell(" & CStr(i) & ", 6)"
    End If
    End With
Next

'Col 7: CusNum
' yellow if empty
For i = 3 To numOfRow
    With sht.Cells(i, 7)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty CusNum on cell(" & CStr(i) & ", 7)"
    ElseIf Len(CStr(.Value)) <> 9 Or IsNumeric(.Value) = False Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Wrong CusNum on cell(" & CStr(i) & ", 7)"
    End If
    
    End With
Next

'Col 8:RelNum
' yellow if empty
For i = 3 To numOfRow
    With sht.Cells(i, 8)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty RelNum on cell(" & CStr(i) & ", 8)"
    ElseIf Len(CStr(.Value)) <> 9 Or IsNumeric(.Value) = False Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Wrong RelNum on cell(" & CStr(i) & ", 8)"
    End If
  
    End With
Next

'Col 9: ref no
' red if empty
For i = 3 To numOfRow
    With sht.Cells(i, 9)
    If .Value = "" Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Empty RefNo on cell(" & CStr(i) & ", 9)"
    End If
    End With
Next

'Col 12: NTB/OTB
' red if empty
For i = 3 To numOfRow
    With sht.Cells(i, 12)
    If .Value = "" Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Empty NTB/OTB on cell(" & CStr(i) & ", 12)"
    End If
    End With
Next

'Col 14: Credit RCAO
' Six digit number or empty(in yellow color)
For i = 3 To numOfRow
    With sht.Cells(i, 14)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty Credit RCAO on cell(" & CStr(i) & ", 14)"
    ElseIf Len(CStr(.Value)) <> 6 Or IsNumeric(.Value) = False Then
        .Interior.Color = vbRed
        count = count + 1
        logsht.Cells(count, 1) = "Wrong Credit RCAO on cell(" & CStr(i) & ", 14)"
    End If
    End With
Next

'Col 16: Credit BBM
' yellow if empty
For i = 3 To numOfRow
    With sht.Cells(i, 16)
    If .Value = "" Then
        .Interior.Color = vbYellow
        count = count + 1
        logsht.Cells(count, 1) = "Empty Credit BBM on cell(" & CStr(i) & ", 16)"
    End If
    End With
Next
sht.Activate
End Sub


Private Sub getARDeal()
Dim bk As Workbook
Dim Tsht As Worksheet
Dim apsht As Worksheet

Set bk = ActiveWorkbook
Set Tsht = bk.Worksheets("AdjustmentT")

Tsht.Activate
Tsht.UsedRange.AutoFilter 36, "TRUE"

Tsht.Cells.Select
Selection.Copy
Set apsht = bk.Worksheets.Add
apsht.Activate
apsht.Paste
apsht.Name = "Approved Deals"
apsht.Columns.AutoFit
ActiveWindow.Zoom = 80
'---
Tsht.Activate
Tsht.UsedRange.AutoFilter 36, "FALSE"

Tsht.Cells.Select
Selection.Copy
Set apsht = bk.Worksheets.Add
apsht.Activate
apsht.Paste
apsht.Name = "Rejected Deals"

apsht.UsedRange.Sort key1:=apsht.Range("K2"), order1:=xlAscending, Header:=xlYes
apsht.Columns.AutoFit
ActiveWindow.Zoom = 80

Application.DisplayAlerts = False
Tsht.Delete
bk.SaveAs saveToPath & "\" & monthstr & "AdjustmentT (All Deals).xlsx", ConflictResolution:=xlLocalSessionChanges
bk.Worksheets("Approved Deals").Copy
ActiveWorkbook.SaveAs saveToPath & "\" & monthstr & "AdjustmentT (Approved Deals).xlsx", ConflictResolution:=xlLocalSessionChanges
ActiveWorkbook.Close
bk.Close
Application.DisplayAlerts = True


End Sub

Private Sub CreateUploadFile()
Dim bk As Workbook
Dim sht As Worksheet
Set bk = ActiveWorkbook
Set sht = ActiveSheet
Dim apsht As Worksheet
Dim tmpsht As Worksheet
Set tmpsht = ThisWorkbook.Worksheets("Template")
Set apsht = bk.Worksheets.Add
Dim numOfRow As Long
numOfRow = sht.UsedRange.Rows.count

tmpsht.Activate
tmpsht.Rows("1:2").Copy
apsht.Activate
apsht.Rows("1:2").Select
ActiveSheet.Paste

For i = 1 To apsht.UsedRange.Columns.count
    For j = 1 To sht.UsedRange.Columns.count
        If apsht.Cells(2, i).Value = sht.Cells(1, j).Value Then
            sht.Activate
            sht.Range(sht.Cells(2, j), sht.Cells(numOfRow, j)).Select
            Selection.Copy
            apsht.Activate
            apsht.Cells(3, i).Select
            ActiveSheet.Paste
        End If
    Next
Next

apsht.Copy
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs saveToPath & "\" & monthstr & "_Adjustment.xlsx"
ActiveWorkbook.Close
bk.Close
Application.DisplayAlerts = True
End Sub


Private Sub SOP_In(strpath1 As String)

Set tarbk = ActiveWorkbook
Workbooks.Open strpath1, UpdateLinks:=False
Set cpbk = ActiveWorkbook
Set Tsht = cpbk.Worksheets("SOP")

Dim ColIdxArr(5) As Integer
Dim ColTitle As Variant
ColTitle = Array("RCAO", "RM", "Mgr", "BBD", "Type")
For i = 1 To 5
    ColIdxArr(i) = 0
Next

For i = 1 To Tsht.UsedRange.Columns.count
    For j = 0 To 4
        If Tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = Tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 5
    If ColIdxArr(k) <> 0 Then
    Set Y = Union(Y, Tsht.UsedRange.Columns(ColIdxArr(k)))
    End If
Next
Tsht.Activate
Y.Select
Selection.Copy
tarbk.Activate
tarbk.Worksheets.Add
Set tarsht = ActiveSheet
tarsht.Name = "SOP"


tarsht.Cells(1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set Tsht = cpbk.Worksheets("ROE & others")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To Tsht.UsedRange.Columns.count
    For j = 0 To 3
        If Tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = Tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, Tsht.UsedRange.Columns(ColIdxArr(k)))
Next
Tsht.Activate
Y.Select
Selection.Copy
tarsht.Activate

tarsht.Cells(tarsht.UsedRange.Rows.count + 1, 1).PasteSpecial xlPasteValues

'--------------------------------
Set Tsht = cpbk.Worksheets("GEB & Remote codes")

For i = 1 To 4
    ColIdxArr(i) = 0
Next

For i = 1 To Tsht.UsedRange.Columns.count
    For j = 0 To 3
        If Tsht.Cells(1, i).Value = ColTitle(j) And ColIdxArr(j + 1) = 0 Then
            ColIdxArr(j + 1) = i
        End If
    Next
Next
Set Y = Tsht.UsedRange.Columns(ColIdxArr(1))
    
For k = 2 To 4
    Set Y = Union(Y, Tsht.UsedRange.Columns(ColIdxArr(k)))
Next
Tsht.Activate
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

End Sub
