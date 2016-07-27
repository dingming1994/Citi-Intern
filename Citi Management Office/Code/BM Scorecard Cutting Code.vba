Public saveToPath As String

Sub Cut_CC(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("Computation_Case")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim stdidx As Integer
Dim tmp As Integer
Dim idxset As New Collection
cutsht.Cells(6, 3).ClearContents
For Each nm In namelist
    If nm <> "" Then
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        stdidx = 0
        If cutsht.Cells(j, 3) = nm Or cutsht.Cells(j, 7) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            tmp = idx
            
            While IsError(cutsht.Cells(tmp, 9).Value) Or Not (cutsht.Cells(tmp, 8).Value = "" And cutsht.Cells(tmp, 3).Value = "" And cutsht.Cells(tmp, 9).Value <> "")
                tmp = tmp - 1
            Wend
            stdidx = tmp
            idxset.Add idx
            idxset.Add stdidx
        End If
        
    Next
    End If
Next

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 6 To cutsht.UsedRange.Rows.Count
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If
        
Next
End Sub


Sub Cut_BREV(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("BREV")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If

Next
End Sub

Sub Cut_NCG(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("NCG")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If

Next
End Sub



Sub Cut_AUM(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("AUM")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1
idxset.Add 2

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If
Next

If (Not namelist(0) Like "Tony Wong") And (Not namelist(0) Like "Valerie") Then
Application.DisplayAlerts = False
cutsht.Delete
Application.DisplayAlerts = True
End If
End Sub


Sub Cut_BWP(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("BWP")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If

Next
End Sub

Sub Cut_NPS(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("NPS")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1
idxset.Add 2

'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If

Next
End Sub


Sub Cut_ABU(namelist As Variant)
Dim cutbk As Workbook
Dim cutsht As Worksheet
Set cutsht = ActiveWorkbook.Worksheets("ABU")
cutsht.Activate
cutsht.Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Dim nm As Variant
Dim idx As Integer
Dim idxset As New Collection

For Each nm In namelist
    For j = 1 To cutsht.UsedRange.Rows.Count
        idx = 0
        If cutsht.Cells(j, 6) = nm Or cutsht.Cells(j, 10) = nm Then
            idx = j
        End If
        
        If idx <> 0 Then
            idxset.Add idx
        End If
    Next
Next
idxset.Add 1
idxset.Add 2
'-----
'clear rows which are not in idxset
Dim flag As Integer
For i = 1 To cutsht.UsedRange.Rows.Count
    If cutsht.Cells(i, 5) = "Note:" Then Exit For
    flag = 0
    For j = 1 To idxset.Count
        If i = idxset(j) Then
            flag = 1
        End If
    Next
    If flag = 0 Then
        cutsht.Rows(i).ClearContents
    End If

Next
End Sub

Sub main()
Dim nmsht As Worksheet
Set nmsht = ThisWorkbook.Worksheets("name")
Dim oripath As String

If saveToPath = "" Then
MsgBox "Please select the saving directory first."
Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the BM scorecard file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\PLUS\"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    oripath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else: Exit Sub
End If


Application.FileDialog(msoFileDialogOpen).Title = "Select the BSC file"
Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\PLUS\"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    bscpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
End If

If Not oripath Like "*Q* BM scorecard*" Then
MsgBox "Please select the correct BM scorecard file"
Exit Sub
End If
'oripath = "I:\CAP_Profile_PRD65\Desktop\Scorecard Cutting\Q215 BM scorecard target v2 (for cutting).xlsx"
qtrstr = Right(oripath, Len(oripath) - InStrRev(oripath, "\"))
qtrstr = Left(qtrstr, 4)

If bscpath <> "" Then
    Set bscbk = Workbooks.Open(bscpath)
    Set bsc_issht = bscbk.Worksheets("Individual_Statement")
     bsc_issht.Range("A1:G23").UnMerge
    Set bksht = bscbk.Worksheets("Staff_Payput_Details")
    Set mgtsht = bscbk.Worksheets("MGT_BSC_SUMMARY")
End If
Application.ScreenUpdating = False
Dim savePath As String
Dim nm As String
Dim nmlist As Variant
Dim pos As String
For i = 2 To nmsht.UsedRange.Rows.Count
    nm = nmsht.Cells(i, 1).Value
    pos = nmsht.Cells(i, 2).Value
    geid = nmsht.Cells(i, 7).Value
    If pos = "SBM" Then
        nmlist = Split(nmsht.Cells(i, 3).Value & ", " & nm, ", ")
    Else
        nmlist = Array(nm)
    End If
    
    If nm <> "" And pos <> "" Then
        Workbooks.Open oripath, Password:="Summer" & Right(qtrstr, 2)
        savePath = saveToPath & "\" & qtrstr & " BM scorecard target_" & nm & ".xlsx"
        ActiveWorkbook.SaveAs savePath
        Set tarbk = ActiveWorkbook
        Cut_CC nmlist
        Cut_BREV nmlist
        Cut_NCG nmlist
        Cut_AUM nmlist
        Cut_BWP nmlist
        Cut_NPS nmlist
        Cut_ABU nmlist
        
        'start BSC
        If bscpath <> "" Then
            Set issht = tarbk.Worksheets.Add(after:=tarbk.Worksheets(tarbk.Worksheets.Count))
            bsc_issht.Cells(4, 3) = "'" & geid
            bsc_issht.Range("A1:G23").Copy
            tarbk.Activate
            issht.Range("A1").PasteSpecial xlPasteAll
            issht.Range("A1").PasteSpecial xlPasteValues
            issht.Columns("B:B").ColumnWidth = 30
            issht.Columns("C:C").ColumnWidth = 8
            issht.Columns("D:F").AutoFit
            issht.Rows(12).AutoFit
            issht.name = "Individual_Statement (BSC)"
            
            ' individual banker statement
            Set psht = tarbk.Worksheets.Add(after:=tarbk.Worksheets(tarbk.Worksheets.Count))
            
                fld = 0
                If pos = "SBM" Then fld = 10
                If pos = "BM" Then fld = 6
                If pos = "BBD" Then fld = 8
                
                If fld <> 0 Then
                With bksht
                .AutoFilterMode = False
                .Range("A1:AA1").AutoFilter
                .Range("A1:AA1").AutoFilter Field:=fld, Criteria1:=nm
                .Range("A1").CurrentRegion.Copy
                End With
                psht.Activate
                psht.Range("A1").PasteSpecial xlPasteValues
                psht.name = "Staff_Payput_Details (BSC)"
                End If
            
            'pivot table
            createBSCPT
            psht.Visible = False
            
            
            ' update first page summary data
            Set ccsht = tarbk.Worksheets("Computation_Case")
            ccsht.Range("V4").Value = "BSC  Deduction"
            ccsht.Range("W4").Value = "BSC Grade"
            ccsht.Range("O4:O100").Copy
            ccsht.Range("V4:W100").PasteSpecial xlPasteFormats
            ccsht.Range("V4:W100").FormatConditions.Delete
            ccsht.Range("V4:V100").NumberFormat = "0.00%"
            
            For k = 6 To ccsht.UsedRange.Rows.Count
                If ccsht.Cells(k, 7) <> "" Then
                    nm = ccsht.Cells(k, 7)
                    For j = 2 To mgtsht.Range("C1").End(xlDown).Row
                        If Trim(mgtsht.Cells(j, 3)) = Trim(nm) Then
                            ccsht.Cells(k, 22) = mgtsht.Cells(j, 7)
                            ccsht.Cells(k, 23) = mgtsht.Cells(j, 9)
                        End If
                    Next
                End If
            Next
            
            ccsht.Columns("P").ShowDetail = False
        End If
        tarbk.Save
        tarbk.Close
    End If
    
    
Next
Application.ScreenUpdating = True
MsgBox "BM Scorecard Cutting is Done." & vnnewline & "Please check the file in selected folder."
End Sub


Sub createBSCPT()
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim numOfRow As Long

On Error Resume Next
Set bk = ActiveWorkbook
Set shet = bk.Worksheets("Staff_Payput_Details (BSC)")
Set sht = bk.Worksheets("Individual_Statement (BSC)")
shet.Activate
numOfRow = shet.Range("A2").End(xlDown).Row
ridx = sht.UsedRange.Rows.Count + 5
 
 'SrcData = shet.Name & "!" & Range("A2:N" & numOfRow).Address(ReferenceStyle:=xlR1C1)
 SrcData = "'" & shet.name & "'!" & Range("A1:AA" & numOfRow).Address(ReferenceStyle:=xlR1C1)
 StartPvt = "'" & sht.name & "'!" & sht.Range("B" & CStr(ridx)).Address(ReferenceStyle:=xlR1C1)
  Set pvtCache = bk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    pvt.AddFields RowFields:=Array("Staff_Name", "Grade")
    
    pvt.AddDataField pvt.PivotFields("Pre_BSC_Payout"), "Average of Pre_BSC_Payout", xlAverage
    pvt.AddDataField pvt.PivotFields("Post_BSC_Payout"), "Average of Post_BSC_Payout", xlAverage
    pvt.InGridDropZones = True
    pvt.RowAxisLayout xlTabularRow
    pvt.ColumnGrand = True
    pvt.SubtotalLocation xlAtBottom
    pvt.PivotFields("Staff_Name").Subtotals(1) = False
    pvt.PivotFields("Grade").Subtotals(1) = False
        
    sht.Cells(ridx + 1, 6).Value = "BSC Deduction"
    endrow = sht.Range("B" & CStr(ridx + 2)).End(xlDown).Row
    If endrow < 10000 Then
     For i = ridx + 2 To endrow
         sht.Cells(i, 6).FormulaR1C1 = "=1-RC[-1]/RC[-2]"
         If IsError(sht.Cells(i, 6)) Then
             sht.Cells(i, 6) = 0
         End If
     Next
    ' sht.Range("E" & CStr(ridx ) & ":E" & CStr(endrow)).Select
     sht.Range("E" & CStr(ridx) & ":E" & CStr(endrow)).Copy
     sht.Range("F" & CStr(ridx) & ":F" & CStr(endrow)).PasteSpecial xlPasteFormats
     sht.Range("D" & CStr(ridx) & ":E" & CStr(endrow)).NumberFormat = "#,0"
     sht.Range("F" & CStr(ridx) & ":F" & CStr(endrow)).NumberFormat = "0.00%"
     
     sht.Cells(ridx - 2, 2) = "Banker Individual Statement"
     sht.Range("B1").Copy
     sht.Cells(ridx - 2, 2).PasteSpecial xlPasteFormats
    End If
End Sub

Sub sendMail()

Application.ScreenUpdating = False

If saveToPath = "" Then
MsgBox "Please select the saving directory first."
Exit Sub
End If
Dim dlstr As String
dlstr = InputBox(prompt:="Please indicate the deadline for reversion")
If dlstr = "" Then
MsgBox "Please indicate the revertion deadline"
Exit Sub
End If

Dim oApp As Object
Dim outApp As Outlook.Application
Set oApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFSO.GetFolder(saveToPath)
Set thiswkbk = ActiveWorkbook

Dim toStr As String
Dim ccStr As String
Dim subjectStr As String
Dim bodyStr As String
Dim monStr As String
Dim yearStr As String
Dim nameStr As String
Dim BMname As String
Dim emailname As String
Dim BRcode As String

Dim zipCount As Integer
Dim mailCount As Integer
zipCount = 0
mailCount = 0

Dim qtrstr As String




For Each objfile In objfolder.Files
    
    If objfile.name Like "*scorecard target_*.xlsx" Then
        toStr = ""
        ccStr = ""
        subjectStr = ""
        bodyStr = ""
  
        BMname = Right(objfile.name, Len(objfile.name) - InStrRev(objfile.name, "_"))
        BMname = Left(BMname, InStr(BMname, ".") - 1)
        qtrstr = Left(objfile.name, 4)
        qtrstr = Left(qtrstr, 2) & " 20" & Right(qtrstr, 2)
        nameStr = ""
        zipCount = zipCount + 1
        monStr = Mid(objfile.name, 5, 3)
        yearStr = Mid(objfile.name, 9, 4)
        Set sht = ThisWorkbook.Worksheets("name")
        
        '***************
        For i = 2 To sht.UsedRange.Rows.Count
            If sht.Cells(i, 1).Value = BMname Then
                emailname = sht.Cells(i, 4).Value
                toStr = sht.Cells(i, 5).Value
                ccStr = sht.Cells(i, 6).Value
            End If
        Next
        
        subjectStr = qtrstr & "BM scorecard - Achievement (Deadline " & dlstr & ") (" & emailname & ")"
        
        bodyStr = "Hi " & emailname & "," & vbNewLine & vbNewLine & _
                      "Attached is your individual " & qtrstr & " BM scorecard - Achievement." & vbNewLine & vbNewLine & _
                      "Please revert by " & dlstr & "." & vbNewLine & vbNewLine & _
                      "Password : Su****" & Right(qtrstr, 2) & vbNewLine & vbNewLine & _
                      "Thank you!" & vbNewLine & vbNewLine & _
                      "Best Regards," & vbNewLine & "Anna"
    
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
MsgBox "Email sending is Done."
End Sub


Private Sub CheckListButton_Click()
    ThisWorkbook.Worksheets("name").Activate
    ActiveSheet.Cells(1, 1).Select
End Sub

Private Sub CheckPathButton_Click()
    If saveToPath = "" Then
        MsgBox "You have not select any path."
    Else
        MsgBox saveToPath
    End If
End Sub

Private Sub CuttingButton_Click()
    main
End Sub

Private Sub EmailButton_Click()
    sendMail
End Sub

Private Sub SelectPathButton_Click()
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
