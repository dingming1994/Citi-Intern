Public P_HRBK As Workbook
Public P_PRBK As Workbook
Public P_SOPBK As Workbook
Public P_CUBK As Workbook

Public P_HRPATH, P_PRPATH, P_SOPPATH, P_CUPATH, RBWMPath As String
Public monthstr As String
Public savetopath As String



Private Sub SaveDirButton_Click()
    savetopath = GetFolder
End Sub

Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS\"
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function


Private Sub GetInputFileButton_Click()
    Application.FileDialog(msoFileDialogOpen).Title = "Select the HR Report"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
    Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\2015 Attrition Report\(8) Aug'15 Attrition Report\"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        P_HRPATH = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not P_HRPATH Like "*Active*" Then
            MsgBox "Please Select The Correct File!"
            Exit Sub
        End If
    Else: Exit Sub
    End If
    
    Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP Masterlist"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
    Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\(SK) SOP clean up files\08. Aug15 SOP Masterlist\"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        P_SOPPATH = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not P_SOPPATH Like "*SOP Masterlist*" Then
            MsgBox "Please Select The Correct File!"
            Exit Sub
        End If
    Else: Exit Sub
    End If
    
    Application.FileDialog(msoFileDialogOpen).Title = "Select the Attrition Report for Previous Month"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
    Application.FileDialog(msoFileDialogOpen).InitialFileName = "I:\CAP_Profile_PRD65\Desktop\Attrition Report Automation"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        P_PRPATH = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not P_PRPATH Like "*Attrition Report*" Then
            MsgBox "Please Select The Correct File!"
            Exit Sub
        End If
    Else: Exit Sub
    End If
    
     Application.FileDialog(msoFileDialogOpen).Title = "Select the RB_WM Hiring Report"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
    Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\2015 Attrition Report\"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        RBWMPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not RBWMPath Like "*RB_WM Hiring*" Then
            MsgBox "Please Select The Correct File!"
            Exit Sub
        End If
    End If
    
    monthstr = InputBox(Prompt:="Please input the year and month." & vbNewLine & "Format: Aug15", _
          Title:="Please input the year and month")
    monthstr = Left(monthstr, 3) & "'" & Right(monthstr, 2)
End Sub



Private Sub GenerateReport_Stage1_Click()

If P_HRPATH = "" Or P_SOPPATH = "" Or P_PRPATH = "" Or monthstr = "" Then
    MsgBox "Please Select the Input Files!"
    Exit Sub
End If

If savetopath = "" Then
    MsgBox "Please Select the Saving Path!"
    Exit Sub
End If
    

Application.DisplayAlerts = False
Set P_HRBK = Workbooks.Open(P_HRPATH, False)
Set P_SOPBK = Workbooks.Open(P_SOPPATH, False)
Set P_PRBK = Workbooks.Open(P_PRPATH, False)



For Each sht In P_PRBK.Worksheets
    sht.Visible = True
    If sht.Name = "FTE" Or sht.Name = "Tracker" Or sht.Name = "PBRM Table" Or sht.Name = "Summary" Or sht.Name = "Attrition" Then
        sht.Delete
    End If
    
Next
        

consoFTE
getResignRMInfo
getResignPBInfo
updateLeaver
getBankerStatus
If RBWMPath <> "" Then
    fillInExtraStatus
End If
P_PRBK.SaveAs savetopath & "\Attrition Report " & Left(monthstr, 3) & Right(monthstr, 2) & "(Stage1).xlsx"
Application.DisplayAlerts = False
P_PRBK.Close
P_SOPBK.Close
P_HRBK.Close


Application.DisplayAlerts = True
End Sub

Private Sub GenerateReport_Stage2_Click()


If savetopath = "" Then
    MsgBox "Please Select the Saving Path!"
    Exit Sub
End If


Application.FileDialog(msoFileDialogOpen).Title = "Select the amended attrition report"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' INIT PATH NEED  TO CHANGE TO G:\PLUS\SOP CHEAN UP FILES
    Application.FileDialog(msoFileDialogOpen).InitialFileName = "G:\Plus\2015 Attrition Report\(8) Aug'15 Attrition Report\"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        P_CUPATH = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not P_CUPATH Like "*Attrition Report*" Then
            MsgBox "Please Select The Correct File!"
            Exit Sub
        End If
    Else: Exit Sub
    End If
    
monthstr = InputBox(Prompt:="Please input the year and month." & vbNewLine & "Format: Aug15", _
          Title:="Please input the year and month")
If Len(monthstr) <> 5 Then
    MsgBox "Please input the correct month"
    Exit Sub
End If
monthstr = Left(monthstr, 3) & "'" & Right(monthstr, 2)

Set P_CUBK = Workbooks.Open(P_CUPATH, False)

FTEPage

TrackerPageForSummary
TrackerPage
PBRMTable
PT_Leaver
getBranchAttrition
getAttritionTable
createSummarySheet
orientSheet2
Set pvt = P_CUBK.Worksheets("Sales Tracker").PivotTables(1)
pvt.PivotCache.Refresh

P_CUBK.SaveAs savetopath & "\Attrition Report " & Left(monthstr, 3) & Right(monthstr, 2) & ".xlsx"
End Sub


Sub orientSheet2()
For Each sht In P_CUBK.Worksheets
    flag = 0
    If sht.Name = "FTE" Or sht.Name = "Attrition" Or sht.Name = "PBRM Table" Or sht.Name = "Sales Tracker" Or sht.Name = "Summary" Then
        flag = 1
    End If
    If flag = 0 Then
        sht.Visible = xlSheetHidden
    End If
Next
With P_CUBK
    .Sheets("Attrition").Move After:=.Sheets("Summary")
    .Sheets("FTE").Move After:=.Sheets("Attrition")
    .Sheets("Sales Tracker").Move After:=.Sheets("FTE")
    .Sheets("PBRM Table").Move After:=.Sheets("Sales Tracker")
End With
End Sub

Sub consoFTE()


Dim ftebk As Workbook
Set ftebk = Workbooks.Add
Dim hrbk As Workbook
Set hrbk = P_HRBK
Dim attbk As Workbook
Set attbk = P_PRBK
hrbk.Activate

hrbk.Worksheets("Active List").Copy ftebk.Worksheets(1)

ftebk.Worksheets("Active List").Activate
Dim colCount As Integer
Dim str As Variant
With ftebk.Worksheets("Active List")
    colCount = .UsedRange.Columns.count
    .Cells(1, colCount + 1).Value = "CC"
    .Cells(1, colCount + 2).Value = "Category"
    .Cells(1, colCount + 3).Value = "Description"
    .Cells(1, colCount + 4).Value = "BBD"
    .Cells(1, colCount + 5).Value = "Month"
    For i = 2 To .UsedRange.Rows.count
        .Cells(i, colCount + 1).Value = Right(.Cells(i, 4), 4)
        str = Application.VLookup(.Cells(i, 9).Value, ThisWorkbook.Worksheets("Category").UsedRange, 3, False)
        If IsError(str) Then
            .Cells(i, colCount + 2).Value = "ERROR"
        Else
            .Cells(i, colCount + 2).Value = CStr(str)
        End If
        str = Application.VLookup(.Cells(i, colCount + 1).Value, ThisWorkbook.Worksheets("Description").Range("B:D"), 3, False)
        If IsError(str) Then
            .Cells(i, colCount + 3).Value = "ERROR"
        Else
            .Cells(i, colCount + 3).Value = CStr(str)
        End If
        str = Application.VLookup(.Cells(i, colCount + 3).Value, ThisWorkbook.Worksheets("Zone Map").Range("A:B"), 2, False)
        If IsError(str) Then
            .Cells(i, colCount + 4).Value = "ERROR"
        Else
            .Cells(i, colCount + 4).Value = CStr(str)
        End If
        .Cells(i, colCount + 5).Value = monthstr
    Next
End With

'IBCpage
'MBpage
'Servicepage


Dim atrowNo, rowct As Integer
Set consosht = attbk.Worksheets("Active List Conso")
Set listsht = ftebk.Worksheets("Active List")
atrowNo = consosht.Range("A1").End(xlDown).Row
rowct = listsht.Range("A1").End(xlDown).Row
For i = 1 To consosht.UsedRange.Columns.count
    For j = 1 To listsht.UsedRange.Columns.count
        If listsht.Cells(1, j) = consosht.Cells(1, i) Then
            With listsht
            .Activate
            .Range(.Cells(2, j), .Cells(rowct, j)).Select
            Selection.Copy
            End With
            With consosht
            .Activate
            .Cells(atrowNo + 1, i).Select
            Selection.PasteSpecial xlPasteValues
            End With
        End If
    Next
            
Next

Application.DisplayAlerts = False
ftebk.Close
Application.DisplayAlerts = True

End Sub
  
Private Function getColString(colidx As Integer) As String
Dim str As String
str = Cells(1, colidx).Address
str = Mid(str, 2, Len(str) - 3)
getColString = str
End Function

Sub getResignPBInfo()
Dim SelectionStr1 As String
SelectionStr1 = "Yes, No, unknown"
Dim SelectionStr2 As String
For i = 2 To ThisWorkbook.Worksheets("FTE Config").Range("D1").End(xlDown).Row
    If i <> 2 Then SelectionStr2 = SelectionStr2 + ", "
    SelectionStr2 = SelectionStr2 + ThisWorkbook.Worksheets("FTE Config").Cells(i, 4)
Next
Dim SelectionStr3 As String
For i = 2 To ThisWorkbook.Worksheets("FTE Config").Range("E1").End(xlDown).Row
    If i <> 2 Then SelectionStr3 = SelectionStr3 + ", "
    SelectionStr3 = SelectionStr3 + ThisWorkbook.Worksheets("FTE Config").Cells(i, 5)
Next

Dim path1 As String
Dim path2 As String
Dim hrbk As Workbook
Dim sopbk As Workbook
Set hrbk = P_HRBK
Set sopbk = P_SOPBK
Set appbk = P_PRBK
Set termsht = hrbk.Worksheets("Termination")
Dim pbsht As Worksheet
For Each sht In sopbk.Worksheets
    If sht.Name Like "??? PB" Then
    Set pbsht = sht
    End If
Next
termsht.Activate
If termsht.UsedRange.Rows.count = 1 Then
Exit Sub
End If
For Each sht In appbk.Worksheets
    If sht.Name Like "PB Resignation*" Then
    Set resultsht = sht
    End If
Next
Dim count As Integer
count = resultsht.Range("A1").End(xlDown).Row
Dim geid As String
For i = 2 To termsht.UsedRange.Rows.count
    termsht.Activate
    If CStr(Format(termsht.Cells(i, 3).Value, "MM")) = "08" And CStr(termsht.Cells(i, 6).Value) = "Personal Banker" Then
        geid = CStr(termsht.Cells(i, 1).Value)
        pbsht.Activate
        For j = 2 To pbsht.UsedRange.Rows.count
            If CStr(pbsht.Cells(j, 8).Value) = geid Then
                count = count + 1
                resultsht.Cells(count, 1).Value = CStr(Format(termsht.Cells(i, 3).Value, "YYYY"))
                resultsht.Cells(count, 2).Value = CStr(pbsht.Cells(j, 2).Value)
                resultsht.Cells(count, 3).Value = CStr(pbsht.Cells(j, 4).Value)
                resultsht.Cells(count, 4).Value = CStr(pbsht.Cells(j, 7).Value)
                resultsht.Cells(count, 5).Value = CStr(pbsht.Cells(j, 11).Value)
                resultsht.Cells(count, 6).Value = CStr(pbsht.Cells(j, 12).Value)
                resultsht.Cells(count, 7).Value = CStr(pbsht.Cells(j, 13).Value)
                resultsht.Cells(count, 8).Value = termsht.Cells(i, 3).Value
                resultsht.Cells(count, 9).Value = CStr(pbsht.Cells(j, 14).Value)
                resultsht.Cells(count, 10).Value = "'" & monthstr
                resultsht.Cells(count, 14).Value = CStr(pbsht.Cells(j, 19).Value)
                resultsht.Cells(count, 15).Value = CStr(pbsht.Cells(j, 27).Value)
                If resultsht.Cells(count, 15).Value = "" Then resultsht.Cells(count, 15).Value = "Non-Regrettable"
                resultsht.Cells(count, 16).Value = getQuarter()
                resultsht.Cells(count, 6).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 7).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 8).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 9).NumberFormat = "0.0"
                
            End If
        Next
    End If
    
Next
End Sub

Sub getResignRMInfo()
Dim SelectionStr1 As String
SelectionStr1 = "Yes, No, unknown"
Dim SelectionStr2 As String
For i = 2 To ThisWorkbook.Worksheets("FTE Config").Range("D1").End(xlDown).Row
    If i <> 2 Then SelectionStr2 = SelectionStr2 + ", "
    SelectionStr2 = SelectionStr2 + ThisWorkbook.Worksheets("FTE Config").Cells(i, 4)
Next
Dim SelectionStr3 As String
For i = 2 To ThisWorkbook.Worksheets("FTE Config").Range("E1").End(xlDown).Row
    If i <> 2 Then SelectionStr3 = SelectionStr3 + ", "
    SelectionStr3 = SelectionStr3 + ThisWorkbook.Worksheets("FTE Config").Cells(i, 5)
Next


Dim path1 As String
Dim path2 As String

Dim hrbk As Workbook
Dim sopbk As Workbook
Set hrbk = P_HRBK
Set sopbk = P_SOPBK
Set appbk = P_PRBK
hrbk.Activate
Set termsht = hrbk.Worksheets("Termination")
Dim rmsht As Worksheet
For Each sht In sopbk.Worksheets
    If sht.Name Like "??? RM" Then
    Set rmsht = sht
    End If
Next
termsht.Activate
If termsht.UsedRange.Rows.count = 1 Then
Exit Sub
End If
For Each sht In appbk.Worksheets
    If sht.Name Like "RM Resignation*" Then
    Set resultsht = sht
    End If
Next
Dim count As Integer
count = resultsht.Range("A1").End(xlDown).Row
Dim geid As String
For i = 2 To termsht.UsedRange.Rows.count
    termsht.Activate
    If CStr(Format(termsht.Cells(i, 3).Value, "MM")) = "08" And (CStr(termsht.Cells(i, 6).Value) = "Relationship Manager" Or CStr(termsht.Cells(i, 6).Value) = "Associate Banker, CPC" Or CStr(termsht.Cells(i, 6).Value) = "Direct Banking Specialist") Then
        geid = CStr(termsht.Cells(i, 1).Value)
        rmsht.Activate
        For j = 2 To rmsht.UsedRange.Rows.count
            If CStr(rmsht.Cells(j, 8).Value) = geid Then
                count = count + 1
                resultsht.Cells(count, 1).Value = CStr(Format(termsht.Cells(i, 3).Value, "YYYY"))
                resultsht.Cells(count, 2).Value = CStr(rmsht.Cells(j, 2).Value)
                resultsht.Cells(count, 3).Value = CStr(rmsht.Cells(j, 4).Value)
                resultsht.Cells(count, 4).Value = CStr(rmsht.Cells(j, 7).Value)
                resultsht.Cells(count, 5).Value = CStr(rmsht.Cells(j, 11).Value)
                resultsht.Cells(count, 6).Value = CStr(rmsht.Cells(j, 12).Value)
                resultsht.Cells(count, 7).Value = CStr(rmsht.Cells(j, 13).Value)
                resultsht.Cells(count, 8).Value = termsht.Cells(i, 3).Value
                resultsht.Cells(count, 9).Value = CStr(rmsht.Cells(j, 14).Value)
                resultsht.Cells(count, 10).Value = "'" & monthstr
                resultsht.Cells(count, 11).Value = "RM"
                If rmsht.Cells(j, 8).Font.Color <> vbBlack Then
                    resultsht.Cells(count, 11).Value = "CPC RM"
                    resultsht.Cells(count, 11).Font.Color = vbRed
                End If
                If CStr(termsht.Cells(i, 6).Value) = "Associate Banker, CPC" Then
                    resultsht.Cells(count, 11).Value = "AB"
                    resultsht.Cells(count, 11).Font.Color = vbRed
                End If
                If CStr(termsht.Cells(i, 6).Value) = "Direct Banking Specialist" Then
                    resultsht.Cells(count, 11).Value = "DB Specialist"
                    resultsht.Cells(count, 11).Font.Color = vbRed
                End If
                resultsht.Cells(count, 15).Value = CStr(rmsht.Cells(j, 19).Value)
                resultsht.Cells(count, 16).Value = CStr(rmsht.Cells(j, 27).Value)
                If resultsht.Cells(count, 16).Value = "" Then resultsht.Cells(count, 16).Value = "Non-Regrettable"
                resultsht.Cells(count, 17).Value = getQuarter()
                resultsht.Cells(count, 6).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 7).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 8).NumberFormat = "DD-MMM-YY"
                resultsht.Cells(count, 9).NumberFormat = "0.0"
                
            End If
        Next
    End If
    
Next
End Sub

Private Function getRowIndexes(sht As Worksheet, colidx As Integer) As Variant
Dim curstr As String
Dim idxarr(6) As String
For i = 1 To sht.UsedRange.Rows.count
    If CStr(sht.Cells(i, colidx).Value) = "Within Ind" Then
        curstr = "Within Ind"
        idxarr(1) = i + 1
    End If
    If CStr(sht.Cells(i, colidx).Value) = "Bank" Then
        curstr = "Bank"
        idxarr(3) = i + 1
    End If
    If CStr(sht.Cells(i, colidx).Value) = "Reason Specific" Then
        curstr = "Reason Specific"
        idxarr(5) = i + 1
    End If
    If CStr(sht.Cells(i, colidx).Value) = "Grand Total" Then
        If curstr = "Within Ind" Then
             idxarr(2) = i - 1
        End If
        If curstr = "Bank" Then
             idxarr(4) = i - 1
        End If
        If curstr = "Reason Specific" Then
             idxarr(6) = i - 1
        End If
    End If
        
Next
getRowIndexes = idxarr
End Function






Sub getBankerStatus()

Qstr = getQuarter()

Dim sopbk, rptbk As Workbook
Set sopbk = P_SOPBK
Set rptbk = P_PRBK
' abk is the workbook the attach, output file
Set asht = rptbk.Worksheets("Banker Status")
Dim rmsht As Worksheet
Dim pbsht As Worksheet
For Each sht In sopbk.Worksheets
    If sht.Name Like ("??? RM") Then
        Set rmsht = sht
    End If
    If sht.Name Like ("??? PB") Then
        Set pbsht = sht
    End If
Next
Dim count As Integer
' should be equal to the count of previous month records. appending
count = asht.UsedRange.Rows.count - 1

'RM
For i = 4 To rmsht.Range("A1").End(xlDown).Row
    count = count + 1
    asht.Cells(count + 1, 1) = rmsht.Cells(i, 4)
    asht.Cells(count + 1, 2) = rmsht.Cells(i, 7)
    'job type
    If rmsht.Cells(i, 7).Font.Color <> vbBlack Then
        asht.Cells(count + 1, 3).Value = "CPC RM"
    End If
    If rmsht.Cells(i, 7).Interior.Color = vbYellow Then
        asht.Cells(count + 1, 3).Value = "AB"
    End If
    If rmsht.Cells(i, 7).Font.Color = vbBlack Then
        asht.Cells(count + 1, 3).Value = "RM"
    End If
    ' status
    If rmsht.Cells(i, 22).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Unranked"
    ElseIf rmsht.Cells(i, 24).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Training"
    ElseIf rmsht.Cells(i, 25).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Transfer"
    ElseIf rmsht.Cells(i, 26).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Resigned"
    ElseIf rmsht.Cells(i, 23).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Ranked"
    Else
        asht.Cells(count + 1, 4).Value = "ERROR"
    End If
    'promotion
    If rmsht.Cells(i, 30).Value = "Y" Then
        asht.Cells(count + 1, 5).Value = "CG to CPC"
    ElseIf rmsht.Cells(i, 29).Value = "Y" Then
        asht.Cells(count + 1, 5).Value = "PB to RM"
    End If
    asht.Cells(count + 1, 6).Value = "'" & Left(monthstr, 3) & Right(monthstr, 2)
    asht.Cells(count + 1, 7).Value = Qstr
    
Next
'PB

For i = 4 To pbsht.Range("A1").End(xlDown).Row
    count = count + 1
    asht.Cells(count + 1, 1) = pbsht.Cells(i, 4)
    asht.Cells(count + 1, 2) = pbsht.Cells(i, 7)
    'job type
    If pbsht.Cells(i, 4).Value = "Direct Banking Specialist" Or pbsht.Cells(i, 4).Value Like "DBB*" Then
        asht.Cells(count + 1, 3).Value = "DB Specialist"
    Else
        asht.Cells(count + 1, 3).Value = "PB"
    End If
    ' status
    If pbsht.Cells(i, 22).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Unranked"
    ElseIf pbsht.Cells(i, 24).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Training"
    ElseIf pbsht.Cells(i, 25).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Transfer"
    ElseIf pbsht.Cells(i, 26).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Resigned"
    ElseIf pbsht.Cells(i, 23).Value = "Y" Then
        asht.Cells(count + 1, 4).Value = "Ranked"
    Else
        asht.Cells(count + 1, 4).Value = "ERROR"
    End If
    'promotion
    If pbsht.Cells(i, 29).Value = "Y" Then
        asht.Cells(count + 1, 5).Value = "PB to RM"
    End If
    asht.Cells(count + 1, 6).Value = "'" & Left(monthstr, 3) & Right(monthstr, 2)
    asht.Cells(count + 1, 7).Value = Qstr
    
Next


End Sub

Private Sub TrackerPage()
Dim trackersht As Worksheet
Set trackersht = P_CUBK.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = P_CUBK.Worksheets("Banker Status")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim rangeStr As String
Dim StartPvt As String
Dim NumofRows As String


On Error Resume Next

NumofRows = datasht.UsedRange.Rows.count

rangeStr = "A1:G" & NumofRows

SrcData = datasht.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)

trackersht.Name = "Sales Tracker"
StartPvt = "'" & trackersht.Name & "'" & "!" & trackersht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = P_CUBK.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  
  pvt.AddDataField pvt.PivotFields("Status"), "Count of Status", xlCount
  pvt.AddFields RowFields:=Array("Quarter", "Month"), ColumnFields:="Status"
  pvt.PivotFields("Branch").Orientation = xlPageField
  pvt.PivotFields("Type").Orientation = xlPageField
  For Each itm In pvt.PivotFields("Quarter").PivotItems
    If itm.Name = "(blank)" Then
        itm.Visible = False
    End If
  Next
  pvt.PivotFields("Type").EnableMultiplePageItems = True
  pvt.PivotCache.Refresh
End Sub

Private Sub TrackerPageForSummary()
Dim trackersht As Worksheet
Set trackersht = P_CUBK.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = P_CUBK.Worksheets("Banker Status")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim rangeStr As String
Dim StartPvt As String
Dim NumofRows As String


On Error Resume Next

NumofRows = datasht.UsedRange.Rows.count

rangeStr = "A1:G" & NumofRows

SrcData = datasht.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)

trackersht.Name = "Tracker(S)"
StartPvt = "'" & trackersht.Name & "'" & "!" & trackersht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = P_CUBK.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    
  
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlTabularRow
   pvt.ColumnGrand = True
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = False
    pvt.EnableDrilldown = True
   
  
  pvt.AddDataField pvt.PivotFields("Status"), "Count of Status", xlCount
  pvt.AddFields RowFields:="Month", ColumnFields:=Array("Status", "Type")
  
End Sub

Sub updateLeaver()


Dim hrbk, rptbk As Workbook
Dim tsht, asht, activesht As Worksheet
Set hrbk = P_HRBK
Set rptbk = P_PRBK
Set tsht = hrbk.Worksheets("Termination")
Set asht = rptbk.Worksheets("Leavers")
Set activesht1 = ThisWorkbook.Worksheets("Active List 1")
Set activesht2 = ThisWorkbook.Worksheets("Active List 2")
Dim count, sum As Integer
Dim descr As String
count = asht.Range("A1").End(xlDown).Row
sum = 0
Dim test As Variant
Dim test2 As Variant
For i = 2 To tsht.Range("A1").End(xlDown).Row
    If Format(tsht.Cells(i, 4).Value, "MMM") = Left(monthstr, 3) Then
     
        count = count + 1
        asht.Cells(count, 1) = tsht.Cells(i, 1)
        asht.Cells(count, 2) = tsht.Cells(i, 2)
        asht.Cells(count, 3) = tsht.Cells(i, 4)
        asht.Cells(count, 3).NumberFormat = "M/D/YYYY"
        asht.Cells(count, 4) = monthstr
        asht.Cells(count, 5) = tsht.Cells(i, 3)
        asht.Cells(count, 5).NumberFormat = "M/D/YYYY"
        
        asht.Cells(count, 6) = tsht.Cells(i, 6)
        test = Application.VLookup(asht.Cells(count, 1).Value, activesht1.UsedRange, 5, False)
        test2 = Application.VLookup(asht.Cells(count, 1).Value, activesht2.UsedRange, 5, False)
        If IsError(test) And IsError(test2) Then  ' which means cannot find in previous month activelist
            If asht.Cells(count, 6) Like "*Direct Banking*" Then
                asht.Cells(count, 8) = "Direct Banking Specialist"
                For p = 9 To 12
                    asht.Cells(count, p) = "#N/A"
                Next
            ElseIf asht.Cells(count, 6) Like "*Bancassurance*" Then
                asht.Cells(count, 8) = "BANK ASSURANCE"
                For p = 9 To 12
                    asht.Cells(count, p) = "#N/A"
                Next
            Else
                For p = 8 To 12
                    asht.Cells(count, p) = "#N/A"
                    asht.Cells(count, p).Interior.Color = vbRed
                Next
            End If
        Else
            If IsError(test) Then test = test2
            descr = CStr(test)
            asht.Cells(count, 8) = descr
        
            test = Application.VLookup(descr, ThisWorkbook.Worksheets("Leaver Map").UsedRange, 2, False)
            If IsError(test) Then
                ' no branch maped
                asht.Cells(count, 9) = "#N/A"
                asht.Cells(count, 10) = "#N/A"
                asht.Cells(count, 11) = "#N/A"
                asht.Cells(count, 12) = "#N/A"
            Else
                asht.Cells(count, 9) = CStr(test)
                test = Application.VLookup(asht.Cells(count, 6).Value, ThisWorkbook.Worksheets("Category").UsedRange, 3, False)
                If IsError(test) Then
                    asht.Cells(count, 10) = "#N/A"
                Else
                    asht.Cells(count, 10) = CStr(test)
                End If
                asht.Cells(count, 11) = monthstr
           
                test = Application.VLookup(asht.Cells(count, 9).Value, ThisWorkbook.Worksheets("Zone Map").UsedRange, 2, False)
                If IsError(test) Then
                    asht.Cells(count, 12) = "#N/A"
                Else
                    asht.Cells(count, 12) = CStr(test)
                    asht.Cells(count, 13) = 1
                    sum = sum + 1
                End If
            End If
        End If
    End If
Next
asht.Cells(count, 14) = sum
asht.Cells(count, 14).Interior.Color = RGB(255, 204, 255)


End Sub

Private Sub PT_Leaver()
Dim trackersht As Worksheet
Set trackersht = P_CUBK.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = P_CUBK.Worksheets("Leavers")
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim rangeStr As String
Dim StartPvt As String
Dim NumofRows As String


On Error Resume Next

NumofRows = datasht.UsedRange.Rows.count

rangeStr = "A1:" & "M" & NumofRows

SrcData = datasht.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)

trackersht.Name = "Attrition"
StartPvt = "'" & trackersht.Name & "'" & "!" & trackersht.Range("A1").Address(ReferenceStyle:=xlR1C1)

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
   
  
  pvt.AddDataField pvt.PivotFields("Eff Date"), "Count of Eff Date", xlCount
  pvt.AddFields RowFields:=Array("Zone", "Description"), ColumnFields:="Category"
  pvt.PivotFields("Month").Orientation = xlPageField
  For Each itm In pvt.PivotFields("Zone").PivotItems
    If IsError(itm.Name) Then
        itm.Visible = False
    ElseIf itm.Name = "(blank)" Or itm.Name = "Specialist" Or itm.Name = "#N/A" Then
        itm.Visible = False
    End If
  Next
End Sub

Sub getBranchAttrition()

Set attsht = P_CUBK.Worksheets("Attrition")
Set pvt = attsht.PivotTables(1)
curMonth = Month("1 " & Left(monthstr, 3))
Set ftesht = P_CUBK.Worksheets("FTE")
Set fpvt = ftesht.PivotTables(1)

startRow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.count).Row + 3
pvt.TableRange1.Copy
attsht.Cells(startRow, 1).PasteSpecial xlPasteValues
attsht.Range(attsht.Cells(startRow, 3), attsht.Cells(attsht.UsedRange.Rows.count, 8)).Clear
attsht.Cells(startRow + 1, 3) = "Attrition Rate"
Dim arr1, arr2 As Variant
Dim ar(4) As Double
Application.DisplayAlerts = False
'On Error Resume Next
For i = 5 To startRow - 5
    
    For j = 8 To 8
        zone = attsht.Cells(i, 1)
        tmp = i
        While zone = ""
            zone = attsht.Cells(tmp - 1, 1)
            tmp = tmp - 1
        Wend
        br = attsht.Cells(i, 2)
        job = attsht.Cells(4, j)
        If br <> "" Then
            attcount = attsht.Cells(i, j)
            
            'new code. break down by quarter
            attsht.Cells(i, j).Select
            Selection.ShowDetail = True
            
            Set tmpsht = ActiveSheet
            arr1 = getQuarterData(monthstr)
            tmpsht.Delete
            'ActiveSheet.Delete
            ' get FTE for the cell
            termMonth = ""
            
            With ThisWorkbook.Worksheets("Zone Map")
                For k = 2 To .UsedRange.Rows.count
                     If Not IsError(.Cells(k, 2).Value) Then
                        If .Cells(k, 1).Value = "Resigned" And .Cells(k, 2).Value = zone Then
                            termMonth = .Cells(k, 3).Value
                        End If
                    End If
                Next
            End With
            
            If termMonth = "" Then
                termMonth = monthstr
            End If
            
            curMonth = Month("1 " & Left(termMonth, 3))
            
            ftesht.Activate
            fpvt.PivotFields("Month").CurrentPage = "All"
            fpvt.PivotSelect "BBD[" & zone & "]", xlDataOnly, True
            Set selection1 = Selection
            fpvt.PivotSelect "Description[" & br & "]", xlDataOnly, True
            Set selection2 = Selection
            If j <> 8 Then
                fpvt.PivotSelect "Category[" & job & "]", xlDataOnly, True
            Else
                ftesht.Range("H:H").Select
            End If
            Set selection3 = Selection
            Intersect(selection1, selection2, selection3).Select
            'new
            Selection.ShowDetail = True
            Set tmpsht = ActiveSheet
            arr2 = getQuarterData(monthstr)
            tmpsht.Delete
            For k = 0 To 3
                If curMonth < (k + 1) * 3 Then
                    ar(k) = arr2(k) / (curMonth - k * 3)
                Else
                    
                    ar(k) = arr2(k) / 3
                End If
            Next
                   
                   
            'test
            If br = "CS" Then sa = ""
             sa = ""
            sb = ""
            
            For k = 0 To 3
                sa = sa + " " + CStr(arr1(k))
                sb = sb + " " + CStr(ar(k))
            Next
            
           ' MsgBox sa & vbNewLine & sb
            
            attrate = 0
            qtr = 0
            For k = 0 To 3
                If arr2(k) <> 0 Then
                    attrate = attrate + arr1(k) / ar(k)
                    qtr = qtr + 1
                End If
            Next
            attrate = attrate / qtr * 4
            
            attsht.Activate
            attsht.Cells(startRow + i - 3, 3) = attrate
        End If
               
    Next
Next

i = startRow + 1
While (attsht.Cells(i, 1) <> "") Or (attsht.Cells(i, 3) <> "")
    If attsht.Cells(i, 2) = "" Then
        attsht.Cells(i, 2).EntireRow.Delete
        i = i - 1
    Else
        i = i + 1
    End If
Wend


'formatting
With attsht
    .Cells(startRow, 1) = "Branch Annualized Attrition Rate"
    .Range(.Cells(startRow, 1), .Cells(startRow + 1, 3)).Interior.Color = RGB(220, 230, 241)
    .Range(.Cells(startRow, 1), .Cells(startRow + 1, 3)).Font.Bold = True
    .Range(.Cells(startRow + 1, 1), .Cells(startRow + 1, 3)).Borders(xlBottom).Color = RGB(83, 141, 213)
    For i = startRow + 2 To .Cells(startRow + 2, 2).End(xlDown).Row
    If .Cells(i, 1) <> "" And .Cells(i + 1, 1) = "" Then
        .Range(.Cells(i, 1), .Cells(i, 3)).Borders(xlTop).Color = RGB(83, 141, 213)
    End If
    Next
    .Range(.Cells(i, 1), .Cells(i, 3)).Borders(xlTop).Color = RGB(83, 141, 213)
    .Range(.Cells(startRow + 1, 1), .Cells(.Cells(startRow + 2, 2).End(xlDown).Row, 1)).Font.Bold = True
    .Range(.Cells(startRow + 2, 3), .Cells(.Cells(startRow + 2, 2).End(xlDown).Row, 3)).NumberFormat = "0.00%"
End With


Application.DisplayAlerts = True
End Sub

Sub tset()
    Selection.Interior.Color = RGB(220, 230, 241)
End Sub

Sub getAttritionTable()
Dim yearstr As String


yearstr = monthstr
Set attsht = P_CUBK.Worksheets("Attrition")
Dim startRow, zoneCount As Integer
zoneCount = 0
Dim range1 As Range
Dim range2 As Range
Dim colidx, subsum, bbdrow, catCount As Integer
catCount = 0
Dim arr As Variant


'--  fill in Quarter + total attrition for zone
Set pvt = attsht.PivotTables(1)
startRow = attsht.UsedRange.Rows.count + 3

'attsht.Rows(CStr(startRow) & ":" & CStr(attsht.UsedRange.Rows.count)).Delete

For Each itm In pvt.PivotFields("Zone").PivotItems
    If itm.Visible = True Then
        zoneCount = zoneCount + 1
         pvt.PivotSelect "Zone[" & itm.Name & ";Total]", xlDataOnly, True
        Set range1 = Selection
        For Each itm2 In pvt.PivotFields("Category").PivotItems
            If itm2.Visible = True And itm2.Name <> "#N/A" Then
                catCount = catCount + 1
                pvt.PivotSelect "Category[" & itm2.Name & "]", xlDataOnly, True
                Set range2 = Selection
                Intersect(range1, range2).Select
                colidx = Selection.Column
                subsum = Selection.Value
                Selection.ShowDetail = True
                Set tmpsht = ActiveSheet
                arr = getQuarterData(monthstr)
                For k = 0 To 3
                    attsht.Cells(startRow + (zoneCount - 1) * 15 + 2 + k, colidx) = arr(k)
                Next
                attsht.Cells(startRow + (zoneCount - 1) * 15 + 6, colidx) = subsum
                attsht.Cells(startRow + (zoneCount - 1) * 15 + 2, 1) = itm.Name
                Application.DisplayAlerts = False
                tmpsht.Delete
                attsht.Activate
                Application.DisplayAlerts = True
            End If
        Next
    End If
Next
catCount = Int(catCount / zoneCount)
' Fill in Quarter/ total FTE for zone
Set ftesht = ActiveWorkbook.Worksheets("FTE")
Set pvt = ftesht.PivotTables(1)
For Each itm In pvt.PivotFields("BBD").PivotItems
Dim curMonth As Integer
Dim termMonth As String
termMonth = ""
curMonth = Month("1 " & Left(monthstr, 3))

    If itm.Visible = True Then
    If itm.Name = "John" Then
        termMonth = ""
    End If
    
    
        Set c = attsht.Range(attsht.Cells(startRow, 1), attsht.Cells(attsht.UsedRange.Rows.count, 1)).Find(itm.Name)
        If c Is Nothing Then
            Exit For
        Else
            bbdrow = c.Row
        End If
        
        With ThisWorkbook.Worksheets("Zone Map")
            For k = 2 To .UsedRange.Rows.count
                If Not IsError(.Cells(k, 2).Value) Then
                    If .Cells(k, 2).Value = itm.Name And .Cells(k, 1).Value = "Resigned" Then
                        termMonth = .Cells(k, 3).Value
                    End If
                End If
            Next
        End With
        
        If termMonth <> "" Then
            pvt.PivotFields("Month").CurrentPage = termMonth
        End If
        

            If termMonth = "" Then
                termMonth = monthstr
            End If
            
        curMonth = Month("1 " & Left(termMonth, 3))
        ftesht.Activate
        'pvt.PivotSelect "BBD[" & itm.Name & ";Total]", xlDataOnly, True
        'Set range1 = Selection
        For Each itm2 In pvt.PivotFields("Category").PivotItems
            If itm2.Visible = True And itm2.Name <> "#N/A" Then
                pvt.PivotFields("Month").CurrentPage = "All"
                ftesht.Activate
                pvt.PivotSelect "BBD[" & itm.Name & ";Total]", xlDataOnly, True
                Set range1 = Selection
                pvt.PivotSelect "Category[" & itm2.Name & "]", xlDataOnly, True
                Set range2 = Selection
                Intersect(range1, range2).Select
                colidx = Selection.Column
                subsum = Selection.Value
                Selection.ShowDetail = True
                Set tmpsht = ActiveSheet
                arr = getQuarterData(monthstr)
                For k = 0 To 3
                    If curMonth < (k + 1) * 3 Then
                        attsht.Cells(bbdrow + 5 + k, colidx) = arr(k) / (curMonth - k * 3)
                    Else
                        attsht.Cells(bbdrow + 5 + k, colidx) = arr(k) / 3
                    End If
                Next
                
                If termMonth = "" Then
                    pvt.PivotFields("Month").CurrentPage = monthstr
                Else
                    pvt.PivotFields("Month").CurrentPage = termMonth
                End If
                ftesht.Activate
                pvt.PivotSelect "Category[" & itm2.Name & "]", xlDataOnly, True
                Set range2 = Selection
                pvt.PivotSelect "BBD[" & itm.Name & ";Total]", xlDataOnly, True
                Set range1 = Selection
                Intersect(range1, range2).Select
                subsum = Selection.Value
                attsht.Cells(bbdrow + 9, colidx) = subsum
                
                Application.DisplayAlerts = False
                tmpsht.Delete
                attsht.Activate
                Application.DisplayAlerts = True
            End If
        Next
    End If
Next
'FIll in last column : Total col
Dim ssum, ssum1 As Integer
ssum = 0
For i = 1 To zoneCount
    bbdrow = startRow + 2 + (i - 1) * 15
    For j = 0 To 9
        ssum = 0
        For k = 1 To catCount
            ssum = ssum + attsht.Cells(bbdrow + j, k + 2)
        Next
        attsht.Cells(bbdrow + j, catCount + 3) = ssum
    Next
Next


'Fill in Quarter / annualized attrition rate
For i = 1 To zoneCount
    
    termMonth = ""
    bbdrow = startRow + 2 + (i - 1) * 15
    
    With ThisWorkbook.Worksheets("Zone Map")
        For k = 2 To .UsedRange.Rows.count
            If Not IsError(.Cells(k, 2).Value) Then
                If .Cells(k, 2).Value = attsht.Cells(bbdrow, 1).Value And .Cells(k, 1).Value = "Resigned" Then
                    termMonth = .Cells(k, 3).Value
                End If
            End If
        Next
    End With
    
    
    
    For k = 1 To catCount + 1
        For j = 1 To 4
        If attsht.Cells(bbdrow + 4 + j, k + 2) <> 0 Then
            attsht.Cells(bbdrow + 9 + j, k + 2) = attsht.Cells(bbdrow - 1 + j, k + 2) / attsht.Cells(bbdrow + 4 + j, k + 2)
        Else
            attsht.Cells(bbdrow + 9 + j, k + 2) = 0
        End If
        Next
        If attsht.Cells(bbdrow + 9, k + 2) <> 0 Then
            'If termMonth = "" Then
            '    attsht.Cells(bbdrow + 14, k + 2) = attsht.Cells(bbdrow + 4, k + 2) / attsht.Cells(bbdrow + 9, k + 2) / curMonth * 12
            'Else
            '    attsht.Cells(bbdrow + 14, k + 2) = attsht.Cells(bbdrow + 4, k + 2) / attsht.Cells(bbdrow + 9, k + 2) / getMonthNumber(Left(termMonth, 3)) * 12
            'End If
            rt = 0
            ct = 0
            For p = 0 To 3
                If attsht.Cells(bbdrow + 5 + p, k + 2) <> 0 Then
                    rt = rt + attsht.Cells(bbdrow + p, k + 2) / attsht.Cells(bbdrow + 5 + p, k + 2)
                    ct = ct + 1
                End If
            Next
            attsht.Cells(bbdrow + 14, k + 2) = rt / ct * 4
            
            
        Else
            attsht.Cells(bbdrow + 14, k + 2) = 0
        End If
    Next
Next

'total summary
Set ftesht = P_CUBK.Worksheets("FTE")

ftesht.PivotTables(1).PivotFields("Month").CurrentPage = monthstr
Dim totrow As Integer
totrow = startRow + 2 + zoneCount * 15
With attsht
    .Cells(totrow, 1) = "Total"
    .Cells(totrow, 2) = "Attrition"
    .Cells(totrow + 1, 2) = "FTE"
    .Cells(totrow + 2, 2) = "Annualized Attrition Rate"
    For i = 1 To catCount + 1
        ssum = 0
        ssum1 = 0
        For j = 1 To zoneCount
            bbdrow = startRow + 2 + (j - 1) * 15
            ssum = ssum + .Cells(bbdrow + 4, i + 2)
        Next
        .Cells(totrow, i + 2) = ssum
        .Cells(totrow + 1, i + 2) = ftesht.Cells(ftesht.UsedRange.Rows.count, i + 2)
        ssum1 = .Cells(totrow + 1, i + 2)
        .Cells(totrow + 2, i + 2) = ssum / ssum1 / curMonth * 12
    Next
End With


' formatting
With attsht
.Cells(startRow, 1) = "Annualized 20" & Right(monthstr, 2) & " Attrition Rate - breakdown by new zone structure"
.Range(.Cells(startRow, 1), .Cells(startRow, catCount + 3)).Merge
.Cells(startRow, 1).Font.Bold = True
For i = 1 To catCount
    .Cells(startRow + 1, i + 2) = attsht.Cells(4, i + 2)
Next
.Cells(startRow + 1, catCount + 3) = "Total"
.Range(.Cells(startRow + 1, 3), .Cells(startRow + 1, catCount + 3)).Interior.Color = RGB(217, 217, 217)
For i = 1 To zoneCount
    bbdrow = startRow + 2 + (i - 1) * 15
    For q = 1 To 4
        .Cells(bbdrow - 1 + q, 2) = "Attrition Q" & CStr(q)
        .Cells(bbdrow + 4 + q, 2) = "FTE Q" & CStr(q)
        .Cells(bbdrow + 9 + q, 2) = "Quarterly Attrition Rate Q" & CStr(q)
    Next
    .Cells(bbdrow + 4, 2) = "Attrition YTD"
    .Cells(bbdrow + 9, 2) = "FTE"
    .Range(.Cells(bbdrow + 9, 2), .Cells(bbdrow + 9, catCount + 3)).Interior.Color = RGB(218, 238, 243)
    .Cells(bbdrow + 14, 2) = "Annualized Attrition Rate"
    .Range(.Cells(bbdrow + 14, 2), .Cells(bbdrow + 14, catCount + 3)).Interior.Color = RGB(255, 204, 255)
Next
.Range(.Cells(totrow + 1, 2), .Cells(totrow + 1, catCount + 3)).Interior.Color = RGB(218, 238, 243)
.Range(.Cells(totrow + 2, 2), .Cells(totrow + 2, catCount + 3)).Interior.Color = RGB(255, 204, 255)

For i = 1 To zoneCount
    bbdrow = startRow + 2 + (i - 1) * 15
    .Range(.Cells(bbdrow, 3), .Cells(bbdrow + 9, catCount + 3)).NumberFormat = "0"
    .Range(.Cells(bbdrow + 10, 3), .Cells(bbdrow + 14, catCount + 3)).NumberFormat = "0%"
    .Range(.Cells(bbdrow, 1), .Cells(bbdrow + 14, 1)).Merge
    .Cells(bbdrow, 1).VerticalAlignment = xlTop
    .Cells(bbdrow, 1).HorizontalAlignment = xlCenter
Next
.Range(.Cells(totrow, 3), .Cells(totrow + 1, catCount + 3)).NumberFormat = "0"
.Range(.Cells(totrow + 2, 3), .Cells(totrow + 2, catCount + 3)).NumberFormat = "0%"
.Range(.Cells(totrow, 1), .Cells(totrow + 2, 1)).Merge
.Cells(totrow, 1).VerticalAlignment = xlTop
.Cells(totrow, 1).HorizontalAlignment = xlCenter

.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlTop).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlBottom).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlRight).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(totrow + 2, catCount + 3)).Borders(xlLeft).LineStyle = XlLineStyle.xlContinuous

'.Rows(CStr(startRow) & ":" & CStr(totrow + 2)).Ungroup
For i = 1 To zoneCount
    bbdrow = startRow + 2 + (i - 1) * 15
    .Rows(CStr(bbdrow) & ":" & CStr(bbdrow + 3)).Group
    .Rows(bbdrow).ShowDetail = False
    .Rows(CStr(bbdrow + 5) & ":" & CStr(bbdrow + 8)).Group
    .Rows(bbdrow + 5).ShowDetail = False
    .Rows(CStr(bbdrow + 10) & ":" & CStr(bbdrow + 13)).Group
    .Rows(bbdrow + 10).ShowDetail = False
Next
.Columns("B").ColumnWidth = 25
End With

End Sub


Function getQuarterData(yearstr As String) As Variant
Dim Qdata(4) As Integer
Set sht = ActiveSheet
Dim monthCol As Integer
For i = 1 To sht.UsedRange.Columns.count
    If sht.Cells(1, i) = "Month" Then
        monthCol = i
    End If
Next

For i = 2 To sht.UsedRange.Rows.count
    If Right(sht.Cells(i, monthCol), 2) = Right(yearstr, 2) Then
        Select Case Left(sht.Cells(i, monthCol), 3)
            Case "Jan": Qdata(0) = Qdata(0) + 1
            Case "Feb": Qdata(0) = Qdata(0) + 1
            Case "Mar": Qdata(0) = Qdata(0) + 1
            Case "Apr": Qdata(1) = Qdata(1) + 1
            Case "May": Qdata(1) = Qdata(1) + 1
            Case "Jun": Qdata(1) = Qdata(1) + 1
            Case "Jul": Qdata(2) = Qdata(2) + 1
            Case "Aug": Qdata(2) = Qdata(2) + 1
            Case "Sep": Qdata(2) = Qdata(2) + 1
            Case "Oct": Qdata(3) = Qdata(3) + 1
            Case "Nov": Qdata(3) = Qdata(3) + 1
            Case "Dec": Qdata(3) = Qdata(3) + 1
        End Select
    End If
Next

getQuarterData = Qdata

End Function

Private Sub FTEPage()

Dim trackersht As Worksheet
Set trackersht = P_CUBK.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
Dim datasht As Worksheet
Set datasht = P_CUBK.Worksheets("Active List Conso")

Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim rangeStr As String
Dim StartPvt As String
Dim NumofRows As String


On Error Resume Next

NumofRows = datasht.UsedRange.Rows.count

rangeStr = "A1:" & getColString(datasht.UsedRange.Columns.count) & NumofRows

SrcData = datasht.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)

trackersht.Name = "FTE"
StartPvt = "'" & trackersht.Name & "'" & "!" & trackersht.Range("A1").Address(ReferenceStyle:=xlR1C1)

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
   
  
  pvt.AddDataField pvt.PivotFields("Descr"), "Count of Descr", xlCount
  pvt.AddFields RowFields:=Array("BBD", "Description"), ColumnFields:="Category"
  pvt.PivotFields("Month").Orientation = xlPageField
  For Each itm In pvt.PivotFields("BBD").PivotItems
    If ThisWorkbook.Worksheets("Zone Map").Range("B:B").Find(itm.Name) Is Nothing Or itm.Name = "Specialist" Then
        itm.Visible = False
    End If
  Next
  For Each itm In pvt.PivotFields("Category").PivotItems
    If itm.Name = "ERROR" Then
        itm.Visible = False
    End If
  Next
  pvt.PivotFields("Month").CurrentPage = monthstr
End Sub

Sub PBRMTable()

Set bk = P_CUBK
Set sht = bk.Worksheets.Add
sht.Name = "PBRM Table"


Dim pvtCache As PivotCache
Dim pvt As PivotTable

Dim SrcData, rangeStr, StartPvt, NumofRows As String

On Error Resume Next

Set datasht = bk.Worksheets("RM Resignation")
NumofRows = datasht.UsedRange.Rows.count
rangeStr = "A:Q"
SrcData = datasht.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A4").Address(ReferenceStyle:=xlR1C1)
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
  pvt.AddDataField pvt.PivotFields("RM"), "Count of RM", xlCount
  pvt.AddFields RowFields:="Count of Within Industry", ColumnFields:="Quarter"
  For Each itm In pvt.PivotFields("Quarter").PivotItems
    If itm.Name = "(blank)" Then
        itm.Visible = False
    End If
  Next
  

StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A12").Address(ReferenceStyle:=xlR1C1)
'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
'Create Pivot table from Pivot Cache
  Set pvt3 = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable3")
  pvt3.ManualUpdate = False
  pvt3.InGridDropZones = True
  pvt3.RowAxisLayout xlTabularRow
  pvt3.ColumnGrand = True
  pvt3.RowGrand = True
  pvt3.EnableDataValueEditing = False
  pvt3.EnableDrilldown = True
  pvt3.AddDataField pvt.PivotFields("RM"), "Count of RM", xlCount
  pvt3.AddFields RowFields:="Reason Specific", ColumnFields:="Quarter"
  For Each itm In pvt3.PivotFields("Quarter").PivotItems
    If itm.Name = "(blank)" Then
        itm.Visible = False
    End If
  Next
  
Set datasht1 = bk.Worksheets("PB Resignation")
NumofRows = datasht1.UsedRange.Rows.count
rangeStr = "A:P"
SrcData = datasht1.Name & "!" & Range(rangeStr).Address(ReferenceStyle:=xlR1C1)
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("I4").Address(ReferenceStyle:=xlR1C1)
'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
'Create Pivot table from Pivot Cache
  Set pvt1 = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")
  pvt1.ManualUpdate = False
  pvt1.InGridDropZones = True
  pvt1.RowAxisLayout xlTabularRow
  pvt1.ColumnGrand = True
  pvt1.RowGrand = True
  pvt1.EnableDataValueEditing = False
  pvt1.EnableDrilldown = True
  pvt1.AddDataField pvt1.PivotFields("PB"), "Count of PB", xlCount
  pvt1.AddFields RowFields:="Count of Within Industry", ColumnFields:="Quarter"
  For Each itm In pvt1.PivotFields("Quarter").PivotItems
    If itm.Name = "(blank)" Then
        itm.Visible = False
    End If
  Next
  
  StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("I12").Address(ReferenceStyle:=xlR1C1)
'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
'Create Pivot table from Pivot Cache
  Set pvt4 = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable4")
  pvt4.ManualUpdate = False
  pvt4.InGridDropZones = True
  pvt4.RowAxisLayout xlTabularRow
  pvt4.ColumnGrand = True
  pvt4.RowGrand = True
  pvt4.EnableDataValueEditing = False
  pvt4.EnableDrilldown = True
  pvt4.AddDataField pvt.PivotFields("RM"), "Count of RM", xlCount
  pvt4.AddFields RowFields:="Reason Specific", ColumnFields:="Quarter"
  For Each itm In pvt4.PivotFields("Quarter").PivotItems
    If itm.Name = "(blank)" Then
        itm.Visible = False
    End If
  Next
  
'Formatting

sht.Cells(1, 1) = "20" & Right(monthstr, 2) & " - RM"
sht.Cells(1, 9) = "20" & Right(monthstr, 2) & " - PB"
sht.Cells(2, 1) = "Average Vintage: " + CStr(Format(Application.WorksheetFunction.Average(datasht.Columns("I")), "0.0"))
sht.Cells(2, 9) = "Average Vintage: " + CStr(Format(Application.WorksheetFunction.Average(datasht1.Columns("I")), "0.0"))
With sht.Range("A1,I1")
.Interior.Color = vbYellow
.Font.Bold = True
End With
Dim col As Integer
col = pvt.TableRange1.Columns(pvt.TableRange1.Columns.count).Column
sht.Columns("B:" & getColString(col - 1)).ColumnWidth = 3
col = pvt1.TableRange1.Columns(pvt1.TableRange1.Columns.count).Column
sht.Columns("J:" & getColString(col - 1)).ColumnWidth = 3

End Sub


Sub createSummarySheet()

Dim curMonth As Integer
curMonth = Month("1 " & Left(monthstr, 3))

Dim zoneCount As Integer
zoneCount = 0
Set bk = P_CUBK
Set sht = bk.Worksheets.Add
sht.Name = "Summary"

'Table 1
sht.Cells(2, 1) = "Table 1: Breakdown by Zone"
With bk.Worksheets("Attrition")
.Activate
Set pvt = .PivotTables(1)
startRow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.count).Row + 3
startRow = .Cells(startRow + 2, 2).End(xlDown).Row + 4
.Rows(CStr(startRow) & ":" & CStr(.UsedRange.Rows.count)).Copy
End With
With sht
.Activate
.Cells(4, 1).PasteSpecial xlPasteAll
Selection.Ungroup
Selection.UnMerge
Selection.EntireRow.Hidden = False
For i = 5 To .UsedRange.Rows.count
    If .Cells(i, 1) <> "" And .Cells(i, 1) <> "Total" Then
        zoneCount = zoneCount + 1
    End If
Next

Dim dlrange As Range
For i = 1 To zoneCount
    If dlrange Is Nothing Then
        Set dlrange = .Rows(CStr((i - 1) * 15 + 6) & ":" & CStr((i - 1) * 15 + 9))
    Else
        Set dlrange = Union(dlrange, .Rows(CStr((i - 1) * 15 + 6) & ":" & CStr((i - 1) * 15 + 9)))
    End If
    Set dlrange = Union(dlrange, .Rows(CStr((i - 1) * 15 + 11) & ":" & CStr((i - 1) * 15 + 19)))
    .Cells((i - 1) * 15 + 10, 1) = .Cells((i - 1) * 15 + 6, 1)
Next

dlrange.Delete
.Rows(.UsedRange.Rows.count + 1).Insert
.Rows(.UsedRange.Rows.count + 1).Insert
.Cells(8 + 2 * zoneCount, 2) = "Plan FTE"
.Cells(9 + 2 * zoneCount, 2) = "YTD Attrition Rate"
For i = 3 To .UsedRange.Columns.count
    .Cells(9 + 2 * zoneCount, i) = .Cells(10 + 2 * zoneCount, i) / 12 * curMonth
    .Cells(9 + 2 * zoneCount, i).NumberFormat = "0%"
    'Fill in PLAN FTE
Next

.Columns("B").Insert
For i = 1 To zoneCount
    .Range(.Cells(6 + (i - 1) * 2, 1), .Cells(7 + (i - 1) * 2, 1)).Merge
Next
.Range(.Cells(6 + zoneCount * 2, 1), .Cells(10 + zoneCount * 2, 1)).Merge
For i = 5 To 10 + zoneCount * 2
    .Range(.Cells(i, 2), .Cells(i, 3)).Merge
Next
.UsedRange.Columns.ColumnWidth = 11
End With

'add in branch table
'With bk.Worksheets("Attrition")
'Set pvt = .PivotTables(1)
'startRow = pvt.TableRange1.Rows(pvt.TableRange1.Rows.count).Row + 3
'.Cells(startRow, 1).CurrentRegion.Copy
'cct = sht.UsedRange.Rows.count + 4
'sht.Cells(sht.UsedRange.Rows.count + 4, 1).Select
'Selection.PasteSpecial xlPasteAll
'sht.Cells(cct, 1) = "Table 2: Branch Annualized Attrition Rate"
'End With
    

'Table 2
With sht
startRow = .UsedRange.Rows.count + 4
.Cells(startRow, 1) = "Table 2: Sales FTE"
.Cells(startRow, 1).Font.Bold = True
.Cells(startRow + 1, 1) = "'" & Left(monthstr, 3) & Right(monthstr, 2)
.Cells(startRow + 2, 1) = "Ranked FTE"
.Cells(startRow + 3, 1) = "Training + Pending RIN"
.Cells(startRow + 4, 1) = "Unranked"
.Cells(startRow + 5, 1) = "Transferred Out"
.Cells(startRow + 6, 1) = "Resigned"
.Cells(startRow + 7, 1) = "Total (excl. Resigned)"

.Cells(startRow + 1, 3) = "CPC RM"
.Cells(startRow + 1, 4) = "RM"
.Cells(startRow + 1, 5) = "PB"
.Cells(startRow + 1, 6) = "AB"
.Cells(startRow + 1, 7) = "DB Specialist"
.Cells(startRow + 1, 8) = "Total"

Set pvt = ActiveWorkbook.Worksheets("Tracker(S)").PivotTables(1)
On Error Resume Next
For i = 1 To 5
   ' MsgBox .Cells(startRow + 1, i + 2).Value
   ' MsgBox pvt.GetPivotData("Status", "'Tracker(S)'!$A$1", "Type", .Cells(startRow + 1, i + 2).Value)
    'MsgBox Left(monthstr, 3) & Right(monthstr, 2)
    
    .Cells(startRow + 2, i + 2) = pvt.GetPivotData("Status", "Type", .Cells(startRow + 1, i + 2).Value, "Status", "Ranked", "Month", Left(monthstr, 3) & Right(monthstr, 2))
    .Cells(startRow + 3, i + 2) = pvt.GetPivotData("Status", "Type", .Cells(startRow + 1, i + 2).Value, "Status", "Training", "Month", Left(monthstr, 3) & Right(monthstr, 2))
    .Cells(startRow + 4, i + 2) = pvt.GetPivotData("Status", "Type", .Cells(startRow + 1, i + 2).Value, "Status", "Unranked", "Month", Left(monthstr, 3) & Right(monthstr, 2))
    .Cells(startRow + 5, i + 2) = pvt.GetPivotData("Status", "Type", .Cells(startRow + 1, i + 2).Value, "Status", "Transfer", "Month", Left(monthstr, 3) & Right(monthstr, 2))
    .Cells(startRow + 6, i + 2) = pvt.GetPivotData("Status", "Type", .Cells(startRow + 1, i + 2).Value, "Status", "Resigned", "Month", Left(monthstr, 3) & Right(monthstr, 2))
    .Cells(startRow + 7, i + 2) = Application.WorksheetFunction.sum(.Range(.Cells(startRow + 2, i + 2), .Cells(startRow + 5, i + 2)))
Next
For i = 2 To 7
    .Cells(startRow + i, 8) = Application.WorksheetFunction.sum(.Range(.Cells(startRow + i, 3), .Cells(startRow + i, 7)))
    .Range(.Cells(startRow + i, 1), .Cells(startRow + i, 2)).Merge
Next
.Range(.Cells(startRow + 1, 1), .Cells(startRow + 1, 2)).Merge
.Range(.Cells(startRow, 1), .Cells(startRow, 8)).Merge
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlTop).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlBottom).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlRight).LineStyle = XlLineStyle.xlContinuous
.Range(.Cells(startRow, 1), .Cells(startRow + 7, 8)).Borders(xlLeft).LineStyle = XlLineStyle.xlContinuous

End With


'table 4
startRow = sht.UsedRange.Rows.count + 3
With sht
.Cells(startRow, 1) = "Table 3: S&D Attrition Numbers / New Hires"
.Cells(startRow, 1).Font.Bold = True
.Cells(startRow + 1, 2) = "Resignation from S&D"
.Cells(startRow + 1, 5) = "New Hires"
.Range(.Cells(startRow + 1, 2), .Cells(startRow + 1, 4)).Merge
.Range(.Cells(startRow + 1, 5), .Cells(startRow + 1, 6)).Merge
.Cells(startRow + 2, 1) = "Month"
.Cells(startRow + 2, 2) = "Total"
.Cells(startRow + 2, 3) = "PB"
.Cells(startRow + 2, 4) = "RM"
.Cells(startRow + 2, 5) = "PB"
.Cells(startRow + 2, 6) = "RM"
.Cells(startRow + 3, 1) = "Jan"
.Cells(startRow + 4, 1) = "Feb"
.Cells(startRow + 5, 1) = "Mar"
.Cells(startRow + 6, 1) = "Apr"
.Cells(startRow + 7, 1) = "May"
.Cells(startRow + 8, 1) = "Jun"
.Cells(startRow + 9, 1) = "Jul"
.Cells(startRow + 10, 1) = "Aug"
.Cells(startRow + 11, 1) = "Sep"
.Cells(startRow + 12, 1) = "Oct"
.Cells(startRow + 13, 1) = "Nov"
.Cells(startRow + 14, 1) = "Dec"
.Cells(startRow + 15, 1) = "Total"

Set asht = ActiveWorkbook.Worksheets("Leavers")
For i = 2 To asht.Range("A1").End(xlDown).Row
    If Not IsError(asht.Range("J" & CStr(i))) Then
        Select Case Left(asht.Range("K" & CStr(i)), 3)
        Case "Jan": .Cells(startRow + 3, 2) = .Cells(startRow + 3, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 3, 3) = .Cells(startRow + 3, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 3, 4) = .Cells(startRow + 3, 4) + 1
                    End If
        Case "Feb": .Cells(startRow + 4, 2) = .Cells(startRow + 4, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 4, 3) = .Cells(startRow + 4, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 4, 4) = .Cells(startRow + 4, 4) + 1
                    End If
        Case "Mar": .Cells(startRow + 5, 2) = .Cells(startRow + 5, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 5, 3) = .Cells(startRow + 5, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 5, 4) = .Cells(startRow + 5, 4) + 1
                    End If
        Case "Apr": .Cells(startRow + 6, 2) = .Cells(startRow + 6, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 6, 3) = .Cells(startRow + 6, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 6, 4) = .Cells(startRow + 6, 4) + 1
                    End If
        Case "May": .Cells(startRow + 7, 2) = .Cells(startRow + 7, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 7, 3) = .Cells(startRow + 7, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 7, 4) = .Cells(startRow + 7, 4) + 1
                    End If
        Case "Jun": .Cells(startRow + 8, 2) = .Cells(startRow + 8, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 8, 3) = .Cells(startRow + 8, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 8, 4) = .Cells(startRow + 8, 4) + 1
                    End If
        Case "Jul": .Cells(startRow + 9, 2) = .Cells(startRow + 9, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 9, 3) = .Cells(startRow + 9, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 9, 4) = .Cells(startRow + 9, 4) + 1
                    End If
        Case "Aug": .Cells(startRow + 10, 2) = .Cells(startRow + 10, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 10, 3) = .Cells(startRow + 10, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 10, 4) = .Cells(startRow + 10, 4) + 1
                    End If
        Case "Sep": .Cells(startRow + 11, 2) = .Cells(startRow + 11, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 11, 3) = .Cells(startRow + 11, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 11, 4) = .Cells(startRow + 11, 4) + 1
                    End If
        Case "Oct": .Cells(startRow + 12, 2) = .Cells(startRow + 12, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 12, 3) = .Cells(startRow + 12, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 12, 4) = .Cells(startRow + 12, 4) + 1
                    End If
        Case "Nov": .Cells(startRow + 13, 2) = .Cells(startRow + 13, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 13, 3) = .Cells(startRow + 13, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 13, 4) = .Cells(startRow + 13, 4) + 1
                    End If
        Case "Dec": .Cells(startRow + 14, 2) = .Cells(startRow + 14, 2) + 1
                    If asht.Range("F" & CStr(i)) = "Personal Banker" Then
                        .Cells(startRow + 14, 3) = .Cells(startRow + 14, 3) + 1
                    ElseIf asht.Range("F" & CStr(i)) = "Relationship Manager" Or asht.Range("F" & CStr(i)) = "CPC Relationship Manager" Then
                        .Cells(startRow + 14, 4) = .Cells(startRow + 14, 4) + 1
                    End If
        End Select
    End If
Next


ThisWorkbook.Worksheets("New Hire").Range("A3:B14").Copy
sht.Cells(startRow + 3, 5).PasteSpecial xlPasteValues
sht.Range(.Cells(startRow, 1), .Cells(startRow, 6)).Merge

sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlTop).LineStyle = XlLineStyle.xlContinuous
sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlBottom).LineStyle = XlLineStyle.xlContinuous
sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlRight).LineStyle = XlLineStyle.xlContinuous
sht.Cells(startRow + 3, 5).CurrentRegion.Borders(xlLeft).LineStyle = XlLineStyle.xlContinuous
For i = 2 To 6
    .Cells(startRow + 15, i) = Application.WorksheetFunction.sum(.Range(.Cells(startRow + 3, i), .Cells(startRow + 14, i)))
Next
End With

End Sub

Function getQuarter() As String
Select Case Left(monthstr, 3)
Case "Jan": getQuarter = "Q1"
Case "Feb": getQuarter = "Q1"
Case "Mar": getQuarter = "Q1"
Case "Apr": getQuarter = "Q2"
Case "May": getQuarter = "Q2"
Case "Jun": getQuarter = "Q2"
Case "Jul": getQuarter = "Q3"
Case "Aug": getQuarter = "Q3"
Case "Sep": getQuarter = "Q3"
Case "Oct": getQuarter = "Q4"
Case "Nov": getQuarter = "Q4"
Case "Dec": getQuarter = "Q4"
End Select

End Function

Function getMonthNumber(str As String) As Integer
Dim zName As String
Dim zNum As Integer

zName = UCase(str)

Select Case zName

Case UCase("Jan")
zNum = 1
Case UCase("Feb")
zNum = 2
Case UCase("Mar")
zNum = 3
Case UCase("Apr")
zNum = 4
Case UCase("May")
zNum = 5
Case UCase("Jun")
zNum = 6
Case UCase("Jul")
zNum = 7
Case UCase("Aug")
zNum = 8
Case UCase("Sep")
zNum = 9
Case UCase("Oct")
zNum = 10
Case UCase("Nov")
zNum = 11
Case UCase("Dec")
zNum = 12

End Select

getMonthNumber = zNum
End Function


Sub fillInExtraStatus()

'Set P_PRBK = ActiveWorkbook 'test
'monthstr = "Aug'15" 'test

Set bssht = P_PRBK.Worksheets("Banker Status")
   
    Workbooks.Open RBWMPath, False, , , "coffee"
    Set rbbk = ActiveWorkbook
    Set rbsht = rbbk.Worksheets("Sales Data")
    
    
    For k = 4 To 6
    For j = 3 To 6
    For i = 1 To rbsht.Cells(k, j)
        r = bssht.UsedRange.Rows.count + 1
        bssht.Cells(r, 3) = rbsht.Cells(1, j)
        bssht.Cells(r, 4) = rbsht.Cells(k, 2)
        bssht.Cells(r, 6) = monthstr
        bssht.Cells(r, 7) = getQuarter()
    Next
    Next
    Next
    
        
    Application.DisplayAlerts = False
    rbbk.Close
    Application.DisplayAlerts = True
    
    
End Sub


