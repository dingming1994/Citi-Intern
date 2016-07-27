Public saveToPath As String

Sub ConcFiles_CB_REV1()

Application.FileDialog(msoFileDialogOpen).Title = "Select the contracted revenue files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show

Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
    'For i = 1 To fileCount
    '    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    '    Set fileArr(i) = ActiveWorkbook
    'Next
End If

Dim isFound As Integer
isFound = 0
For i = 1 To fileCount
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CB_Contracted_Product_Revenue_?" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
    If isFound = 0 Then
        MsgBox "Please select the correct files"
     
        Exit Sub
    End If
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 31, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        
    Next
    
 

    correctsht.Activate
    correctsht.Rows("2:" & CStr(correctsht.UsedRange.Rows.Count)).Select
   
    Selection.Copy
    If i <> 1 Then
        fileArr(1).Activate
        For Each sht In fileArr(1).Worksheets
        If sht.Name Like "CB_Contracted_Product_Revenue_?" Then
            Set pastesht = sht
            isFound = 1
        End If
        Next
        
        pastesht.Activate
        pastesht.Cells(pastesht.UsedRange.Rows.Count + 1, 1).Select
        Selection.PasteSpecial xlPasteValues
        
        Application.DisplayAlerts = False
        fileArr(i).Close
        Application.DisplayAlerts = True
        
    End If
        
    
Next

End Sub

Sub ConcFiles_CB_REV()

Application.FileDialog(msoFileDialogOpen).Title = "Select the CB contracted revenue files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show


Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
Exit Sub
End If


Set newbk = Workbooks.Add
Application.DisplayAlerts = False
newbk.Worksheets(2).Delete
newbk.Worksheets(2).Delete
Set appsht = newbk.Worksheets(1)
appsht.Name = "RAW_DATA"
Application.DisplayAlerts = True

Dim isFound As Integer
isFound = 0
For i = 1 To fileCount
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CB_Contracted_Product_Revenue_?" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
    If isFound = 0 Then
        MsgBox "Please select the correct files"
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArr(10) As Integer
    Dim ColTitle As Variant
    ColTitle = Array("RM", "GEID", "Mgr", "Status", "INS_DATA", "TOTAL_Effort_REVENUE", "NRFF", "TRAILER", "FINAL_REVENUE", "MONTH")
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 31, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        If correctsht.Cells(j, 1) <> "" And correctsht.Cells(j, 2) <> "" Then
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        End If
    Next

    correctsht.Activate
    
    For j = 1 To correctsht.UsedRange.Columns.Count
        Dim str As String
        str = correctsht.Cells(1, j).Value
        For k = 0 To 9
            If LCase(str) = LCase(ColTitle(k)) Then
                ColIdxArr(k + 1) = j
            End If
        Next
            
    Next
    
    Set Y = correctsht.UsedRange.Columns(ColIdxArr(1))
    
    For k = 2 To 10
        Set Y = Union(Y, correctsht.UsedRange.Columns(ColIdxArr(k)))
    Next
        
    Y.Select
    Selection.Copy
    appsht.Activate

    appsht.Cells(appsht.UsedRange.Rows.Count + 1, 1).Select
    Selection.PasteSpecial xlPasteValues
        
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next

    appsht.Activate
appsht.Rows(1).Delete
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 1).Value = "RM" And appsht.Cells(i, 2).Value = "GEID" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
    If appsht.Cells(i, 1).Value = "" And appsht.Cells(i, 2).Value = "" Then
        appsht.Rows(i).Delete
    End If
Next

appsht.Activate
createPivotTable_ALL_CB
End Sub

Sub ConcFiles_CGG3_REV()

Application.FileDialog(msoFileDialogOpen).Title = "Select the CGG3 contracted revenue files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show


Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
Exit Sub
End If

Set newbk = Workbooks.Add

Application.DisplayAlerts = False
newbk.Worksheets(2).Delete
newbk.Worksheets(2).Delete
Set appsht = newbk.Worksheets(1)
appsht.Name = "RAW_DATA"
Application.DisplayAlerts = True

Dim isFound As Integer
isFound = 0
For i = 1 To fileCount
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CG_Contracted_Product_Revenue_?" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
    If isFound = 0 Then
        MsgBox "Please select the correct files"
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArr(11) As Integer
    Dim ColTitle As Variant
    ColTitle = Array("RM", "GEID", "G3_IND", "Mgr", "Status", "INS_DATA", "TOTAL_Effort_REVENUE", "NRFF", "TRAILER", "TOTAL_REVENUE", "MONTH")
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 33, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        If correctsht.Cells(j, 1) <> "" And correctsht.Cells(j, 2) <> "" Then
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        End If
    Next

    correctsht.Activate
    
    For j = 1 To correctsht.UsedRange.Columns.Count
        Dim str As String
        str = correctsht.Cells(1, j).Value
        For k = 0 To 10
            If LCase(str) = LCase(ColTitle(k)) Then
                ColIdxArr(k + 1) = j
            End If
        Next
            
    Next
    
    Set Y = correctsht.UsedRange.Columns(ColIdxArr(1))
    
    For k = 2 To 11
        Set Y = Union(Y, correctsht.UsedRange.Columns(ColIdxArr(k)))
    Next
        
    Y.Select
    Selection.Copy
    appsht.Activate

    appsht.Cells(appsht.UsedRange.Rows.Count + 1, 1).Select
    Selection.PasteSpecial xlPasteValues
        
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next

    appsht.Activate
appsht.Rows(1).Delete
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 1).Value = "RM" And appsht.Cells(i, 2).Value = "GEID" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
    If appsht.Cells(i, 1).Value = "" And appsht.Cells(i, 2).Value = "" Then
        appsht.Rows(i).Delete
    End If
Next

appsht.Activate
createPivotTable_ALL_CGG3
End Sub

Sub ConcFiles_CGCB_CREDIT()

Application.FileDialog(msoFileDialogOpen).Title = "Select the CGGB contracted credit files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show


Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
    Exit Sub
End If

Set newbk = Workbooks.Add

Application.DisplayAlerts = False
newbk.Worksheets(3).Delete
Set appshtCG = newbk.Worksheets(1)
appshtCG.Name = "CG_RAW_DATA"
Set appshtCB = newbk.Worksheets(2)
appshtCB.Name = "CB_RAW_DATA"
Application.DisplayAlerts = True

Dim isFound As Integer
For i = 1 To fileCount
isFound = 0
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CG_CREDITS_Summary" Then
            Set correctshtCG = sht
            isFound = isFound + 1
        End If
        If sht.Name Like "CB_CREDITS_Summary" Then
            Set correctshtCB = sht
            isFound = isFound + 1
        End If
    Next
    
    If isFound <> 2 Then
        MsgBox "Please select the correct files"
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArrCB(9) As Integer
    Dim ColIdxArrCG(11) As Integer
    Dim ColTitleCB As Variant
    Dim ColTitleCG As Variant
    ColTitleCB = Array("RM", "GEID", "Mgr", "Status", "Monthly_Credit_Threshold", "Monthly_Credit_Threshold_Prorate", "Total_Sales_Credits", "ABU_Ranking", "MONTH")
    ColTitleCG = Array("RM", "GEID", "Mgr", "Status", "Monthly_Credit_Threshold", "Monthly_Credit_Threshold_Prorate", "Total_Sales_Credits", "ABU_Ranking", "Rank", "G3_IND", "MONTH")
    
    With correctshtCB
        .Activate
        fileName = fileArr(i).Name
        colCount = .UsedRange.Columns.Count + 1
        monthStr = Mid(fileName, 49, 5)
        .Cells(1, colCount).Value = "MONTH"
        For j = 2 To .UsedRange.Rows.Count
            If .Cells(j, 1) <> "" And .Cells(j, 2) <> "" Then
            .Cells(j, colCount).Value = getMonthStr(monthStr)
            End If
        Next
    
        .Activate
        
        For j = 1 To .UsedRange.Columns.Count
            Dim str As String
            str = .Cells(1, j).Value
            For k = 0 To 8
                If LCase(str) = LCase(ColTitleCB(k)) Then
                    ColIdxArrCB(k + 1) = j
                End If
            Next
                
        Next
        
        Set Y = .UsedRange.Columns(ColIdxArrCB(1))
        
        For k = 2 To 9
            Set Y = Union(Y, .UsedRange.Columns(ColIdxArrCB(k)))
        Next
            
        Y.Select
        Selection.Copy
        appshtCB.Activate
        If i = 1 Then
            appshtCB.Cells(1, 1).Select
        Else
            appshtCB.Cells(appshtCB.UsedRange.Rows.Count + 1, 1).Select
        End If
        Selection.PasteSpecial xlPasteValues
    End With
    
    
    With correctshtCG
        .Activate
        fileName = fileArr(i).Name
        colCount = .UsedRange.Columns.Count + 1
        monthStr = Mid(fileName, 49, 5)
        .Cells(1, colCount).Value = "MONTH"
        For j = 2 To .UsedRange.Rows.Count
            If .Cells(j, 1) <> "" And .Cells(j, 2) <> "" Then
            .Cells(j, colCount).Value = getMonthStr(monthStr)
            End If
        Next
    
        .Activate
        
        For j = 1 To .UsedRange.Columns.Count
            str = .Cells(1, j).Value
            For k = 0 To 10
                If LCase(str) = LCase(ColTitleCG(k)) Then
                    ColIdxArrCG(k + 1) = j
                End If
            Next
                
        Next
        
        Set Y = .UsedRange.Columns(ColIdxArrCG(1))
        
        For k = 2 To 11
            Set Y = Union(Y, .UsedRange.Columns(ColIdxArrCG(k)))
        Next
            
        Y.Select
        Selection.Copy
        appshtCG.Activate
        If i = 1 Then
            appshtCG.Cells(1, 1).Select
        Else
            appshtCG.Cells(appshtCG.UsedRange.Rows.Count + 1, 1).Select
        End If
        Selection.PasteSpecial xlPasteValues
    End With
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next
    With appshtCB
        .Activate
    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 1).Value = "RM" And .Cells(i, 2).Value = "GEID" Then
            .Rows(i).Delete
            i = i - 1
        End If
        If .Cells(i, 1).Value = "" And .Cells(i, 2).Value = "" Then
            .Rows(i).Delete
        End If
    Next
    End With
    
    With appshtCG
        .Activate

    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 1).Value = "RM" And .Cells(i, 2).Value = "GEID" Then
            .Rows(i).Delete
            i = i - 1
        End If
        If .Cells(i, 1).Value = "" And .Cells(i, 2).Value = "" Then
            .Rows(i).Delete
        End If
    Next
    End With

appshtCG.Activate

createPivotTable_ALL_CGCB_Credit
End Sub

Function ConcFiles_PB_UPGRADE() As Integer

ConcFiles_PB_UPGRADE = 1
Application.FileDialog(msoFileDialogOpen).Title = "Select the contracted credit files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show

Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
    ConcFiles_PB_UPGRADE = 0
    Exit Function
End If

Set newbk = Workbooks.Add
Application.DisplayAlerts = False
newbk.Worksheets(2).Delete
newbk.Worksheets(2).Delete
Set appsht = newbk.Worksheets(1)
Application.DisplayAlerts = True


Dim isFound As Integer
isFound = 0
For i = 1 To fileCount
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CB_CREDITS_Summary" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
     If isFound = 0 Then
        MsgBox "Please select the correct files"
        ConcFiles_PB_UPGRADE = 0
        Application.DisplayAlerts = False
        fileArr(i).Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArr(4) As Integer
    
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 49, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        
    Next
    
   
    correctsht.Activate
    
    For j = 1 To correctsht.UsedRange.Columns.Count
        Dim str As String
        str = correctsht.Cells(i, j)
        If correctsht.Cells(1, j).Value = "RM" Then
            ColIdxArr(1) = j
        End If
        If correctsht.Cells(1, j).Value = "Mgr" Then
            ColIdxArr(2) = j
        End If
        If correctsht.Cells(1, j).Value = "ABU_Ranking" Then
            ColIdxArr(3) = j
        End If
        If correctsht.Cells(1, j).Value = "MONTH" Then
            ColIdxArr(4) = j
        End If
    Next
            
    
    Union(correctsht.UsedRange.Columns(ColIdxArr(1)), correctsht.UsedRange.Columns(ColIdxArr(2)), correctsht.UsedRange.Columns(ColIdxArr(3)), correctsht.UsedRange.Columns(ColIdxArr(4))).Select
   
    Selection.Copy
    appsht.Activate

    appsht.Cells(appsht.UsedRange.Rows.Count + 1, 1).Select
    Selection.PasteSpecial xlPasteValues
        
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next

appsht.Activate
appsht.Rows(1).Delete
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 1).Value = "RM" And appsht.Cells(i, 2).Value = "Mgr" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
    If appsht.Cells(i, 1).Value = "" And appsht.Cells(i, 2).Value = "" Then
        appsht.Rows(i).Delete
    End If
Next

' ABU_Ranking: A = 1 B = 2 C = 3
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 3).Value = "A" Then
        appsht.Cells(i, 3).Value = 100
    End If
    If appsht.Cells(i, 3).Value = "B" Then
        appsht.Cells(i, 3).Value = 1
    End If
    If appsht.Cells(i, 3).Value = "U" Then
        appsht.Cells(i, 3).Value = 0.01
    End If
Next

createPivotTable_ABU ("PB")
newbk.Activate
InsertCodePT ActiveWorkbook, "PB"
newbk.Worksheets("PB ABU Pivot").PivotTables(1).RefreshTable

Application.DisplayAlerts = False
appsht.Delete
Application.DisplayAlerts = True
newbk.Activate

End Function

Function ConcFiles_RM_UPGRADE() As Integer

ConcFiles_RM_UPGRADE = 1
Application.FileDialog(msoFileDialogOpen).Title = "Select the contracted credit files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show

Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
    ConcFiles_RM_UPGRADE = 0
    Exit Function
End If

Set newbk = Workbooks.Add
Application.DisplayAlerts = False
newbk.Worksheets(2).Delete
newbk.Worksheets(2).Delete
Set appsht = newbk.Worksheets(1)
Application.DisplayAlerts = True


Dim isFound As Integer
isFound = 0
For i = 1 To fileCount
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CG_CREDITS_Summary" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
    If isFound = 0 Then
        MsgBox "Please select the correct files"
        ConcFiles_RM_UPGRADE = 0
        Application.DisplayAlerts = False
        fileArr(i).Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArr(4) As Integer
    
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 49, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        
    Next
    
    correctsht.Activate
    
    For j = 1 To correctsht.UsedRange.Columns.Count
        Dim str As String
        str = correctsht.Cells(i, j)
        If correctsht.Cells(1, j).Value = "RM" Then
            ColIdxArr(1) = j
        End If
        If correctsht.Cells(1, j).Value = "Mgr" Then
            ColIdxArr(2) = j
        End If
        If correctsht.Cells(1, j).Value = "ABU_Ranking" Then
            ColIdxArr(3) = j
        End If
        If correctsht.Cells(1, j).Value = "MONTH" Then
            ColIdxArr(4) = j
        End If
    Next
            
    
    Union(correctsht.UsedRange.Columns(ColIdxArr(1)), correctsht.UsedRange.Columns(ColIdxArr(2)), correctsht.UsedRange.Columns(ColIdxArr(3)), correctsht.UsedRange.Columns(ColIdxArr(4))).Select
   
    Selection.Copy
    appsht.Activate

    appsht.Cells(appsht.UsedRange.Rows.Count + 1, 1).Select
    Selection.PasteSpecial xlPasteValues
        
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next

appsht.Activate
appsht.Rows(1).Delete
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 1).Value = "RM" And appsht.Cells(i, 2).Value = "Mgr" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
    If appsht.Cells(i, 1).Value = "" And appsht.Cells(i, 2).Value = "" Then
        appsht.Rows(i).Delete
    End If
    If appsht.Cells(i, 1).Value <> "" And appsht.Cells(i, 3).Value = "" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
Next

' ABU_Ranking: A = 1 B = 2 C = 3
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 3).Value = "A" Then
        appsht.Cells(i, 3).Value = 100
    End If
    If appsht.Cells(i, 3).Value = "B" Then
        appsht.Cells(i, 3).Value = 1
    End If
    If appsht.Cells(i, 3).Value = "U" Then
        appsht.Cells(i, 3).Value = 0.01
    End If
Next

createPivotTable_ABU ("RM")
newbk.Activate
InsertCodePT ActiveWorkbook, "RM"
newbk.Worksheets("RM ABU Pivot").PivotTables(1).RefreshTable

Application.DisplayAlerts = False
appsht.Delete
Application.DisplayAlerts = True
    
End Function
Function ConcFiles_CPC_UPGRADE() As Integer

ConcFiles_CPC_UPGRADE = 1
Application.FileDialog(msoFileDialogOpen).Title = "Select the contracted credit files"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\nascet03vdm03\dm_v002\BP-INCENTIVE\Mgt_Office\SIPS"
intchoice = Application.FileDialog(msoFileDialogOpen).Show

Dim fileArr() As Workbook
Dim fileCount As Integer
Dim correctsht As Worksheet
If intchoice <> 0 Then
    fileCount = Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
    ReDim fileArr(fileCount)
Else
    ConcFiles_CPC_UPGRADE = 0
    Exit Function
End If

Set newbk = Workbooks.Add
Application.DisplayAlerts = False
newbk.Worksheets(2).Delete
newbk.Worksheets(2).Delete
Set appsht = newbk.Worksheets(1)
Application.DisplayAlerts = True

Dim isFound As Integer

For i = 1 To fileCount
    isFound = 0
    Workbooks.Open (Application.FileDialog(msoFileDialogOpen).SelectedItems(i))
    Set fileArr(i) = ActiveWorkbook
    For Each sht In fileArr(i).Worksheets
        If sht.Name Like "CPC_ABU_Ranking" Then
            Set correctsht = sht
            isFound = 1
        End If
    Next
    
     If isFound = 0 Then
        MsgBox "Please select the correct files"
        ConcFiles_CPC_UPGRADE = 0
        Application.DisplayAlerts = False
        fileArr(i).Close
        newbk.Close
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    Dim fileName As String
    Dim monthStr As String
    Dim colCount As Integer
    Dim ColIdxArr(4) As Integer
    
    fileName = fileArr(i).Name
    colCount = correctsht.UsedRange.Columns.Count + 1
    monthStr = Mid(fileName, 49, 5)
    correctsht.Cells(1, colCount).Value = "MONTH"
    For j = 2 To correctsht.UsedRange.Rows.Count
        correctsht.Cells(j, colCount).Value = getMonthStr(monthStr)
        
    Next
    
   
    correctsht.Activate
    
    For j = 1 To correctsht.UsedRange.Columns.Count
        Dim str As String
        str = correctsht.Cells(i, j)
        If correctsht.Cells(1, j).Value = "RM" Then
            ColIdxArr(1) = j
        End If
        If correctsht.Cells(1, j).Value = "Mgr" Then
            ColIdxArr(2) = j
        End If
        If correctsht.Cells(1, j).Value = "ABU_Ranking" Then
            ColIdxArr(3) = j
        End If
        If correctsht.Cells(1, j).Value = "MONTH" Then
            ColIdxArr(4) = j
        End If
    Next
            
    
    Union(correctsht.UsedRange.Columns(ColIdxArr(1)), correctsht.UsedRange.Columns(ColIdxArr(2)), correctsht.UsedRange.Columns(ColIdxArr(3)), correctsht.UsedRange.Columns(ColIdxArr(4))).Select
   
    Selection.Copy
    appsht.Activate

    appsht.Cells(appsht.UsedRange.Rows.Count + 1, 1).Select
    Selection.PasteSpecial xlPasteValues
        
    Application.DisplayAlerts = False
    fileArr(i).Close
    Application.DisplayAlerts = True

Next

appsht.Activate
appsht.Rows(1).Delete
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 1).Value = "RM" And appsht.Cells(i, 2).Value = "Mgr" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
    If appsht.Cells(i, 1).Value = "" And appsht.Cells(i, 2).Value = "" Then
        appsht.Rows(i).Delete
    End If
    If appsht.Cells(i, 1).Value <> "" And appsht.Cells(i, 3).Value = "" Then
        appsht.Rows(i).Delete
        i = i - 1
    End If
Next

' ABU_Ranking: A = 1 B = 2 C = 3
For i = 2 To appsht.UsedRange.Rows.Count
    If appsht.Cells(i, 3).Value = "A" Then
        appsht.Cells(i, 3).Value = 100
    End If
    If appsht.Cells(i, 3).Value = "B" Then
        appsht.Cells(i, 3).Value = 1
    End If
    If appsht.Cells(i, 3).Value = "U" Then
        appsht.Cells(i, 3).Value = 0.01
    End If
Next

createPivotTable_ABU ("CPC")
newbk.Activate
InsertCodePT ActiveWorkbook, "CPC"
newbk.Worksheets("CPC ABU Pivot").PivotTables(1).RefreshTable
Application.DisplayAlerts = False
appsht.Delete
Application.DisplayAlerts = True
newbk.Activate
End Function

Sub createPivotTable_ALL_CB()
Set sht = ActiveSheet
createPivotTable_TER
sht.Activate
createPivotTable_TR (0)
End Sub

Sub createPivotTable_ALL_CGG3()
Set sht = ActiveSheet
createPivotTable_TER
sht.Activate
createPivotTable_TR (1)
End Sub

Sub createPivotTable_ALL_CGCB_Credit()
Set bk = ActiveWorkbook
bk.Worksheets("CB_RAW_DATA").Activate
createPivotTable_TSC
ActiveSheet.Name = "PB PIVOT (Total Sales Credit)"
bk.Worksheets("CG_RAW_DATA").Activate
createPivotTable_TSC
ActiveSheet.Name = "RM PIVOT (Total Sales Credit)"
End Sub

Sub createPivotTable_ABU(tp As String)

Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As String

NumOfRows = ActiveSheet.UsedRange.Rows.Count

RangeStr = "A1:D" & NumOfRows

SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

Set sht = Sheets.Add
sht.Name = tp & " ABU Pivot"
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

 'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.createPivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
   
  pvt.AddFields RowFields:="RM", ColumnFields:="Month"
  pvt.AddDataField pvt.PivotFields("ABU_Ranking"), "Sum of ABU_Ranking", xlSum
  pvt.PivotFields("Mgr").Orientation = xlPageField
  pvt.PivotFields("Sum of ABU_Ranking").NumberFormat = "[=100]""A"";[=1] ""B""; ""U"""
  pvt.GrandTotalName = "    A-B-U"
  pvt.ManualUpdate = False
  
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = False
   pvt.RowGrand = True
   pvt.EnableDataValueEditing = True
    pvt.EnableDrilldown = False
End Sub


Sub createPivotTable_TSC()

Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As Integer
Dim NumOfCols As Integer

NumOfRows = ActiveSheet.UsedRange.Rows.Count
NumOfCols = ActiveSheet.UsedRange.Columns.Count
RangeStr = "A1:" & ActiveSheet.Cells(NumOfRows, NumOfCols).Address
SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

Set sht = Sheets.Add
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

 'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.createPivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
   
  pvt.AddFields RowFields:="RM", ColumnFields:="Month"
  pvt.AddDataField pvt.PivotFields("Total_Sales_Credits"), "Sum of Total_Sales_Credits", xlSum
  pvt.PivotFields("Mgr").Orientation = xlPageField
  pvt.ManualUpdate = False
  
  sht.Range(ActiveSheet.Cells(5, 2), ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).NumberFormat = "#,###.##"
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom

End Sub

Sub createPivotTable_TER()

Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As Integer
Dim NumOfCols As Integer

NumOfRows = ActiveSheet.UsedRange.Rows.Count
NumOfCols = ActiveSheet.UsedRange.Columns.Count
RangeStr = "A1:" & ActiveSheet.Cells(NumOfRows, NumOfCols).Address
SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

Set sht = Sheets.Add
sht.Name = "PIVOT (Effort Rev)"
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

 'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.createPivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
   
  pvt.AddFields RowFields:="RM", ColumnFields:="Month"
  pvt.AddDataField pvt.PivotFields("TOTAL_Effort_REVENUE"), "Sum of TOTAL_Effort_REVENUE", xlSum
  pvt.PivotFields("Mgr").Orientation = xlPageField
  pvt.ManualUpdate = False
  
  sht.Range(ActiveSheet.Cells(5, 2), ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).NumberFormat = "#,###.##"
  pvt.InGridDropZones = True
   pvt.RowAxisLayout xlOutlineRow
   pvt.ColumnGrand = True
   pvt.SubtotalLocation xlAtBottom

End Sub

Sub createPivotTable_TR(id As Integer)
' id is the identifier telling the Name of FINAL/TOTAL revenue
' In CB it is called FINAL_REVENUE in CGG3 it is called TOTAL_REVENUE
Dim dataRangeStr As String
If id = 0 Then
    dataRangeStr = "FINAL_REVENUE"
Else
    dataRangeStr = "TOTAL_REVENUE"
End If
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim sht As Worksheet
Dim SrcData As String
Dim RangeStr As String
Dim StartPvt As String
Dim NumOfRows As Integer
Dim NumOfCols As Integer

NumOfRows = ActiveSheet.UsedRange.Rows.Count
NumOfCols = ActiveSheet.UsedRange.Columns.Count
RangeStr = "A1:" & ActiveSheet.Cells(NumOfRows, NumOfCols).Address

SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

Set sht = Sheets.Add
sht.Name = "PIVOT (Total Rev)"
StartPvt = "'" & sht.Name & "'" & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

 'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.createPivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
 
   
  pvt.AddFields RowFields:="RM", ColumnFields:="Month"
  pvt.AddDataField pvt.PivotFields(dataRangeStr), "Sum of " & dataRangeStr, xlSum
  pvt.PivotFields("Mgr").Orientation = xlPageField
  pvt.ManualUpdate = False
  
  sht.Range(ActiveSheet.Cells(5, 2), ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).NumberFormat = "#,###.##"
  pvt.InGridDropZones = True
  pvt.RowAxisLayout xlOutlineRow
  pvt.ColumnGrand = True
  pvt.SubtotalLocation xlAtBottom
   
   

End Sub


Sub CreatePivotTableForAll_ZT()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim NumOfRows As Long
Dim RangeStr As String


NumOfRows = ActiveSheet.UsedRange.Rows.Count
RangeStr = "A1:AH" & NumOfRows


 

 SrcData = ActiveSheet.Name & "!" & Range(RangeStr).Address(ReferenceStyle:=xlR1C1)

 Set sht = Sheets.Add
 sht.Name = "PivotTableAllZT"

'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.createPivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
   
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
    fName = saveToPath & "\" & "AUM " & monthStr & " " & yearStr & " - Breakdown by Zone and Tiers"
                
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, Password:="Universal" & Right(yearStr, 2)
        .Close 0
    End With
    Application.DisplayAlerts = True

End Sub

Function getMonthStr(str As String) As String
    Dim ans As String
    ans = "20" & Right(str, 2) & "."
    Select Case Left(str, 3)
    Case "JAN":
        ans = ans & "01"
    Case "FEB":
        ans = ans & "02"
    Case "MAR":
        ans = ans & "03"
    Case "APR":
        ans = ans & "04"
    Case "MAY":
        ans = ans & "05"
    Case "JUN":
        ans = ans & "06"
    Case "JUL":
        ans = ans & "07"
    Case "AUG":
        ans = ans & "08"
    Case "SEP":
        ans = ans & "09"
    Case "OCT":
        ans = ans & "10"
    Case "NOV":
        ans = ans & "11"
    Case "DEC":
        ans = ans & "12"
    Case Else:
        ans = ans & "##"
    End Select
         
    getMonthStr = ans

End Function


Sub InsertCodePT(wkbk As Workbook, tp As String)

    Dim Command1 As String
    Dim command2 As String
    Command1 = getCommandStr1()

    Dim VBComps As VBComponents

    Set VBComps = wkbk.VBProject.VBComponents

    Dim VBComp As VBComponent
    Dim VBCodeMod As CodeModule

    Dim oSheet As Worksheet
    For Each oSheet In wkbk.Worksheets
        If oSheet.Name = tp & " ABU Pivot" Then
            Set VBComp = VBComps(oSheet.CodeName)
            Set VBCodeMod = VBComp.CodeModule
            InsertCode VBCodeMod, Command1
        End If
    Next oSheet
    
End Sub

Private Function InsertCode(VBCodeMod As CodeModule, Command1 As String)

    Dim LineNum As Long
    With VBCodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, Command1
    End With

End Function

Function getCommandStr1() As String
Dim ans As String
ans = Chr(13) + "Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable)"
ans = ans & Chr(13) & "Target.EnableDataValueEditing = True"
ans = ans & Chr(13) & "Dim TotCol As Integer"
ans = ans & Chr(13) & "Dim TotRow As Integer"
ans = ans & Chr(13) & " TotRow = 0"
ans = ans & Chr(13) & " For i = 1 To ActiveSheet.UsedRange.Columns.Count"
ans = ans & Chr(13) & "    If ActiveSheet.Cells(4, i).Value Like ""*A-B-U"" Then"
ans = ans & Chr(13) & "        TotCol = i"
ans = ans & Chr(13) & "    End If"
ans = ans & Chr(13) & " Next"
ans = ans & Chr(13) & " For i = 4 To ActiveSheet.UsedRange.Rows.Count"
ans = ans & Chr(13) & "    If ActiveSheet.Cells(i, 1).Value = """" Then"
ans = ans & Chr(13) & "        TotRow = i"
ans = ans & Chr(13) & "        Exit For"
ans = ans & Chr(13) & "   End If"
ans = ans & Chr(13) & " Next"
ans = ans & Chr(13) & " If TotRow = 5 Then"
ans = ans & Chr(13) & " ActiveSheet.Columns(TotCol + 1).Resize(, 100).Delete"
ans = ans & Chr(13) & " Exit Sub"
ans = ans & Chr(13) & " End If"
ans = ans & Chr(13) & " ActiveSheet.Columns(TotCol).NumberFormat = ""@"""
ans = ans & Chr(13) & " ActiveSheet.Columns(TotCol + 1).Resize(, 100).Delete"
ans = ans & Chr(13) & " If TotRow <> 0 Then"
ans = ans & Chr(13) & "    If TotRow = 5 Then TotRow = 6"
ans = ans & Chr(13) & "    ActiveSheet.Rows(TotRow).Resize(100).Delete'"
ans = ans & Chr(13) & " End If"
ans = ans & Chr(13) & " ActiveSheet.Columns(TotCol + 1).NumberFormat = ""0.00%"""
ans = ans & Chr(13) & " ActiveSheet.Cells(4, TotCol + 1).Value = ""AB/ABU"""
ans = ans & Chr(13) & " ActiveSheet.Cells(4, TotCol + 1).Font.Bold = True"
ans = ans & Chr(13) & " ActiveSheet.Cells(4, TotCol + 1).HorizontalAlignment = xlRight"
ans = ans & Chr(13) & " ActiveSheet.Columns(TotCol).HorizontalAlignment = xlRight"
ans = ans & Chr(13) & " ActiveSheet.Cells(4, TotCol + 1).Interior.Color = RGB(220, 230, 241)"
ans = ans & Chr(13) & " ActiveSheet.Cells(3, TotCol + 1).Interior.Color = RGB(220, 230, 241)"
ans = ans & Chr(13) & " Dim NumA As Integer"
ans = ans & Chr(13) & " Dim NumB As Integer"
ans = ans & Chr(13) & " Dim NumU As Integer"
ans = ans & Chr(13) & " Dim ComboValue As Double"
ans = ans & Chr(13) & " For i = 5 To ActiveSheet.UsedRange.Rows.Count"
ans = ans & Chr(13) & "    If ActiveSheet.Cells(i, 1) <> """" And (Not ActiveSheet.Cells(i, TotCol).Value Like ""*-*-*"") Then"
ans = ans & Chr(13) & "        ComboValue = ActiveSheet.Cells(i, TotCol).Value"
ans = ans & Chr(13) & "        NumA = Int((ComboValue / 100))"
ans = ans & Chr(13) & "        NumB = Int((ComboValue - NumA * 100))"
ans = ans & Chr(13) & "        NumU = Round(100 * (ComboValue - Int(ComboValue)))"
ans = ans & Chr(13) & "        ActiveSheet.Cells(i, TotCol).Value = CStr(NumA) & "" - "" & CStr(NumB) & "" - "" & CStr(NumU)"
ans = ans & Chr(13) & "        ActiveSheet.Cells(i, TotCol + 1).Value = (NumA + NumB) / (NumA + NumB + NumU)"
ans = ans & Chr(13) & "    End If"
ans = ans & Chr(13) & " Next"
ans = ans & Chr(13) & "Target.EnableDataValueEditing = False"
ans = ans & Chr(13) & "End Sub"

ans = ans & Chr(13) & "Private Sub Worksheet_Activate()"
ans = ans & Chr(13) & "Application.DisplayAlerts = False"
ans = ans & Chr(13) & "End Sub"
ans = ans & Chr(13) & "Private Sub Worksheet_Deactivate()"
ans = ans & Chr(13) & "Application.DisplayAlerts = True"
ans = ans & Chr(13) & "End Sub"
getCommandStr1 = ans
End Function


Private Sub CPC_ABU_Button_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

If ConcFiles_CPC_UPGRADE = 0 Then
    Application.ScreenUpdating = True
    Exit Sub
End If
'MsgBox ActiveWorkbook.Name

ProtectVBProject ActiveWorkbook, "mo"
Application.Wait (Now + TimeValue("0:00:10"))
Dim fName As String
fName = saveToPath & "\" & Format(Date, "YYYY") & " YTD CPCs ABU (" & Format(Date, "YYYY.MM.DD") & ").xlsm"
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, FileFormat:=52, Password:="Citi2015"
ActiveWorkbook.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Private Sub PB_ABU_Button_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

If ConcFiles_PB_UPGRADE = 0 Then
    Application.ScreenUpdating = True
    Exit Sub
End If

ProtectVBProject ActiveWorkbook, "mo"
Application.Wait (Now + TimeValue("0:00:10"))
Dim fName As String
fName = saveToPath & "\" & Format(Date, "YYYY") & " YTD PBs ABU (" & Format(Date, "YYYY.MM.DD") & ").xlsm"
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, FileFormat:=52, Password:="Citi2015"
ActiveWorkbook.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Private Sub RM_ABU_Button_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

If ConcFiles_RM_UPGRADE = 0 Then
    Application.ScreenUpdating = True
    Exit Sub
End If


ProtectVBProject ActiveWorkbook, "mo"
Application.Wait (Now + TimeValue("0:00:05"))
Dim fName As String
fName = saveToPath & "\" & Format(Date, "YYYY") & " YTD RMs ABU (" & Format(Date, "YYYY.MM.DD") & ").xlsm"
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, FileFormat:=52, Password:="Citi2015"
ActiveWorkbook.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub



Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable)
Target.EnableDataValueEditing = True
Dim TotCol As Integer
Dim TotRow As Integer
 TotRow = 0
 For i = 1 To ActiveSheet.UsedRange.Columns.Count
    If ActiveSheet.Cells(4, i).Value = "Grand Total" Then
        TotCol = i
    End If
 Next
 For i = 4 To ActiveSheet.UsedRange.Rows.Count
    If ActiveSheet.Cells(i, 1).Value = "" Then
        TotRow = i
        Exit For
   End If
 Next
 ActiveSheet.Columns(TotCol).NumberFormat = "@"
 ActiveSheet.Columns(TotCol + 1).Resize(, 100).Delete
 If TotRow <> 0 Then
    ActiveSheet.Rows(TotRow).Resize(100).Delete '
 End If
 ActiveSheet.Columns(TotCol + 1).NumberFormat = "0.00%"
 ActiveSheet.Cells(4, TotCol + 1).Value = "AB/ABU"
 ActiveSheet.Cells(4, TotCol + 1).Font.Bold = True
 ActiveSheet.Cells(4, TotCol + 1).HorizontalAlignment = xlRight
 ActiveSheet.Columns(TotCol).HorizontalAlignment = xlRight
 ActiveSheet.Cells(4, TotCol + 1).Interior.Color = RGB(220, 230, 241)
 ActiveSheet.Cells(3, TotCol + 1).Interior.Color = RGB(220, 230, 241)
 Dim NumA As Integer
 Dim NumB As Integer
 Dim NumU As Integer
 Dim ComboValue As Double
 For i = 5 To ActiveSheet.UsedRange.Rows.Count
    If ActiveSheet.Cells(i, 1) <> "" And (Not ActiveSheet.Cells(i, TotCol).Value Like "*-*-*") Then
        ComboValue = ActiveSheet.Cells(i, TotCol).Value
        NumA = Int((ComboValue / 100))
        NumB = Int((ComboValue - NumA * 100))
        NumU = Round(100 * (ComboValue - Int(ComboValue)))
        ActiveSheet.Cells(i, TotCol).Value = CStr(NumA) & " - " & CStr(NumB) & " - " & CStr(NumU)
        ActiveSheet.Cells(i, TotCol + 1).Value = (NumA + NumB) / (NumA + NumB + NumU)
    End If
 Next
Target.EnableDataValueEditing = False
End Sub

Private Sub SaveDirButton_Click()
    saveToPath = GetFolder
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

Private Sub CheckPathButton_Click()
If saveToPath = "" Then
MsgBox ("YOU HAVE NOT SELECT ANY PATH")
Else
MsgBox (saveToPath)
End If
End Sub


Private Sub CBRevButton_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

ConcFiles_CB_REV

Dim fName As String
fName = saveToPath & "\CB_Contracted_Revenue_Summary_YTD (" & Format(Date, "YYYY.MM.DD") & ").xlsx"
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
ActiveWorkbook.Close
Application.ScreenUpdating = True

End Sub

Private Sub CGG3RevButton_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

ConcFiles_CGG3_REV

Dim fName As String
fName = saveToPath & "\CGG3_Contracted_Revenue_Summary_YTD (" & Format(Date, "YYYY.MM.DD") & ").xlsx"
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
ActiveWorkbook.Close
Application.ScreenUpdating = True
End Sub

Private Sub CGCBCreditButton_Click()
Application.ScreenUpdating = False
If saveToPath = "" Then
MsgBox ("YOU HAVE TO SELECT A FOLDER FOR SAVING FILES")
Exit Sub
End If

ConcFiles_CGCB_CREDIT

Dim fName As String
fName = saveToPath & "\CGCB_Contracted_Credit_Summary_YTD (" & Format(Date, "YYYY.MM.DD") & ").xlsx"
ActiveWorkbook.SaveAs fileName:=fName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
ActiveWorkbook.Close
Application.ScreenUpdating = True
End Sub




Sub ProtectVBProject(ByRef WB As Workbook, ByVal Password As String)
Dim vbProj As Object

Set vbProj = WB.VBProject

If vbProj.Protection = 1 Then Exit Sub

Set Application.VBE.ActiveVBProject = vbProj

Application.VBE.CommandBars(1).FindControl(id:=2578, recursive:=True).Execute

SendKeys "+{TAB}{RIGHT}%V{+}{TAB}" & Password & "{TAB}" & Password & "~", True

End Sub



