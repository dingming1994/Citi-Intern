Public START_DAY_IN_MONTH As String
Public END_DAY_IN_MONTH As String
Public saveToPath As String
Public monthstr As String

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

Private Sub CreateFileButton_Click()
CreateSSFile
End Sub

Private Sub CheckingButton_Click()
UploadFileChecking
End Sub

Private Sub sepSSDButton_Click()
SSD_USER_SEPERATION
End Sub


Private Sub CreateSSFile()
 
If saveToPath = "" Then
    MsgBox "Please select the saving directory!"
    Exit Sub
End If

Application.FileDialog(msoFileDialogOpen).Title = "Select the SOP masterlist"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
Else: Exit Sub
End If
monthstr = InputBox(Prompt:="Please input the month and year (MMMYY)", _
          Title:="Please input the month and year")
START_DAY_IN_MONTH = InputBox(Prompt:="Please input the start date of month (e.g. 10/1/2015)", _
          Title:="Please input the start date of month")
END_DAY_IN_MONTH = InputBox(Prompt:="Please input the last date of month (e.g. 10/31/2015)", _
          Title:="Please input the last date of month")

'On Error Resume Next
Dim wopbk As Workbook
Set sopbk = Workbooks.Open(strpath, False)
Dim sopsht As Worksheet
Set sopsht = sopbk.worksheets("SOP")
RCAO_MAPPING sopsht
SUPERVISOR_MAPPING sopsht
SSD_USER sopsht
SOP_TARGET sopsht
createSSDeleteFile
Application.DisplayAlerts = False
sopbk.Close
Application.DisplayAlerts = True
End Sub

Sub createSSDeleteFile()
ThisWorkbook.worksheets("RCAO_MAPPING_D").Copy
ActiveSheet.Name = "Sheet1"
ActiveSheet.Cells(2, 2) = START_DAY_IN_MONTH
ActiveSheet.Cells(2, 3) = END_DAY_IN_MONTH
ActiveWorkbook.SaveAs saveToPath & "\SSD_RCAO_MAPPING_DELETE_G2C_" & monthstr & ".xlsx"
ActiveWorkbook.Close

ThisWorkbook.worksheets("SUPERVISOR_MAPPING_D").Copy
ActiveSheet.Name = "Sheet1"
ActiveSheet.Cells(2, 2) = START_DAY_IN_MONTH
ActiveSheet.Cells(2, 3) = END_DAY_IN_MONTH
ActiveWorkbook.SaveAs saveToPath & "\SSD_SUPERVISOR_MAPPING_DELETE_G2C_" & monthstr & ".xlsx"
ActiveWorkbook.Close

ThisWorkbook.worksheets("SSD_USER_D").Copy
ActiveSheet.Name = "Sheet1"
ActiveSheet.Cells(2, 1) = END_DAY_IN_MONTH
ActiveWorkbook.SaveAs saveToPath & "\SSD_USER_DELETE_G2C_" & monthstr & ".xlsx"
ActiveWorkbook.Close

ThisWorkbook.worksheets("SOP_TARGET_D").Copy
ActiveSheet.Name = "Sheet1"
ActiveSheet.Cells(2, 1) = END_DAY_IN_MONTH
ActiveWorkbook.SaveAs saveToPath & "\SSI_SOP_TARGET_DELETE_" & monthstr & ".xlsx"
ActiveWorkbook.Close
End Sub

Sub RCAO_MAPPING(sopsht As Worksheet)

'SetStartEndDate

'Set sopbk = Workbooks.Open("I:\CAP_Profile_PRD65\Desktop\SS Upload System\SOP Masterlist 2015 - Oct 15 (2015.10.16).xls", False)

'Set sopsht = sopbk.worksheets("SOP")

Set outbk = Workbooks.Add
' delete extra sheet and leave only one sheet
Application.DisplayAlerts = False
count = 0
For Each sht In outbk.worksheets
    count = count + 1
    If count <> 1 Then
        sht.Delete
    End If
Next
Application.DisplayAlerts = True

Set outsht = outbk.worksheets(1)
outsht.Name = "Sheet1"

' fill in title row

outsht.Range("A1") = "USER_USER_ID"
outsht.Range("B1") = "USER_RC_CDE"
outsht.Range("C1") = "USER_AO_CDE"
outsht.Range("D1") = "USER_PORT"
outsht.Range("E1") = "USER_REG_CDE"
outsht.Range("F1") = "RCAO_USER_BIZ_UNIT_CDE"
outsht.Range("G1") = "USER_RCAO_EFF_DT"
outsht.Range("H1") = "USER_RCAO_END_DT"

'fill in SOP data
With sopsht
.Activate
endrow = .Range("BG2").End(xlDown).Row
.Range(.Range("BG2"), .Range("BG" & CStr(endrow))).Copy
outsht.Range("A2").PasteSpecial xlPasteValues
.Range(.Range("O2"), .Range("O" & CStr(endrow))).Copy
outsht.Range("B2").PasteSpecial xlPasteValues
.Range(.Range("N2"), .Range("N" & CStr(endrow))).Copy
outsht.Range("C2").PasteSpecial xlPasteValues
outsht.Range("E2:E" & CStr(endrow)).Value = "P"
outsht.Range("F2:F" & CStr(endrow)).Value = "RB"

End With
'fill in back-end data
For i = 2 To ThisWorkbook.worksheets("RCAO_MAPPING_BE").UsedRange.Rows.count
    ThisWorkbook.worksheets("RCAO_MAPPING_BE").Rows(i).Copy
    outsht.Rows(endrow + i - 1).PasteSpecial xlPasteValues
Next
outsht.Range("G2:G" & outsht.UsedRange.Rows.count).Value = START_DAY_IN_MONTH
outsht.Range("H2:H" & outsht.UsedRange.Rows.count).Value = END_DAY_IN_MONTH

'ending row

outsht.Range("A" & CStr(outsht.UsedRange.Rows.count + 1) & ":F" & CStr(outsht.UsedRange.Rows.count + 1)).Value = "*END"

outbk.SaveAs saveToPath & "\SSD_RCAO_MAPPING_BJ_APPEND_" & monthstr & ".xlsx"
outbk.Close

End Sub



Sub SUPERVISOR_MAPPING(sopsht As Worksheet)

'SetStartEndDate
'Set sopbk = Workbooks.Open("I:\CAP_Profile_PRD65\Desktop\SS Upload System\SOP Masterlist 2015 - Oct 15 (2015.10.16).xls", False)

'Set sopsht = sopbk.worksheets("SOP")

Set outbk = Workbooks.Add
' delete extra sheet and leave only one sheet
Application.DisplayAlerts = False
count = 0
For Each sht In outbk.worksheets
    count = count + 1
    If count <> 1 Then
        sht.Delete
    End If
Next
Application.DisplayAlerts = True

Set outsht = outbk.worksheets(1)
outsht.Name = "Sheet1"

' fill in title row

outsht.Range("A1") = "USER_USER_ID"
outsht.Range("B1") = "HIERARCHY_TYPE"
outsht.Range("C1") = "USER_SUPERVISOR_ID"
outsht.Range("D1") = "USER_BIZ_UNIT_CDE"
outsht.Range("E1") = "USER_RELATIONSHIP_LEVEL"
outsht.Range("F1") = "USER_SUP_EFF_DT"
outsht.Range("G1") = "USER_SUP_END_DT"

'fill in SOP data
With sopsht
.Activate
endrow = .Range("BG2").End(xlDown).Row
.Range(.Range("BG2"), .Range("BG" & CStr(endrow))).Copy
outsht.Range("A2").PasteSpecial xlPasteValues
.Range(.Range("BJ2"), .Range("BJ" & CStr(endrow))).Copy
outsht.Range("B2").PasteSpecial xlPasteValues
.Range(.Range("BI2"), .Range("BI" & CStr(endrow))).Copy
outsht.Range("C2").PasteSpecial xlPasteValues
outsht.Range("D2:D" & CStr(endrow)).Value = "RB"
outsht.Range("E2:E" & CStr(endrow)).Value = "7"
End With

'fill in back-end data

For i = 2 To ThisWorkbook.worksheets("RCAO_MAPPING_BE").UsedRange.Rows.count
    ThisWorkbook.worksheets("SUPERVISOR_MAPPING_BE").Rows(i).Copy
    outsht.Rows(endrow + i - 1).PasteSpecial xlPasteValues
Next
outsht.Range("F2:F" & outsht.UsedRange.Rows.count).Value = START_DAY_IN_MONTH
outsht.Range("G2:G" & outsht.UsedRange.Rows.count).Value = END_DAY_IN_MONTH

'Remove duplicates
outsht.UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3)

'ending row
outsht.Range("A" & CStr(outsht.UsedRange.Rows.count + 1) & ":F" & CStr(outsht.UsedRange.Rows.count + 1)).Value = "*END"


outbk.SaveAs saveToPath & "\SSD_SUPERVISOR_MAPPING_BJ_" & monthstr & ".xlsx"
outbk.Close
End Sub


Sub SSD_USER(sopsht As Worksheet)

'Set sopbk = Workbooks.Open("I:\CAP_Profile_PRD65\Desktop\SS Upload System\SOP Masterlist 2015 - Oct 15 (2015.10.16).xls", False)

'SetStartEndDate
'Set sopsht = sopbk.worksheets("SOP")

Set outbk = Workbooks.Add
' delete extra sheet and leave only one sheet
Application.DisplayAlerts = False
count = 0
For Each sht In outbk.worksheets
    count = count + 1
    If count <> 1 Then
        sht.Delete
    End If
Next
Application.DisplayAlerts = True

Set outsht = outbk.worksheets(1)
outsht.Name = "Sheet1"

' fill in title row
ThisWorkbook.worksheets("SSD_USER_BE").Rows(1).Copy
outsht.Rows(1).PasteSpecial xlPasteValues

With sopsht
    .Activate
    endrow = .Range("BG2").End(xlDown).Row
    idx = 2
    prevName = ""
    On Error Resume Next
    For i = 2 To endrow
        nName = .Cells(i, 3)
        If nName <> prevName Then
        outsht.Cells(idx, 1) = .Range("BG" & CStr(i))
        outsht.Cells(idx, 2) = END_DAY_IN_MONTH
        outsht.Cells(idx, 3) = "RB"
        outsht.Cells(idx, 4) = .Range("BH" & CStr(i))
        outsht.Cells(idx, 5) = .Range("BJ" & CStr(i))
        outsht.Cells(idx, 6) = .Range("BD" & CStr(i))
        outsht.Cells(idx, 7) = .Range("BA" & CStr(i))
        outsht.Cells(idx, 13) = nName
        outsht.Range("AV" & CStr(idx)) = .Range("AO" & CStr(i))
        outsht.Range("BV" & CStr(idx)) = .Range("BD" & CStr(i))
        outsht.Range("DY" & CStr(idx)) = .Range("X" & CStr(i))
        outsht.Range("EN" & CStr(idx)) = .Range("AF" & CStr(i))
        outsht.Range("EO" & CStr(idx)) = "P"
        prevName = nName
        idx = idx + 1
        End If
        
    Next
 End With
        
For i = 2 To ThisWorkbook.worksheets("SSD_USER_BE").UsedRange.Rows.count
    ThisWorkbook.worksheets("SSD_USER_BE").Rows(i).Copy
    outsht.Rows(idx + i - 2).PasteSpecial xlPasteValues
Next
outsht.Range("B2:B" & outsht.UsedRange.Rows.count).Value = END_DAY_IN_MONTH

'ending row
outsht.Range("A" & CStr(outsht.UsedRange.Rows.count + 1) & ":EP" & CStr(outsht.UsedRange.Rows.count + 1)).Value = "*END"

outbk.SaveAs saveToPath & "\SSD_USER_BJ_" & monthstr & ".xlsx"
outbk.Close

End Sub


Sub SOP_TARGET(sopsht As Worksheet)

'Set sopbk = Workbooks.Open("I:\CAP_Profile_PRD65\Desktop\SS Upload System\SOP Masterlist 2015 - Oct 15 (2015.10.16).xls", False)
'SetStartEndDate
'Set sopsht = sopbk.worksheets("SOP")

Set outbk = Workbooks.Add
' delete extra sheet and leave only one sheet
Application.DisplayAlerts = False
count = 0
For Each sht In outbk.worksheets
    count = count + 1
    If count <> 1 Then
        sht.Delete
    End If
Next
Application.DisplayAlerts = True

Set outsht = outbk.worksheets(1)
outsht.Name = "Sheet1"

' fill in title row
ThisWorkbook.worksheets("SOP_TARGET_BE").Rows(1).Copy
outsht.Rows(1).PasteSpecial xlPasteValues

With sopsht
    .Activate
    endrow = .Range("BG2").End(xlDown).Row
    idx = 2
    prevName = ""
    On Error Resume Next
    For i = 2 To endrow
        nName = .Cells(i, 3)
        flag = 0
        If .Range("BJ" & CStr(i)) = "CPCRM" Then
            If .Range("Y" & CStr(i)) > 10 And .Range("Y" & CStr(i)) < 999990 Then
                flag = 1
            End If
        Else
            If .Range("AL" & CStr(i)) > 10 And .Range("AL" & CStr(i)) < 999990 Then
                flag = 1
            End If
        End If
        If nName <> prevName And flag = 1 Then
        outsht.Cells(idx, 1) = END_DAY_IN_MONTH
        outsht.Cells(idx, 2) = "RB"
        outsht.Cells(idx, 3) = .Range("BG" & CStr(i))
        
        outsht.Cells(idx, 5) = .Range("BJ" & CStr(i))
        If outsht.Cells(idx, 5) = "CPCRM" Then
            outsht.Cells(idx, 5) = "CPC"
            outsht.Cells(idx, 4) = "Total_Original_SOP_CGCB"
            outsht.Cells(idx, 7) = .Range("Y" & CStr(i))
            
        Else
            outsht.Cells(idx, 4) = "Total_Original_TheresholdCGCB"
            outsht.Cells(idx, 7) = .Range("AL" & CStr(i))
        End If
        
     
        outsht.Cells(idx, 8) = "MTD"
        outsht.Cells(idx, 9) = "I"
        outsht.Cells(idx, 10) = "N"
      

        prevName = nName
        idx = idx + 1
        End If
        
    Next
 End With
        
For i = 2 To ThisWorkbook.worksheets("SOP_TARGET_BE").UsedRange.Rows.count
    ThisWorkbook.worksheets("SOP_TARGET_BE").Rows(i).Copy
    outsht.Rows(idx + i - 2).PasteSpecial xlPasteValues
Next
outsht.Range("A2:A" & outsht.UsedRange.Rows.count).Value = END_DAY_IN_MONTH

'ending row
outsht.Range("A" & CStr(outsht.UsedRange.Rows.count + 1) & ":J" & CStr(outsht.UsedRange.Rows.count + 1)).Value = "*END"
 
outbk.SaveAs saveToPath & "\SSI_SOP_TARGET_" & monthstr & ".xlsx"
outbk.Close


End Sub


Sub UploadFileChecking()

If saveToPath = "" Then
    MsgBox "Please select the saving directory!"
    Exit Sub
End If

Set oApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFSO.GetFolder(saveToPath)
Dim rcaobk, supbk, userbk, outbk As Workbook

For Each objfile In objfolder.Files
        
    If objfile.Name Like "*SSD_RCAO_MAPPING_BJ_APPEND*.xlsx" Then
        Set rcaobk = Workbooks.Open(objfile.Path, False)
        Set rcaosht = rcaobk.worksheets(1)
    End If
    If objfile.Name Like "*SSD_SUPERVISOR_MAPPING_BJ*.xlsx" Then
        Set supbk = Workbooks.Open(objfile.Path, False)
        Set supsht = supbk.worksheets(1)
    End If
    If objfile.Name Like "*SSD_USER_BJ*.xlsx" Then
        Set userbk = Workbooks.Open(objfile.Path, False)
        Set usersht = userbk.worksheets(1)
    End If

Next

Set outbk = Workbooks.Add
Set outsht = outbk.worksheets(1)

With rcaosht
    .Activate
    rowcnt = .Range("A2").End(xlDown).Row - 1
    .Range("A2:A" & CStr(rowcnt)).Copy
End With

outsht.Range("A1").PasteSpecial xlPasteValues

With usersht
    .Activate
    rowcnt = .Range("A2").End(xlDown).Row - 1
    .Range("A2:A" & CStr(rowcnt)).Copy
End With

outsht.Range("B1").PasteSpecial xlPasteValues

With supsht
    .Activate
    rowcnt = .Range("A2").End(xlDown).Row - 1
    .Range("A2:A" & CStr(rowcnt)).Copy
    outsht.Range("C1").PasteSpecial xlPasteValues
    .Range("C2:C" & CStr(rowcnt)).Copy
    outsht.Range("D1").PasteSpecial xlPasteValues
End With

Set osht = ThisWorkbook.worksheets("Checking_Result")
osht.UsedRange.Clear

osht.Range("I1") = "Same user and supervisor in SUPERVISOR_MAPPING"
For i = 2 To supsht.UsedRange.Rows.count - 1
    If LCase(supsht.Cells(i, 1)) = LCase(supsht.Cells(i, 3)) Then
        osht.Cells(osht.Range("I10000").End(xlUp).Row + 1, 9) = supsht.Cells(i, 1)
    End If
Next

Application.DisplayAlerts = False
rcaobk.Close
userbk.Close
supbk.Close
Application.DisplayAlerts = True

outsht.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
outsht.Range("B:B").RemoveDuplicates Columns:=1, Header:=xlNo
outsht.Range("C:C").RemoveDuplicates Columns:=1, Header:=xlNo
outsht.Range("D:D").RemoveDuplicates Columns:=1, Header:=xlNo



osht.Range("A1") = "Missing in SSD_USER but in SUPERVISOR_MAPPING"
' checking 1: all ID in spvsr mapping should be in SSD_USER
With outsht
    For i = 1 To .Range("C1").End(xlDown).Row
        isFound = 0
        For j = 1 To .Range("B1").End(xlDown).Row
            If LCase(.Cells(j, 2)) = LCase(.Cells(i, 3)) Then
                isFound = 1
                Exit For
            End If
        Next
        
        If isFound = 0 Then
            On Error Resume Next
            osht.Cells(osht.Range("A10000").End(xlUp).Row + 1, 1).Value = .Cells(i, 3)
        End If
    Next
End With

osht.Range("C1") = "Missing in RCAO_MAPPING but in SUPERVISOR_MAPPING"
' checking 2: all spvsr in spvsr mapping should be in RCAO Mapping
With outsht
    For i = 1 To .Range("D1").End(xlDown).Row
        isFound = 0
        For j = 1 To .Range("A1").End(xlDown).Row
            If LCase(.Cells(j, 1)) = LCase(.Cells(i, 4)) Then
                isFound = 1
                Exit For
            End If
        Next
        
        If isFound = 0 Then
            On Error Resume Next
            osht.Cells(osht.Range("C10000").End(xlUp).Row + 1, 3).Value = .Cells(i, 4)
        End If
    Next
End With

osht.Range("E1") = "Missing in RCAO_MAPPING but in SSD_USER"
' checking 3: all ID in SSD_USER should be in RCAO Mapping
With outsht
    For i = 1 To .Range("B1").End(xlDown).Row
        isFound = 0
        For j = 1 To .Range("A1").End(xlDown).Row
            If LCase(.Cells(j, 1)) = LCase(.Cells(i, 2)) Then
                isFound = 1
                Exit For
            End If
        Next
        
        If isFound = 0 Then
            On Error Resume Next
            osht.Cells(osht.Range("E10000").End(xlUp).Row + 1, 5).Value = .Cells(i, 2)
        End If
    Next
End With

osht.Range("G1") = "Missing in SUPERVISOR_MAPPING but in SSD_USER"
' checking 4: all ID in SSD_USER should be in supervisor Mapping
With outsht
    For i = 1 To .Range("B1").End(xlDown).Row
        isFound = 0
        For j = 1 To .Range("C1").End(xlDown).Row
            If LCase(.Cells(j, 3)) = LCase(.Cells(i, 2)) Then
                isFound = 1
                Exit For
            End If
        Next
        
        If isFound = 0 Then
            On Error Resume Next
            osht.Cells(osht.Range("G10000").End(xlUp).Row + 1, 7).Value = .Cells(i, 2)
        End If
    Next
End With

With osht
    .Rows(1).RowHeight = 40
    .Rows(1).WrapText = True
    .Columns(1).ColumnWidth = 28
    .Columns(3).ColumnWidth = 28
    .Columns(5).ColumnWidth = 28
    .Columns(7).ColumnWidth = 28
    .Columns(9).ColumnWidth = 28
    .Columns(2).ColumnWidth = 5
    .Columns(4).ColumnWidth = 5
    .Columns(6).ColumnWidth = 5
    .Columns(8).ColumnWidth = 5
    For i = 1 To 9
        If i Mod 2 = 0 Then
            .Columns(i).ColumnWidth = 5
        Else
            .Columns(i).ColumnWidth = 28
            .Cells(1, i).Font.Bold = True
            .Cells(1, i).Interior.Color = rgbDarkGrey
        End If
    Next

End With

Application.DisplayAlerts = False
outbk.Close
Application.DisplayAlerts = True

MsgBox "Data Check is Done. Please see" & vbNewLine & "Checking_Result tab for details."


End Sub

Sub SSD_USER_SEPERATION()
Application.FileDialog(msoFileDialogOpen).Title = "Select the SSD_USER File"

Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show

If intChoice <> 0 Then
    'get the file path selected by the user
    strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
Else: Exit Sub
End If

Set bk = Workbooks.Open(strpath, False)
Set sht = bk.worksheets(1)

Dim numOfBk As Integer
numOfBk = Int((sht.UsedRange.Rows.count - 3) / 100 + 1)

For i = 1 To numOfBk
    Set nbk = Workbooks.Add
    Set nsht = nbk.worksheets(1)
    sht.Rows(1).Copy
    nsht.Rows(1).PasteSpecial xlPasteValues
    If i <> numOfBk Then
        sht.Rows(CStr(2 + 100 * (i - 1)) & ":" & CStr(1 + 100 * i)).Copy
        nsht.Rows(2).PasteSpecial xlPasteValues
        sht.Rows(sht.Range("A1").End(xlDown).Row).Copy
        nsht.Rows(102).PasteSpecial xlPasteValues
    Else
        sht.Rows(CStr(2 + 100 * (i - 1)) & ":" & sht.Range("A1").End(xlDown).Row - 1).Copy
        nsht.Rows(2).PasteSpecial xlPasteValues
        sht.Rows(sht.Range("A1").End(xlDown).Row).Copy
        nsht.Rows(nsht.UsedRange.Rows.count + 1).PasteSpecial xlPasteValues
    End If
    nameofbk = Left(strpath, InStrRev(strpath, ".") - 1) & "_Part_" & CStr(i) & ".xlsx"
    nbk.SaveAs nameofbk
    nbk.Close
Next
    
Application.DisplayAlerts = False
bk.Close
Application.DisplayAlerts = True



End Sub


