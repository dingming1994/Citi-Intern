Public savetopath As String

Sub selectPath()
savetopath = GetFolder
End Sub

Sub GenerateSOPs()
Dim TeamFilter As Variant
Dim RMSheet As Variant
Dim PBSheet As Variant
Dim TeamFieldNum As Variant
Dim EmailRecipient As Variant
Dim BranchSOPFull As Variant
Dim WS As Worksheet


Dim outApp As Object
Dim OutMail As Object

Application.DisplayAlerts = False
Application.ScreenUpdating = False

ThisWorkbook.Worksheets("Control Sheet").Select
RMSheet = Range("RMSheet")
PBSheet = Range("PBSheet")

i = 3
While Not IsEmpty(Cells(i, 1))

    Application.CutCopyMode = False
    Application.Calculate

    Worksheets("Control Sheet").Select
    TeamFilter = Worksheets("Control Sheet").Cells(i, 1)
    BranchWBName = Worksheets("Control Sheet").Cells(i, 6)
    EmailRecipient = Worksheets("Control Sheet").Cells(i, 3)
    BranchSOPFull = Worksheets("Control Sheet").Cells(i, 6)
    BranchWBNameshort = Worksheets("Control Sheet").Cells(i, 5)
    SOPmastershort = Worksheets("Control Sheet").Range("SOPMaster")
    
    Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
    If TeamFilter Like "*Direct Banking*" Then
        RMWSName = RMSheet & " " & "DB"
    Else
        RMWSName = RMSheet & " " & TeamFilter
    End If
    WS.Name = RMWSName
    Worksheets(RMSheet).Select
    TeamFieldNum = Application.WorksheetFunction.Match("BM", Rows("3:3"), 0)
    ActiveSheet.Range("$A:$IV").AutoFilter Field:=TeamFieldNum, Criteria1:="=*" & TeamFilter & "*", Operator:=xlAnd
    Sheets(RMSheet).Range("A:AH").Select
    Selection.Copy
    Sheets(RMWSName).Select
    Sheets(RMWSName).Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Range("A1").Select
    rngendrow = Range("D65500").End(xlUp).Row
    rngstartrow = rngendrow - 3
    'Range(Cells(rngstartrow, 1), Cells(rngendrow, 3)).ClearContents
    Rows(CStr(rngendrow + 1) & ":60000").ClearContents
    Columns("AG:AG").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AA:AE").Select
    Selection.Delete Shift:=xlToLeft
    
    Worksheets(RMSheet).Select
    ActiveSheet.Range("$A:$BK").AutoFilter Field:=TeamFieldNum
    Application.CutCopyMode = False
    
    Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
    Worksheets("Control Sheet").Select
    Application.Calculate
    If TeamFilter Like "Direct Banking*" Then
        PBWSname = PBSheet & " " & "DB"
    Else
        PBWSname = PBSheet & " " & TeamFilter
    End If
    WS.Name = PBWSname
    Worksheets(PBSheet).Select
    TeamFieldNum = Application.WorksheetFunction.Match("BM", Rows("3:3"), 0)
    ActiveSheet.Range("$A:$IV").AutoFilter Field:=TeamFieldNum, Criteria1:="=*" & TeamFilter & "*", Operator:=xlAnd
    Sheets(PBSheet).Range("A:AH").Select
    Selection.Copy
    Sheets(PBWSname).Select
    Sheets(PBWSname).Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    'Sheets(PBWSName).Range("A:AH").Value = Sheets(PBSheet).Range("A:AH").Value
    Columns("AG:AG").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AA:AE").Select
    Selection.Delete Shift:=xlToLeft
    
    Worksheets(PBSheet).Select
    ActiveSheet.Range("$A:$IV").AutoFilter Field:=TeamFieldNum
    Application.CutCopyMode = False
    
    Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
    Worksheets("Control Sheet").Select
    Application.Calculate
    SOPWSName = "SOP " & TeamFilter
    WS.Name = SOPWSName
    Worksheets("SOP").Select
    TeamFieldNum = Application.WorksheetFunction.Match("Team", Rows("1:1"), 0)
    ActiveSheet.Range("$A:$IV").AutoFilter Field:=TeamFieldNum, Criteria1:="=*" & TeamFilter & "*", Operator:=xlAnd
    Range("$A:$D, $F:$H, $J:$J").Select
    Selection.Copy
    Sheets(SOPWSName).Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Cells.Select
    Selection.EntireColumn.AutoFit
    Selection.ClearFormats
    Selection.EntireColumn.AutoFit
    'Sheets(SOPWSName).Range("A:J").Value = Sheets("SOP").Range("A:J").Value
    Range("A1").Select
    Worksheets("SOP").Select
    ActiveSheet.Range("$A:$IV").AutoFilter Field:=TeamFieldNum
    Application.CutCopyMode = False
    
    'Sheets(Array(PBWSName, RMWSName, SOPWSName)).move
    
    ActiveWorkbook.Save
    
    Workbooks.Add
    
    ActiveWorkbook.SaveAs Filename:=savetopath & "\" & CStr(BranchWBNameshort)
    
    PBWSname = PBWSname
    RMWSName = RMWSName
    SOPWSName = SOPWSName
    
    'MsgBox PBWSName & " " & RMWSName & " " & SOPWSName
   
    Windows(SOPmastershort).Activate
    ActiveWorkbook.Worksheets(PBWSname).Select
    ActiveSheet.Range("$A:$X").Select
    Selection.Copy
    Windows(BranchWBNameshort).Activate
    ActiveWorkbook.Worksheets("Sheet1").Select
    Worksheets("Sheet1").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Windows(SOPmastershort).Activate
    ActiveWorkbook.Worksheets("Discount").Select
    ActiveSheet.Rows("11:18").Select
    Selection.Copy
    Windows(BranchWBNameshort).Activate
    Worksheets("Sheet1").Select
    rowidx = Range("A60000").End(xlUp).Row
    Cells(rowidx + 4, 1).Select
    Selection.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    Cells.Select
    Selection.EntireColumn.AutoFit
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Range("A1").Select
    Columns("A:A").Select
    Selection.NumberFormat = "mmm-yy"
    Columns("L:M").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("N:N").Select
    Selection.NumberFormat = "0.00"
    Columns("O:O").Select
    Selection.NumberFormat = "0"
    Columns("R:S").Select
    Selection.NumberFormat = "0%"
    
    ActiveSheet.Name = PBWSname
    Range("A1").Select
    
    'Application.CutCopyMode = False
    
    Windows(SOPmastershort).Activate
    ActiveWorkbook.Worksheets(RMWSName).Select
    ActiveSheet.Range("$A:$Z").Select
    'Application.CutCopyMode = False
    Selection.Copy
    Windows(BranchWBNameshort).Activate
    Worksheets("Sheet2").Select
    ActiveSheet.Paste
    Windows(SOPmastershort).Activate
    ActiveWorkbook.Worksheets("Discount").Select
    ActiveSheet.Rows("1:6").Select
    Selection.Copy
    Windows(BranchWBNameshort).Activate
    Worksheets("Sheet2").Select
    rowidx = Range("A60000").End(xlUp).Row
    Cells(rowidx + 4, 1).Select
    Selection.PasteSpecial xlPasteAll
    
    
    Application.CutCopyMode = False
    Cells.Select
    Selection.EntireColumn.AutoFit
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Range("A1").Select
    Columns("A:A").Select
    Selection.NumberFormat = "mmm-yy"
    Columns("L:M").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("N:O").Select
    Selection.NumberFormat = "0.00"
    Columns("T:U").Select
    Selection.NumberFormat = "0%"
    Columns("V:W").Select
    Selection.NumberFormat = "0,000"
    ActiveSheet.Name = RMWSName
    Range("A1").Select
    
    'Application.CutCopyMode = False
        
    Windows(SOPmastershort).Activate
    ActiveWorkbook.Worksheets(SOPWSName).Select
    'ActiveSheet.Range("$A:$A").Select
    'Application.CutCopyMode = False
    'Selection.Copy
    Range("$A:$J").Select
    Selection.Copy
    Windows(BranchWBNameshort).Activate
    Worksheets("Sheet3").Select
    Range("A1").Select
    'ActiveSheet.Paste
    'Worksheets("Sheet3").Cells(1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Application.CutCopyMode = False
    Cells.Select
    Selection.EntireColumn.AutoFit
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Range("A1").Select
    Columns("B:B").Select
    Selection.NumberFormat = "mmm-yy"
    ActiveSheet.Name = SOPWSName
    Range("A1").Select
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    Windows(SOPmastershort).Activate
    Sheets(Array(PBWSname, RMWSName, SOPWSName)).Delete
    
    Application.CutCopyMode = False
    
    Worksheets("Control Sheet").Select
    Application.Calculate
    
    ActiveWorkbook.Save
    
    BranchSOPFull = ThisWorkbook.Worksheets("Control Sheet").Cells(i, 6)
    attachmentpath = ThisWorkbook.Worksheets("Control Sheet").Cells(i, 7)
    attachfilepath = ThisWorkbook.Worksheets("Control Sheet").Cells(2, 9)
    BranchWBNameshort = ThisWorkbook.Worksheets("Control Sheet").Cells(i, 5)
    attachmentpath2 = attachfilepath & BranchWBNameshort
    
'    Set OutApp = CreateObject("Outlook.Application")
    
'    Set OutMail = OutApp.CreateItem(0)

'    OutMail.To = EmailRecipient
'    OutMail.CC = ""
'    OutMail.BCC = ""
'    OutMail.Subject = ThisWorkbook.Worksheets("Control Sheet").Range("EmailSubj")
'    OutMail.HTMLBody = RangetoHTML(ThisWorkbook.Worksheets("Control Sheet").Range("EmailMsg"))
'    OutMail.Attachments.Add attachmentpath2
'    OutMail.Send
    'OutMail.Display

'    Set OutMail = Nothing
'    Set OutApp = Nothing
    
    i = i + 1

Wend

Application.Calculate
Application.DisplayAlerts = True
Application.ScreenUpdating = True

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


Sub sentMail()

Application.ScreenUpdating = False

Dim oApp As Object
Dim outApp As Outlook.Application
Set oApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFSO.GetFolder(savetopath)
Set thiswkbk = ActiveWorkbook

Dim toStr As String
Dim ccStr As String
Dim subjectStr As String
Dim bodyStr As String
Dim monStr As String
Dim yearStr As String
Dim nameStr As String
Dim BMname As String
Dim BRcode As String


Dim mailCount As Integer
mailCount = 0

For Each objfile In objfolder.Files
    
    'MsgBox objfile.Name & "---------" & objfile.Path
    
    If objfile.Name Like "*SOP Masterlist* for *" Then
        toStr = ""
        ccStr = ""
        subjectStr = ""
        bodyStr = ""
        BMname = ""
        nameStr = ""
        zipCount = zipCount + 1
        
        If objfile.Name Like "SOP Masterlist* for *" Then
            monStr = Mid(objfile.Name, 5, 3)
            yearStr = Mid(objfile.Name, 9, 4)
            Set sht = ThisWorkbook.Worksheets("Control Sheet")
            
            BRcode = Mid(objfile.Name, InStr(1, objfile.Name, "for ") + 4, InStr(1, objfile.Name, ".") - InStr(1, objfile.Name, "for ") - 4)
    
            For i = 3 To sht.UsedRange.Rows.Count
                If sht.Cells(i, 1).Value = BRcode Then
                    toStr = sht.Cells(i, 3).Value
                    BMname = sht.Cells(i, 4).Value
                End If
            Next
            subjectStr = ThisWorkbook.Worksheets("Control Sheet").Range("EmailSubj")
            bodyStr = sht.Range("I18").Value
            bodyStr = Replace(bodyStr, "#####", BMname)
            bodyStr = Replace(bodyStr, "MMMYY", sht.Range("I12").Value + Right(sht.Range("I11").Value, 2))

         End If
            If toStr <> "" Then
                Set outApp = New Outlook.Application
                Set OutMail = outApp.CreateItem(olMailItem)
            
                On Error Resume Next
                With OutMail
                    .To = toStr
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

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
