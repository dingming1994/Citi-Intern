Sub FillContractedData()
Set outbk = ActiveWorkbook
Set pbsht = outbk.Sheets("PB")
Set rmsht = outbk.Sheets("RM")
Set cpcsht = outbk.Sheets("CPC")
Dim pbRevBk, rmRevBk, CreBk As Workbook
pbrevpath = "I:\CAP_Profile_PRD65\Desktop\ABU Ranking System\CB_Contracted_Product_Revenue_SEP15(Final).xlsx"

rmrevpath = "I:\CAP_Profile_PRD65\Desktop\ABU Ranking System\CGG3_Contracted_Product_Revenue_SEP15(Final).xlsx"

crepath = "I:\CAP_Profile_PRD65\Desktop\ABU Ranking System\CGCB_Contracted_Product_Sales_Portfolio_Credits_SEP15(Final).xlsx"

monthstr = "Sep15"
Set pbRevBk = Workbooks.Open(pbrevpath, False)
For Each sht In pbRevBk.Worksheets
    If sht.Name Like "*Contracted_Product_Revenue*" Then
        Set datasht = sht
    End If
Next

curMonth = Month("1 " & Left(monthstr, 3))
ebrcol = 25 + curMonth
revcol = 38 + curMonth
        
For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To pbsht.Range("B3").End(xlDown).Row
        If nm = pbsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        pbsht.Cells(idx, ebrcol) = datasht.Cells(i, 29)
        pbsht.Cells(idx, revcol) = datasht.Cells(i, 32)
    Else
        rowct = pbsht.Range("B2").End(xlDown).Row + 1
        pbsht.Rows(rowct).Insert
        pbsht.Cells(rowct, 2) = nm
        pbsht.Cells(rowct, ebrcol) = datasht.Cells(i, 29)
        pbsht.Cells(rowct, revcol) = datasht.Cells(i, 32)
    End If
    
Next

' Fill in 0 value
For i = 2 To pbsht.Range("B3").End(xlDown).Row
    If pbsht.Cells(i, ebrcol) = "" Then
        pbsht.Cells(i, ebrcol) = 0
    End If
    If pbsht.Cells(i, revcol) = "" Then
        pbsht.Cells(i, revcol) = 0
    End If
Next

Application.DisplayAlerts = False
pbRevBk.Close
Application.DisplayAlerts = True


Set rmRevBk = Workbooks.Open(rmrevpath, False)
For Each sht In rmRevBk.Worksheets
    If sht.Name Like "*Contracted_Product_Revenue*" Then
        Set datasht = sht
    End If
Next

curMonth = Month("1 " & Left(monthstr, 3))
ebrcol = 25 + curMonth
revcol = 38 + curMonth
 'RM
For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To rmsht.Range("B3").End(xlDown).Row
        If nm = rmsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        rmsht.Cells(idx, ebrcol) = datasht.Cells(i, 31)
        rmsht.Cells(idx, revcol) = datasht.Cells(i, 34)
    ElseIf datasht.Cells(i, 4) <> "Y" Then
        rowct = rmsht.Range("B2").End(xlDown).Row + 1
        rmsht.Rows(rowct).Insert
        rmsht.Cells(rowct, 2) = nm
        rmsht.Cells(rowct, ebrcol) = datasht.Cells(i, 31)
        rmsht.Cells(rowct, revcol) = datasht.Cells(i, 34)
    End If
    
Next

' Fill in 0 value
For i = 2 To rmsht.Range("B3").End(xlDown).Row
    If rmsht.Cells(i, ebrcol) = "" Then
        rmsht.Cells(i, ebrcol) = 0
    End If
    If rmsht.Cells(i, revcol) = "" Then
        rmsht.Cells(i, revcol) = 0
    End If
Next

'CPC
For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To cpcsht.Range("B3").End(xlDown).Row
        If nm = cpcsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        cpcsht.Cells(idx, ebrcol) = datasht.Cells(i, 31)
        cpcsht.Cells(idx, revcol) = datasht.Cells(i, 34)
    ElseIf datasht.Cells(i, 4) = "Y" And datasht.Cells(i, 7) <> 999999 Then
        rowct = cpcsht.Range("B2").End(xlDown).Row + 1
        cpcsht.Rows(rowct).Insert
        cpcsht.Cells(rowct, 2) = nm
        cpcsht.Cells(rowct, ebrcol) = datasht.Cells(i, 31)
        cpcsht.Cells(rowct, revcol) = datasht.Cells(i, 34)
    End If
    
Next

' Fill in 0 value
For i = 2 To cpcsht.Range("B3").End(xlDown).Row
    If cpcsht.Cells(i, ebrcol) = "" Then
        cpcsht.Cells(i, ebrcol) = 0
    End If
    If cpcsht.Cells(i, revcol) = "" Then
        cpcsht.Cells(i, revcol) = 0
    End If
Next

Application.DisplayAlerts = False
rmRevBk.Close
Application.DisplayAlerts = True



'PB ABU
Set CreBk = Workbooks.Open(crepath, False)
For Each sht In CreBk.Worksheets
    If sht.Name Like "*CB_CREDITS_Summary*" Then
        Set datasht = sht
    End If
Next

curMonth = Month("1 " & Left(monthstr, 3))
abucol = 4 + curMonth


For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To pbsht.Range("B3").End(xlDown).Row
        If nm = pbsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        pbsht.Cells(idx, abucol) = datasht.Cells(i, 51)
    End If
Next

'RM
For Each sht In CreBk.Worksheets
    If sht.Name Like "*CG_CREDITS_Summary*" Then
        Set datasht = sht
    End If
Next

curMonth = Month("1 " & Left(monthstr, 3))
abucol = 4 + curMonth


For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To rmsht.Range("B3").End(xlDown).Row
        If nm = rmsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        rmsht.Cells(idx, abucol) = datasht.Cells(i, 55)
    End If
Next

'cpc
For Each sht In CreBk.Worksheets
    If sht.Name Like "*CPC_ABU_Ranking*" Then
        Set datasht = sht
    End If
Next

curMonth = Month("1 " & Left(monthstr, 3))
abucol = 4 + curMonth


For i = 2 To datasht.Range("B2").End(xlDown).Row
    nm = datasht.Cells(i, 2)
    idx = 0
    For j = 3 To cpcsht.Range("B3").End(xlDown).Row
        If nm = cpcsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 Then
        cpcsht.Cells(idx, abucol) = datasht.Cells(i, 58)
    End If
Next



Application.DisplayAlerts = False
CreBk.Close
Application.DisplayAlerts = True


End Sub


Sub UpdateSummary()
UpdateSummary_Ranked
UpdateSummary_Resign
UpdateSummary_Ranked_Zone
UpdateSummary_Resign_Zone


End Sub


Sub UpdateSummary_Ranked()
Set outbk = ActiveWorkbook
Set pbsht = outbk.Sheets("PB")
Set rmsht = outbk.Sheets("RM")
Set cpcsht = outbk.Sheets("CPC")
Set sumsht = outbk.Sheets("Summary")
Dim totRow As Integer



'find total row idx
totRow = 0
For i = 7 To sumsht.UsedRange.Rows.Count
    With sumsht
        If .Cells(i, 1) = "" And .Cells(i + 1, 1) = "" And .Cells(i, 2) = "" And .Cells(i + 1, 2) = "" And totRow = 0 Then
            totRow = i
        End If
    End With
Next

'clear data - to do
sumsht.Range(sumsht.Cells(8, 3), sumsht.Cells(totRow, 31)).ClearContents
    
'PB

For i = 3 To pbsht.Range("B3").End(xlDown).Row
    br = pbsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And pbsht.Cells(i, 21) <> "Resigned" And pbsht.Cells(i, 21) <> "Transferred" And Not pbsht.Cells(i, 21) Like "Promoted to RM" Then
        sumsht.Cells(idx, 3) = sumsht.Cells(idx, 3) + 1
        sumsht.Cells(idx, 4) = sumsht.Cells(idx, 4) + pbsht.Cells(i, 25)
        sumsht.Cells(idx, 5) = sumsht.Cells(idx, 5) + pbsht.Cells(i, 22)
        sumsht.Cells(idx, 6) = sumsht.Cells(idx, 6) + pbsht.Cells(i, 23)
        sumsht.Cells(idx, 7) = sumsht.Cells(idx, 7) + pbsht.Cells(i, 24)
        sumsht.Cells(idx, 8) = sumsht.Cells(idx, 8) + pbsht.Cells(i, 38)
        sumsht.Cells(idx, 10) = sumsht.Cells(idx, 10) + pbsht.Cells(i, 51)
    End If
Next
For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 3) > 0 Then
        sumsht.Cells(i, 4) = sumsht.Cells(i, 4) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 5) = sumsht.Cells(i, 5) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 6) = sumsht.Cells(i, 6) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 7) = sumsht.Cells(i, 7) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 9) = sumsht.Cells(i, 8) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 11) = sumsht.Cells(i, 10) / sumsht.Cells(i, 3)
    End If

Next

sumsht.Cells(totRow, 3).Formula = "=sum(C8:C" & CStr(totRow - 1) & ")"
For i = 4 To 11
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(8, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next
st = 3
For i = 8 To totRow - 1
    If sumsht.Cells(i, 1) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
    
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'RM

For i = 3 To rmsht.Range("B3").End(xlDown).Row
    br = rmsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And rmsht.Cells(i, 21) <> "Resigned" And rmsht.Cells(i, 21) <> "Transferred" And Not rmsht.Cells(i, 21) Like "Promoted to CPC" And Not rmsht.Cells(i, 21) Like "Promoted to BM" Then
        sumsht.Cells(idx, 13) = sumsht.Cells(idx, 13) + 1
        sumsht.Cells(idx, 14) = sumsht.Cells(idx, 14) + rmsht.Cells(i, 25)
        sumsht.Cells(idx, 15) = sumsht.Cells(idx, 15) + rmsht.Cells(i, 22)
        sumsht.Cells(idx, 16) = sumsht.Cells(idx, 16) + rmsht.Cells(i, 23)
        sumsht.Cells(idx, 17) = sumsht.Cells(idx, 17) + rmsht.Cells(i, 24)
        sumsht.Cells(idx, 18) = sumsht.Cells(idx, 18) + rmsht.Cells(i, 38)
        sumsht.Cells(idx, 20) = sumsht.Cells(idx, 20) + rmsht.Cells(i, 51)
    End If
Next
For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 13) > 0 Then
        sumsht.Cells(i, 14) = sumsht.Cells(i, 14) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 15) = sumsht.Cells(i, 15) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 16) = sumsht.Cells(i, 16) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 17) = sumsht.Cells(i, 17) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 19) = sumsht.Cells(i, 18) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 21) = sumsht.Cells(i, 20) / sumsht.Cells(i, 13)
    End If
Next

sumsht.Cells(totRow, 13).Formula = "=sum(M8:M" & CStr(totRow - 1) & ")"
For i = 14 To 21
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(8, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next


st = 13
For i = 8 To totRow - 1
     If sumsht.Cells(i, 1) <> "" Then
     sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
   
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

'CPC

For i = 3 To cpcsht.Range("B3").End(xlDown).Row
    br = cpcsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And cpcsht.Cells(i, 21) <> "Resigned" And cpcsht.Cells(i, 21) <> "Transferred" And Not cpcsht.Cells(i, 21) Like "Promoted to BM" Then
        sumsht.Cells(idx, 23) = sumsht.Cells(idx, 23) + 1
        sumsht.Cells(idx, 24) = sumsht.Cells(idx, 24) + cpcsht.Cells(i, 25)
        sumsht.Cells(idx, 25) = sumsht.Cells(idx, 25) + cpcsht.Cells(i, 22)
        sumsht.Cells(idx, 26) = sumsht.Cells(idx, 26) + cpcsht.Cells(i, 23)
        sumsht.Cells(idx, 27) = sumsht.Cells(idx, 27) + cpcsht.Cells(i, 24)
        sumsht.Cells(idx, 28) = sumsht.Cells(idx, 28) + cpcsht.Cells(i, 38)
        sumsht.Cells(idx, 30) = sumsht.Cells(idx, 30) + cpcsht.Cells(i, 51)
    End If
Next
For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 23) > 0 Then
        sumsht.Cells(i, 24) = sumsht.Cells(i, 24) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 25) = sumsht.Cells(i, 25) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 26) = sumsht.Cells(i, 26) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 27) = sumsht.Cells(i, 27) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 29) = sumsht.Cells(i, 28) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 31) = sumsht.Cells(i, 30) / sumsht.Cells(i, 23)
    End If
Next

sumsht.Cells(totRow, 23).Formula = "=sum(W8:W" & CStr(totRow - 1) & ")"
For i = 24 To 31
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(8, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next

st = 23
For i = 8 To totRow - 1
    If sumsht.Cells(i, 1) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

End Sub


Sub UpdateSummary_Resign()
Set outbk = ActiveWorkbook
Set pbsht = outbk.Sheets("PB")
Set rmsht = outbk.Sheets("RM")
Set cpcsht = outbk.Sheets("CPC")
Set sumsht = outbk.Sheets("Summary")
Dim totRow As Integer

'find total row idx
totRow = 0
For i = 7 To sumsht.UsedRange.Rows.Count
    With sumsht
        If .Cells(i, 33) = "" And .Cells(i + 1, 33) = "" And .Cells(i, 34) = "" And .Cells(i + 1, 34) = "" And totRow = 0 Then
            totRow = i
        End If
    End With
Next

'clear data - to do
sumsht.Range(sumsht.Cells(8, 35), sumsht.Cells(totRow, 48)).ClearContents
    
'PB
st = 35
For i = 3 To pbsht.Range("B3").End(xlDown).Row
    br = pbsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (pbsht.Cells(i, 21) = "Resigned" Or pbsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + pbsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + pbsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + pbsht.Cells(i, 24)
    End If
Next

For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AI8:AI" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(8, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 8 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'RM
st = 40
For i = 3 To rmsht.Range("B3").End(xlDown).Row
    br = rmsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (rmsht.Cells(i, 21) = "Resigned" Or rmsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + rmsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + rmsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + rmsht.Cells(i, 24)
    End If
Next

For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AN8:AN" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(8, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 8 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'CPC
st = 45
For i = 3 To cpcsht.Range("B3").End(xlDown).Row
    br = cpcsht.Cells(i, 3)
    idx = 0
    For j = 8 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (cpcsht.Cells(i, 21) = "Resigned" Or cpcsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + cpcsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + cpcsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + cpcsht.Cells(i, 24)
    End If
Next

For i = 8 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) <> "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AS8:AS" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(8, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 8 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

End Sub


Sub UpdateSummary_Ranked_Zone()
Set outbk = ActiveWorkbook
Set pbsht = outbk.Sheets("PB")
Set rmsht = outbk.Sheets("RM")
Set cpcsht = outbk.Sheets("CPC")
Set sumsht = outbk.Sheets("Summary")
Dim totRow As Integer



'find total row idx
totRow = 0
For i = 31 To sumsht.UsedRange.Rows.Count
    With sumsht
        If .Cells(i, 2) = "" And .Cells(i + 1, 2) = "" And totRow = 0 Then
            totRow = i
        End If
    End With
Next

'clear data - to do
sumsht.Range(sumsht.Cells(32, 3), sumsht.Cells(totRow, 31)).ClearContents
    
'PB
st = 3

For i = 3 To pbsht.Range("B3").End(xlDown).Row
    br = pbsht.Cells(i, 4)
    idx = 0
    For j = 31 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And pbsht.Cells(i, 21) <> "Resigned" And pbsht.Cells(i, 21) <> "Transferred" And Not pbsht.Cells(i, 21) Like "Promoted to RM" Then
        sumsht.Cells(idx, 3) = sumsht.Cells(idx, 3) + 1
        sumsht.Cells(idx, 4) = sumsht.Cells(idx, 4) + pbsht.Cells(i, 25)
        sumsht.Cells(idx, 5) = sumsht.Cells(idx, 5) + pbsht.Cells(i, 22)
        sumsht.Cells(idx, 6) = sumsht.Cells(idx, 6) + pbsht.Cells(i, 23)
        sumsht.Cells(idx, 7) = sumsht.Cells(idx, 7) + pbsht.Cells(i, 24)
        sumsht.Cells(idx, 8) = sumsht.Cells(idx, 8) + pbsht.Cells(i, 38)
        sumsht.Cells(idx, 10) = sumsht.Cells(idx, 10) + pbsht.Cells(i, 51)
    End If
Next
For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 3) > 0 Then
        sumsht.Cells(i, 4) = sumsht.Cells(i, 4) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 5) = sumsht.Cells(i, 5) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 6) = sumsht.Cells(i, 6) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 7) = sumsht.Cells(i, 7) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 9) = sumsht.Cells(i, 8) / sumsht.Cells(i, 3)
        sumsht.Cells(i, 11) = sumsht.Cells(i, 10) / sumsht.Cells(i, 3)
    End If

Next

sumsht.Cells(totRow, 3).Formula = "=sum(C32:C" & CStr(totRow - 1) & ")"
For i = 4 To 11
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(32, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
    
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

'RM
st = 13

For i = 3 To rmsht.Range("B3").End(xlDown).Row
    br = rmsht.Cells(i, 4)
    idx = 0
    For j = 31 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And rmsht.Cells(i, 21) <> "Resigned" And rmsht.Cells(i, 21) <> "Transferred" And Not rmsht.Cells(i, 21) Like "Promoted to BM" Then
        sumsht.Cells(idx, 13) = sumsht.Cells(idx, 13) + 1
        sumsht.Cells(idx, 14) = sumsht.Cells(idx, 14) + rmsht.Cells(i, 25)
        sumsht.Cells(idx, 15) = sumsht.Cells(idx, 15) + rmsht.Cells(i, 22)
        sumsht.Cells(idx, 16) = sumsht.Cells(idx, 16) + rmsht.Cells(i, 23)
        sumsht.Cells(idx, 17) = sumsht.Cells(idx, 17) + rmsht.Cells(i, 24)
        sumsht.Cells(idx, 18) = sumsht.Cells(idx, 18) + rmsht.Cells(i, 38)
        sumsht.Cells(idx, 20) = sumsht.Cells(idx, 20) + rmsht.Cells(i, 51)
    End If
Next
For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 3) > 0 Then
        sumsht.Cells(i, 14) = sumsht.Cells(i, 14) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 15) = sumsht.Cells(i, 15) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 16) = sumsht.Cells(i, 16) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 17) = sumsht.Cells(i, 17) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 19) = sumsht.Cells(i, 18) / sumsht.Cells(i, 13)
        sumsht.Cells(i, 21) = sumsht.Cells(i, 20) / sumsht.Cells(i, 13)
    End If

Next

sumsht.Cells(totRow, 13).Formula = "=sum(M32:M" & CStr(totRow - 1) & ")"
For i = 14 To 21
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(32, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
    
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'CPC
st = 23

For i = 3 To cpcsht.Range("B3").End(xlDown).Row
    br = cpcsht.Cells(i, 4)
    idx = 0
    For j = 31 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And cpcsht.Cells(i, 21) <> "Resigned" And cpcsht.Cells(i, 21) <> "Transferred" And Not cpcsht.Cells(i, 21) Like "Promoted to BM" Then
        sumsht.Cells(idx, 23) = sumsht.Cells(idx, 23) + 1
        sumsht.Cells(idx, 24) = sumsht.Cells(idx, 24) + cpcsht.Cells(i, 25)
        sumsht.Cells(idx, 25) = sumsht.Cells(idx, 25) + cpcsht.Cells(i, 22)
        sumsht.Cells(idx, 26) = sumsht.Cells(idx, 26) + cpcsht.Cells(i, 23)
        sumsht.Cells(idx, 27) = sumsht.Cells(idx, 27) + cpcsht.Cells(i, 24)
        sumsht.Cells(idx, 28) = sumsht.Cells(idx, 28) + cpcsht.Cells(i, 38)
        sumsht.Cells(idx, 30) = sumsht.Cells(idx, 30) + cpcsht.Cells(i, 51)
    End If
Next
For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, 3) > 0 Then
        sumsht.Cells(i, 24) = sumsht.Cells(i, 24) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 25) = sumsht.Cells(i, 25) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 26) = sumsht.Cells(i, 26) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 27) = sumsht.Cells(i, 27) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 29) = sumsht.Cells(i, 28) / sumsht.Cells(i, 23)
        sumsht.Cells(i, 31) = sumsht.Cells(i, 30) / sumsht.Cells(i, 23)
    End If

Next

sumsht.Cells(totRow, 23).Formula = "=sum(W32:W" & CStr(totRow - 1) & ")"
For i = 24 To 31
    sumsht.Cells(totRow, i).Formula = "=average(" & sumsht.Cells(32, i).Address & ":" & sumsht.Cells(totRow - 1, i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.ColorIndex = xlNone
    
        If sumsht.Cells(i, st) > 0 Then
            If sumsht.Cells(i, st + 2) >= sumsht.Cells(totRow, st + 2) Then
                sumsht.Cells(i, st + 2).Interior.Color = vbGreen
            End If
            If sumsht.Cells(i, st + 4) >= sumsht.Cells(totRow, st + 4) Then
                sumsht.Cells(i, st + 4).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 6) < sumsht.Cells(totRow, st + 6) Then
                sumsht.Cells(i, st + 6).Interior.Color = vbRed
            End If
            If sumsht.Cells(i, st + 8) < sumsht.Cells(totRow, st + 8) Then
                sumsht.Cells(i, st + 8).Interior.Color = vbRed
            End If
        Else
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 8)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

End Sub


Sub UpdateSummary_Resign_Zone()
Set outbk = ActiveWorkbook
Set pbsht = outbk.Sheets("PB")
Set rmsht = outbk.Sheets("RM")
Set cpcsht = outbk.Sheets("CPC")
Set sumsht = outbk.Sheets("Summary")
Dim totRow As Integer

'find total row idx
totRow = 0
For i = 32 To sumsht.UsedRange.Rows.Count
    With sumsht
        If .Cells(i, 34) = "" And .Cells(i + 1, 34) = "" And totRow = 0 Then
            totRow = i
        End If
    End With
Next

'clear data - to do
sumsht.Range(sumsht.Cells(32, 35), sumsht.Cells(totRow, 48)).ClearContents
    
'PB
st = 35
For i = 3 To pbsht.Range("B3").End(xlDown).Row
    br = pbsht.Cells(i, 4)
    idx = 0
    For j = 32 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (pbsht.Cells(i, 21) = "Resigned" Or pbsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + pbsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + pbsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + pbsht.Cells(i, 24)
    End If
Next

For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AI32:AI" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(32, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'RM
st = 40
For i = 3 To rmsht.Range("B3").End(xlDown).Row
    br = rmsht.Cells(i, 4)
    idx = 0
    For j = 32 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (rmsht.Cells(i, 21) = "Resigned" Or rmsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + rmsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + rmsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + rmsht.Cells(i, 24)
    End If
Next

For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AN32:AN" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(32, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next


'CPC
st = 45
For i = 3 To cpcsht.Range("B3").End(xlDown).Row
    br = cpcsht.Cells(i, 4)
    idx = 0
    For j = 32 To sumsht.UsedRange.Rows.Count
        If br = sumsht.Cells(j, 2) Then
            idx = j
        End If
    Next
    
    If idx <> 0 And (cpcsht.Cells(i, 21) = "Resigned" Or cpcsht.Cells(i, 21) = "Transferred") Then
        sumsht.Cells(idx, st) = sumsht.Cells(idx, st) + 1
        sumsht.Cells(idx, st + 1) = sumsht.Cells(idx, st + 1) + cpcsht.Cells(i, 22)
        sumsht.Cells(idx, st + 2) = sumsht.Cells(idx, st + 2) + cpcsht.Cells(i, 23)
        sumsht.Cells(idx, st + 3) = sumsht.Cells(idx, st + 3) + cpcsht.Cells(i, 24)
    End If
Next

For i = 32 To sumsht.UsedRange.Rows.Count
    If sumsht.Cells(i, 1) = "" And sumsht.Cells(i, 2) <> "" And sumsht.Cells(i, st) > 0 Then
        sumsht.Cells(i, st + 1) = sumsht.Cells(i, st + 1) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 2) = sumsht.Cells(i, st + 2) / sumsht.Cells(i, st)
        sumsht.Cells(i, st + 3) = sumsht.Cells(i, st + 3) / sumsht.Cells(i, st)

    End If

Next

sumsht.Cells(totRow, st).Formula = "=sum(AS32:AS" & CStr(totRow - 1) & ")"
For i = 1 To 3
    sumsht.Cells(totRow, st + i).Formula = "=average(" & sumsht.Cells(32, st + i).Address & ":" & sumsht.Cells(totRow - 1, st + i).Address & ")"
Next

For i = 32 To totRow - 1
    If sumsht.Cells(i, 2) <> "" Then
        sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.ColorIndex = xlNone
    
        If Not sumsht.Cells(i, st) > 0 Then
          
            sumsht.Range(sumsht.Cells(i, st), sumsht.Cells(i, st + 3)).Interior.Color = rgbDarkGrey
        End If
    End If
Next

End Sub
