Attribute VB_Name = "calculate"

Public  Sub gochila()
    'MRGFX手數
    Application.ScreenUpdating = False
    Call tala
    Sheets("OA").Select
     Cells.ClearContents
     Cells.ClearContents
    Range("A1").Select
    ActiveSheet.PasteSpecial Format:="HTML", NoHTMLFormatting:=True
    
        If Cells(1, 1) = "BIBFXOA" Then
            Sheets("Data").Cells(4, 30) = "BIBFXOA"
        Else
            Sheets("Data").Cells(4, 30) = ""
        End If
     
    Columns("A").Select
    Selection.Delete Shift:=xlToLeft
     vito = WorksheetFunction.Match("交易账号", Columns("a"), 0) - 1
    Range("A" & vito, Range("a1").End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
        
    If Sheets("Data").Cells(4, 30) = "BIBFXOA" Then GoTo bigman Else GoTo smallman
bigman:
    X = WorksheetFunction.CountA(Columns("a"))
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:I" & X).AutoFilter Field:=1, Criteria1:=">988000000", Operator:=xlAnd
    Range("N2").Select
    Selection.Copy
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
    
    Range("A1:C1", Range("A1").End(xlDown)).Copy
    Sheets("Data").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("OA").Select
             
    Range("E1", Range("E1").End(xlDown)).Copy
    Sheets("Data").Select
    Range("D1").Select
    ActiveSheet.Paste
    Sheets("OA").Select
             
    Range("H1", Range("H1").End(xlDown)).Copy
    Sheets("Data").Select
    Range("E1").Select
    ActiveSheet.Paste
    y = WorksheetFunction.CountA(Columns("a"))
    GoTo vito
    
smallman:
    Range("N2").Select
    Selection.Copy
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
     y = WorksheetFunction.CountA(Columns("a"))
    Range("a2:c" & y & ",e2:e" & y & ",h2:h" & y).Copy
    Sheets("Data").Select
    Range("a2").Select
    ActiveSheet.Paste
    GoTo vito
    
vito:
    Range("f2") = "=RC[-5]&RC[-4]"
    Range("g2") = "=RC[-3]"
    Range("h2") = "=RC[-3]"
    Range("i2") = "=IFERROR(RC[-2]=VLOOKUP(RC[-3],C[11]:C[12],2,0),"OA多生成")"
    Range("j2") = "=IFERROR(RC[-2]=VLOOKUP(RC[-4],C[10]:C[12],3,0),"OA多生成")"
    Range("f2:j2").Copy
    Range("f2:j" & y).Select
    ActiveSheet.Paste
    Application.ScreenUpdating = True
End Sub

Public  Sub reona()
    'ttr
    Application.ScreenUpdating = False
    Call tara
    Sheets("Sheet1").Select
     Cells.ClearContents
    Sheets("TTR").Select
     Cells.ClearContents
     Cells.ClearContents
    Range("A1").Select
    ActiveSheet.PasteSpecial Format:="Unicode文本", NoHTMLFormatting:=True
    Columns("S:S").Select
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByRows
    WorkDay = Application.Evaluate("WEEKDAY(TODAY())-1")
     tday = Date  '当天上班算表日
     Cells(1, 19) = tday - 1 + TimeValue("22:0:0")
     If WorkDay = 2 Then
        Cells(2, 19) = Cells(1, 19) - 2
     Else
        Cells(2, 19) = Cells(1, 19) - 1
     End If
     X = Cells(1, 19)
     y = Cells(2, 19)
     eve = WorksheetFunction.CountA(Columns("A")) + 1
    Rows("4:4").Select
     Selection.AutoFilter
     ActiveSheet.Range("A4:AJ" & eve).AutoFilter Field:=19, Criteria1:=">=" & y, Operator:=xlAnd, Criteria2:="<" & X
     ActiveSheet.Range("A4:AJ" & eve).AutoFilter Field:=14, Criteria1:="OUT"
    Range("a4", Range("a4").End(xlDown)).Copy
    Sheets("sheet1").Select
    Range("a1").Select
     ActiveSheet.Paste
    Sheets("TTR").Select
    Range("o4", Range("o4").End(xlDown)).Copy
    Sheets("sheet1").Select
    Range("ac1").Select
     ActiveSheet.Paste
     ActiveSheet.Range("a:a").RemoveDuplicates Columns:=1, Header:=xlNo
     ActiveSheet.Range("ac:ac").RemoveDuplicates Columns:=1, Header:=xlNo
    '---------------------
    X = 2
    y = 2
    Do While Cells(X, 29) <> ""
        Cells(1, y) = Cells(X, 29)
        X = X + 1
        y = y + 1
    Loop
    
    Z = 0
    Dim arr As Variant
    arr = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q")
    Do While Cells(1, Z + 2) <> ""
        Cells(2, Z + 2) = "=round(SUMIFS(TTR!$Q:$Q,TTR!$O:$O,$" & arr(Z) & "$1,TTR!$A:$A,A2,TTR!$S:$S,"">=""&TTR!$S$2,TTR!$S:$S,""<""&TTR!$S$1,TTR!$N:$N,""out""),2)"
        Z = Z + 1
    Loop
    vito = WorksheetFunction.CountA(Columns("a"))
    Range("b2", Range("b2").End(xlToRight)).Copy
    Range("b2: " & arr(Z - 1) & vito).Select
    ActiveSheet.Paste
    
    '------------------------
    aa = 2
    For i = 2 To vito
        For j = 2 To Z + 1
            If Cells(i, j) <> 0 Then
                Sheets("Data").Cells(aa, 13) = Cells(i, 1)
                Sheets("Data").Cells(aa, 14) = "0"
                Sheets("Data").Cells(aa, 16) = Cells(i, j)
                Sheets("Data").Cells(aa, 19) = Cells(1, j)
                aa = aa + 1
            End If
        Next j
    Next i
    Sheets("Data").Select
    Call formuuuu
    Application.ScreenUpdating = True
End Sub

Public  Sub twt()
    '计算手數模板
    Range("M1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Call formuuuu
End Sub

Public  Sub tala()
    Sheets("Data").Select
    Range("A2:J2", Range("A2").End(xlDown)).Select
    Selection.ClearContents
    Range("z16", Range("z16").End(xlDown)).Select
    Selection.ClearContents
End Sub

Public  Sub tara()
    Sheets("Data").Select
    Range("M2:X2", Range("M2").End(xlDown)).Select
    Selection.ClearContents
    Range("z16", Range("z16").End(xlDown)).Select
    Selection.ClearContents
End Sub

Public  Sub formuuuu()
    Range("T2") = "=RC[-7]&RC[-1]"
    Range("U2") = "=RC[-7]"
    Range("V2") = "=RC[-6]"
    Range("W2") = "=IFERROR(RC[-2]=VLOOKUP(RC[-3],C[-17]:C[-16],2,0),"OA少生成")"
    Range("X2") = "=IFERROR(RC[-2]=VLOOKUP(RC[-4],C[-18]:C[-16],3,0),"OA少生成")"
     vito = WorksheetFunction.CountA(Columns("M"))
    Range("T2:X2").Copy
    Range("T2:X" & vito).Select
    ActiveSheet.Paste
End Sub

Public  Sub takala()
    Range("a2:x100000").Select
    Selection.ClearContents
    
    Range("z16", Range("z16").End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("OA").Select
    Cells.ClearContents
    Cells.ClearContents
    
    Sheets("TTR").Select
    Cells.ClearContents
    Cells.ClearContents
    
    Sheets("Sheet1").Select
    Cells.ClearContents
    Sheets("Data").Select
End Sub