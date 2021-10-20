Attribute VB_Name = "check_lots"

Public  Sub check_lots()
    Range("z16", Range("z16").End(xlDown)).Select
    Selection.ClearFormats
    Selection.ClearContents

    ABU = WorksheetFunction.CountIf(Range("I2:I10000"), "FALSE")
    ACU = WorksheetFunction.CountIf(Range("J2:J10000"), "FALSE")
    ADU = WorksheetFunction.CountIf(Range("I2:I10000"), "OA多生成")
    BDU = WorksheetFunction.CountIf(Range("W2:W10000"), "OA少生成")
     
    If ADU > 0 Then
    k = WorksheetFunction.CountA(Range("Z4:Z10000")) + 10
     X = 2
     Cells(k, 26) = "OA多生成"
     Cells(k, 26).Font.Bold = True
     'Selection.Font.Bold = False
     k = k + 1
        Do While Cells(X, 1) <> ""
            If Cells(X, 9) = "OA多生成" Then
                Cells(k, 26) = Cells(X, 1) & "   " & Cells(X, 2) & "   " & Cells(X, 3) & "   活动手数" & WorksheetFunction.Text(Cells(X, 4), "0.00") & "   总手数" & WorksheetFunction.Text(Cells(X, 5), "0.00")
                k = k + 1
            End If
            X = X + 1
        Loop
    End If
    
    If BDU > 0 Then
     k = WorksheetFunction.CountA(Range("Z4:Z10000")) + 10
     X = 2
     Cells(k, 26) = "OA少生成"
     Cells(k, 26).Font.Bold = True
     'Selection.Font.Bold = False
     k = k + 1
        Do While Cells(X, 1) <> ""
            If Cells(X, 23) = "OA少生成" Then
                Cells(k, 26) = Cells(X, 13) & "   " & Cells(X, 19) & "   活动手数" & WorksheetFunction.Text(Cells(X, 14), "0.00") & "   总手数" & WorksheetFunction.Text(Cells(X, 16), "0.00")
                k = k + 1
            End If
            X = X + 1
        Loop
    End If
    
    If ABU > 0 Or ACU > 0 Then
     k = WorksheetFunction.CountA(Range("Z4:Z10000")) + 10
     X = 2
     Cells(k, 26) = "OA手數"
     Cells(k, 26).Font.Bold = True
     'Selection.Font.Bold = False
     k = k + 1
        Do While Cells(X, 1) <> ""
            If Cells(X, 9) = "False" Or Cells(X, 10) = "False" Then
                Cells(k, 26) = Cells(X, 1) & "   " & Cells(X, 2) & "   " & Cells(X, 3) & "   活动手数" & WorksheetFunction.Text(Cells(X, 4), "0.00") & "   总手数" & WorksheetFunction.Text(Cells(X, 5), "0.00")
                k = k + 1
            End If
            X = X + 1
        Loop
    End If
    
    If ABU > 0 Or ACU > 0 Then
     k = WorksheetFunction.CountA(Range("Z4:Z10000")) + 10
     X = 2
     Cells(k, 26) = "计算手數"
     Cells(k, 26).Font.Bold = True
     'Selection.Font.Bold = False
     k = k + 1
        Do While Cells(X, 1) <> ""
            If Cells(X, 9) = "False" Or Cells(X, 10) = "False" Then
                vito = WorksheetFunction.Match(Cells(X, 6), Columns("T"), 0)
                
                Cells(k, 26) = Cells(vito, 13) & "   " & Cells(vito, 19) & "   " & Cells(X, 3) & "   活动手数 " & WorksheetFunction.Text(Cells(vito, 14), "0.00") & "   总手数" & WorksheetFunction.Text(Cells(vito, 16), "0.00")
                k = k + 1
            End If
            X = X + 1
        Loop
    End If
End Sub