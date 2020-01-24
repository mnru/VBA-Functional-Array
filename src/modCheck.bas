Attribute VB_Name = "modCheck"
Sub checkTbl(Optional tbln = "check")
    
    Set vardic = CreateObject("Scripting.Dictionary")
    Dim num
    
    numFn = clmNum("function", tbln)
    numVar = clmNum("variable", tbln)
    numAct = clmNum("actual", tbln)
    
    rnum = Range(tbln).Rows.Count
    
    rws = rangeToArys(Range(tbln))
    
    For i = 1 To rnum
        rw = rws(i)
        cnum = lenAry(rw)
        For j = numFn + 1 To cnum
            x = getAryAt(rw, j)
            If TypeName(x) = "String" Then
                If Left(x, 1) = "_" Then
                    y = getVar(x)
                    Call setAryAt(rw, j, vardic(y))
                End If
            End If
        Next j
        vl = evalTblRow(rw, numFn)
        x = rw(numVar)
        If Not IsEmpty(x) And x <> "" Then
            vardic(getVar(x)) = vl
        End If
        Range(tbln & "[" & "actual" & "]")(i, 1) = toString(vl, , , , , True)
        
    Next i
    

End Sub

Function clmNum(clmnn, tbln, Optional bn = "")
    abn = ActiveWorkbook.Name
    If bn = "" Then bn = abn
    Workbooks(bn).Activate
    Dim ret
    With Application.WorksheetFunction
        ret = .Match(clmnn, Range(tbln & "[#headers]"), False)
    End With
    clmNum = ret
    Workbooks(abn).Activate
    
End Function

Function evalTblRow(ary, fncl)
    Dim ret
    Dim ary1, ary2, ary3
 ' ary1 = rangeToAry(lrow)
    ary2 = dropAry(ary, fncl - 1)
    ary3 = dropWhile("info", ary2, -1, "isEmpty")
    ret = evalA(ary3)
    evalTblRow = ret
    
End Function

Function getVar(str)
    Dim ret
    ret = str
    If Left(str, 1) = "_" Then ret = Right(str, Len(str) - 1)
    getVar = ret
End Function


Sub testCheckTbl()
    checkTbl
End Sub

Sub clearActual(Optional tbln = "check")
    Range(tbln & "[actual]").ClearContents
End Sub

Sub setbtn()
    Call addBtn("D1", "checkTbl", "eval")
  Call addBtn("F1", "clearActual", "clear")

End Sub
