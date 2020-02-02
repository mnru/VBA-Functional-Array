Attribute VB_Name = "modCheck"
Sub checkTbl(Optional tbln = "check")
    Set varDic = CreateObject("Scripting.Dictionary")
    Dim num As Long
    Dim numFn As Long
    Dim numVar As Long
    Dim numAct As Long
    Dim x As String
    numFn = clmNum("function", tbln)
    numVar = clmNum("variable", tbln)
    numAct = clmNum("actual", tbln)
    rnum = Range(tbln).Rows.Count
    rws = rangeToArys(Range(tbln))
    Dim i As Long, j As Long
    For i = 1 To rnum
        rw = rws(i)
        cnum = lenAry(rw)
        For j = numFn + 1 To cnum
            x = getAryAt(rw, j)
            If TypeName(x) = "String" Then
                Select Case getHeadNum(x)
                    Case 1
                        y = getVar(x)
                        Call setAryAt(rw, j, varDic(y))
                    Case 2
                        y = getVar(x)
                        Call setAryAt(rw, j, varDic(y), , True)
                    Case Is > 2
                        y = Right(x, Len(x) - 2)
                        Call setAryAt(rw, j, y)
                    Case Else
                End Select
            End If
        Next j
        x = rw(numVar)
        If getHeadNum(x) = 2 Then
            Set vl = evalTblRow(rw, numFn)
            Set varDic(getVar(x)) = vl
        Else
            vl = evalTblRow(rw, numFn)
            If getHeadNum(x) = 1 Then
                varDic(getVar(x)) = vl
            End If
        End If
        Range(tbln & "[" & "actual" & "]")(i, 1) = toString(vl, , , , , True)
    Next i
End Sub

Function getHeadNum(str, Optional chr = "_")
    Dim cnt As Long
    cnt = 0
    For i = 1 To Len(str)
        If Mid(str, i, 1) = chr Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next i
    getHeadNum = cnt
End Function

Sub tsetHeaderNum()
    st = "__qqqq"
    x = getHeadUnderberNum(st)
    Debug.Print x
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
    Dim ary1, ary2
    ary1 = dropAry(ary, fncl - 1)
    ary2 = dropWhile("info", ary1, -1, "isEmpty")
    ret = evalA(ary2)
    evalTblRow = ret
End Function

Function getVar(str As String) As String
    Dim str0 As String
    Dim str1 As String
    Dim n As Long
    n = getHeadNum(str)
    str0 = Right(str, Len(str) - n)
    str1 = Right(str, Len(str) - 2)
    getVar = IIf(n > 2, str1, str0)
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

Sub addBtn(rn, mn, Optional cn = "run", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ThisWorkbook.Name
    Set rg = Workbooks(bn).Sheets(sn).Range(rn)
    Set btn = Workbooks(bn).Sheets(sn).Buttons.Add(rg.Left, rg.Top, rg.width, rg.Height)
    btn.OnAction = mn
    btn.Caption = cn
End Sub
