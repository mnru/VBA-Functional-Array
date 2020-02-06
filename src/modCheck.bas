Attribute VB_Name = "modCheck"
Sub checkTbl(Optional tbln = "check")
    Set varDic = CreateObject("Scripting.Dictionary")
    Dim num   As Long
    Dim rnum   As Long
    Dim numFn  As Long
    Dim numVar As Long
    Dim numAct As Long
    Dim hdNum1 As Long
    Dim hdNum2 As Long
    Dim vx   As String
    Dim el
    Dim z    As String
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
            el = getAryAt(rw, j)
            If TypeName(el) = "String" Then
                headNum1 = getHeadNum(el)
                vx = getVar(CStr(el))
                Select Case headNum1
                    Case 1
                        Call setAryAt(rw, j, varDic(vx))
                    Case 2
                        Call setAryAt(rw, j, varDic(vx), , True)
                    Case Is > 2
                        vx = Right(el, Len(el) - 2)
                        Call setAryAt(rw, j, vx)
                    Case Else
                End Select
            End If
        Next j
        z = rw(numVar)
        vz = getVar(z)
        hdNum2 = getHeadNum(z)
        If hdNum2 = 2 Then
            Set vl = evalTblRow(rw, numFn)
            Set varDic(vz) = vl
        ElseIf hdNum2 = 0 Or hdNum2 = 1 Then
            vl = evalTblRow(rw, numFn)
            varDic(vz) = vl
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
    Dim ret  As String
    Dim num  As Long
    num = getHeadNum(str)
    If num > 2 Then
        ret = Right(str, Len(str) - 2)
    Else
        ret = Right(str, Len(str) - num)
    End If
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

Sub addBtn(rn, mn, Optional cn = "run", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ThisWorkbook.Name
    Set rg = Workbooks(bn).Sheets(sn).Range(rn)
    Set btn = Workbooks(bn).Sheets(sn).Buttons.Add(rg.Left, rg.Top, rg.width, rg.Height)
    btn.OnAction = mn
    btn.Caption = cn
End Sub
