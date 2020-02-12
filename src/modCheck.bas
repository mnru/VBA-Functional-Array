Attribute VB_Name = "modCheck"
Sub checkTbl(Optional tbln = "check")
    Set varDic = CreateObject("Scripting.Dictionary")
    Dim num  As Long
    Dim rnum  As Long
    Dim numFn As Long
    Dim numVar As Long
    Dim numAct As Long
    Dim hdNum1 As Long
    Dim hdNum2 As Long
    Dim vx  As String
    Dim el
    Dim z  As String
    Dim fnAry
    numFn = clmNum("function", tbln)
    numVar = clmNum("variable", tbln)
    numAct = clmNum("actual", tbln)
    numexp = clmNum("expected", tbln)
    numkind = clmNum("kind", tbln)
    rnum = Range(tbln).Rows.Count
    rws = rangeToArys(Range(tbln))
    Dim i As Long, j As Long
    For i = 1 To rnum
        rw = rws(i)
        fnAry0 = tblRowToFncAry(rw, numFn)
        fnAry = fnAry0
        cnum = lenAry(fnAry)
        For j = 2 To cnum
            withAssert = rw(numkind) = "="
            expected = rw(numexp)
            el = getAryAt(fnAry, j)
            If TypeName(el) = "String" Then
                headNum1 = getHeadNum(el)
                vx = getVar(CStr(el))
                Select Case headNum1
                    Case 1
                        Call setAryAt(fnAry, j, varDic(vx))
                    Case 2
                        Call setAryAt(fnAry, j, varDic(vx), , True)
                    Case Is > 2
                        vx = Right(el, Len(el) - 2)
                        Call setAryAt(fnAry, j, vx)
                    Case Else
                End Select
            End If
        Next j
        z = rw(numVar)
        vz = getVar(z)
        hdNum2 = getHeadNum(z)
        If hdNum2 = 2 Then
            retIsObj = True
        Else
            retIsObj = False
        End If
        If hdNum2 = 2 Then
            Set vl = evalFnAry(fnAry, True)
            Set varDic(vz) = vl
        ElseIf hdNum2 = 0 Or hdNum2 = 1 Then
            vl = evalFnAry(fnAry)
            varDic(vz) = vl
        End If
        Range(tbln & "[" & "actual" & "]")(i, 1) = toString(vl, , , , , True)
        Range(tbln & "[" & "statement" & "]")(i, 1) = mkStatement(fnAry0, vz, retIsObj, withAssert, expected)
    Next i
End Sub

Function mkTest(tbl)
    checkTbl (tbl)
    x = filterA("info", rangeToAry(Range(tbl & "[statement]"), "c"), False, "isempty")
    ret = mcJoin(x, vbLf, "Sub test" & tbl & vbLf, vbLf & "End sub")
    ret = Replace(ret, vbLf, vbCrLf)
    mkTest = ret
End Function

Function fnAryElmToStr(elm)
    Dim ret
    If TypeName(elm) = "String" Then
        num = getHeadNum(elm)
        vz = getVar(CStr(elm))
        If num = 1 Or num = 2 Then
            ret = vz
        Else
            ret = """" & vz & """"
        End If
    Else
        ret = CStr(elm)
    End If
    fnAryElmToStr = ret
End Function

Function fnaryToStr(fnAry)
    Dim ret, tmp
    Dim fn
    fn = getAryAt(fnAry, 1)
    tmp = mapA("fnAryElmToStr", dropAry(fnAry, 1))
    If fn = "id_" Then
        ret = getAryAt(tmp, 1)
    ElseIf fn = "l_" Then
        ret = "Array" & mcJoin(tmp, ",", "(", ")")
    ElseIf fn = "calc" Then
        ret = getAryAt(tmp, 1) & Replace(getAryAt(tmp, 3), """", " ") & getAryAt(tmp, 2)
    ElseIf fn = "math" Or fn = "info" Then
        ret = Replace(getAryAt(tmp, 2), """", "") & "(" & getAryAt(tmp, 1) & ")"
    Else
        ret = fn & mcJoin(tmp, ",", "(", ")")
    End If
    fnaryToStr = ret
End Function

Function mkStatement(fnAry, vz, retIsObj, Optional withAssert, Optional expected)
    Dim ret
    If vz = "" Then
        'ret = "Call " & fnaryToStr(fnAry)
        ret = ""
    ElseIf retIsObj Then
        ret = "Set " & vz & " = " & fnaryToStr(fnAry)
    Else
        ret = vz & " = " & fnaryToStr(fnAry)
    End If
    If withAssert Then
        ret = ret & vbLf & "Assert " & vz & " , " & expected
    End If
    mkStatement = ret
End Function

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

Function tblRowToFncAry(rw, fncl)
    Dim ary, ret
    ary = dropAry(rw, fncl - 1)
    ret = dropWhile("info", ary, -1, "isEmpty")
    tblRowToFncAry = ret
End Function

Function evalFnAry(fnAry, Optional retIsObj As Boolean = False)
    If retIsObj Then
        Set evalFnAry = evalObjA(fnAry)
    Else
        evalFnAry = evalA(fnAry)
    End If
End Function

Function getVar(str As String) As String
    Dim ret As String
    Dim num As Long
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

Public Function evalObjA(argAry As Variant) As Variant
    Dim lb As Long
    ary = argAry
    Dim ret As Variant
    lb = LBound(ary)
    Select Case lenAry(ary)
        Case 1: Set ret = Application.Run(ary(lb))
        Case 2: Set ret = Application.Run(ary(lb), ary(lb + 1))
        Case 3: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2))
        Case 4: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3))
        Case 5: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4))
        Case 6: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5))
        Case 7: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6))
        Case 8: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7))
        Case 9: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8))
        Case 10: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9))
        Case 11: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10))
        Case 12: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11))
        Case 13: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12))
        Case 14: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13))
        Case 15: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14))
        Case 16: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15))
        Case 17: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16))
        Case 18: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17))
        Case 19: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18))
        Case 20: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19))
        Case 21: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20))
        Case 22: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21))
        Case 23: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22))
        Case 24: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23))
        Case 25: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24))
        Case 26: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25))
        Case 27: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26))
        Case 28: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27))
        Case 29: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28))
        Case 30: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28), ary(lb + 29))
        Case 31: Set ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28), ary(lb + 29), ary(lb + 30))
        Case Else:
    End Select
    Set evalObjA = ret
End Function
