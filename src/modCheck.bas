Attribute VB_Name = "modCheck"
Enum debugPrint
    faNone = 0
    faDebugPrint = 1
    faPrintAry = 2
End Enum

Sub addBtn(rn, mn, Optional cn = "run", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ThisWorkbook.Name
    Set rg = Workbooks(bn).Sheets(sn).Range(rn)
    Set btn = Workbooks(bn).Sheets(sn).Buttons.Add(rg.Left, rg.Top, rg.width, rg.Height)
    btn.OnAction = mn
    btn.Caption = cn
End Sub

Sub mkCheckTbl(Optional tbln = "", Optional argnum As Long = 5)
    If tbln = "" Then tbln = ActiveCell.Value
    ary0 = Array("check", "assert", "actual", "expected", "variable", "function")
    ary1 = mkSeq(argnum)
    ary2 = mapA("addStr", ary1, "arg")
    ary3 = Array("statement")
    ary = conArys(ary0, ary2, ary3)
    n = lenAry(ary)
    x = ActiveCell.Address(False, False)
    y = ActiveCell.Offset(0, 1).Address(False, False)
    z = ActiveCell.Offset(1, 0).Resize(1, n).Address(False, False)
    Call addBtn(x, "'evalCheckTbl(""" & tbln & """)'", "eval")
    Call addBtn(y, "'clearResult(""" & tbln & """)'", "clear")
    Range(z) = ary
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(z), , xlYes).Name = tbln
    Range(tbln & "[check]").FormulaR1C1 = _
    "=IF(ISBLANK([@assert]),"""",IF(OR([@assert]=""="",[@assert]=""string""),IF([@expected]=[@actual],""pass"",""fail""),IF(OR([@assert]=TRUE,[@assert]=FALSE),IF(AND([@assert]=[@actual],NOT(ISBLANK([@actual]))),""pass"",""fail""),"""")))"
    Range(tbln).ListObject.TableStyle = "TableStyleLight9"
    Range(tbln & "[function]").Interior.ThemeColor = xlThemeColorAccent4
    Range(tbln & "[function]").Interior.TintAndShade = 0.599993896298105
    Range(tbln & "[statement]").Interior.ThemeColor = xlThemeColorAccent4
    Range(tbln & "[statement]").Interior.TintAndShade = 0.599993896298105
    Set fc1 = Range(tbln & "[check]").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""pass""")
    fc1.Interior.Color = 13561798
    Set fc2 = Range(tbln & "[check]").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""fail""")
    fc2.Interior.Color = 13551615
    With Range(tbln & "[assert]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="True,False,=,string"
    End With
End Sub

Sub clearResult(Optional tbln = "check")
    Range(tbln & "[actual]").ClearContents
    Range(tbln & "[statement]").ClearContents
End Sub

Sub evalCheckTbl(Optional tbln = "check", Optional dp As debugPrint = debugPrint.faNone)
    Set varDic = CreateObject("Scripting.Dictionary")
    Dim cNum As Long
    Dim rNum As Long
    Dim numFn As Long
    Dim numVar As Long
    Dim numAct As Long
    Dim numExp As Long
    Dim numKind As Long
    Dim numStt As Long
    Dim underBarNum1 As Long
    Dim underBarNum2 As Long
    Dim vx As String
    Dim el
    Dim vz0 As String
    Dim vz As String
    Dim vl
    Dim fnAry
    numFn = getClmNum("function", tbln)
    numVar = getClmNum("variable", tbln)
    numAct = getClmNum("actual", tbln)
    numExp = getClmNum("expected", tbln)
    numKind = getClmNum("assert", tbln)
    numStt = getClmNum("statement", tbln)
    rNum = Range(tbln).Rows.Count
    rws = rangeToArys(Range(tbln))
    Dim i As Long, j As Long
    For i = 1 To rNum
        rw = rws(i)
        fnAry0 = tblRowToFnAry(rw, numFn, numStt)
        fnAry = fnAry0
        cNum = lenAry(fnAry)
        For j = 2 To cNum
            assertKind = rw(numKind)
            expected = rw(numExp)
            el = getAryAt(fnAry, j)
            If TypeName(el) = "String" Then
                underBarNum1 = underBarCnt(el)
                vx = getVarStr(CStr(el))
                Select Case underBarNum1
                    Case 1, 2
                        Call setAryAt(fnAry, j, varDic(vx))
                    Case Is > 2
                        Call setAryAt(fnAry, j, Right(el, Len(el) - 2))
                    Case Else
                End Select
            End If
        Next j
        vz0 = rw(numVar)
        vz = getVarStr(vz0)
        retIsObj = underBarCnt(vz0) = 2
        If retIsObj Then
            Set vl = evalFnAry(fnAry, True)
            Set varDic(vz) = vl
        Else
            vl = evalFnAry(fnAry)
            varDic(vz) = vl
        End If
        Range(tbln & "[" & "actual" & "]")(i, 1) = toString(vl, , , , , True)
        Range(tbln & "[" & "statement" & "]")(i, 1) = mkStatement(fnAry0, vz, retIsObj, assertKind, expected, dp)
    Next i
End Sub

Private Function mkCheckFncStr(tbln, Optional dp As debugPrint = debugPrint.faNone) As String
    Dim ret As String
    Dim x
    'Call evalCheckTbl(tbl, dp)
    x = filterA("info_", rangeToAry(Range(tbln & "[statement]"), "c"), False, "isempty")
    ret = mcJoin(x, vbLf, "Sub Check" & tbln & vbLf, vbLf & "End sub")
    ret = Replace(ret, vbLf, vbCrLf)
    mkCheckFncStr = ret
End Function

Private Function mkCheckModStr(Optional sn = "", Optional bn = "") As String
    If bn = "" Then bn = ActiveWorkbook.Name
    If sn = "" Then sn = ActiveSheet.Name
    Dim ret As String
    Dim tbl
    ret = "Attribute VB_Name = ""Check" & sn & """" & vbCrLf
    For Each tbl In Workbooks(bn).Sheets(sn).ListObjects
        ret = ret & vbCrLf & mkCheckFncStr(tbl.Name) & vbCrLf
    Next
    mkCheckModStr = ret
End Function

Sub mkCheckFile(Optional sn = "", Optional bn = "")
    If bn = "" Then bn = ActiveWorkbook.Name
    If sn = "" Then sn = ActiveSheet.Name
    Dim str As String
    str = mkCheckModStr(sn, bn)
    pn = ThisWorkbook.Path & "\Check" & sn & ".bas"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stm = fso.CreateTextFile(pn)
    stm.Write (str)
    stm.Close
End Sub

Private Function elmToStr(elm)
    Dim ret As String, vz As String
    Dim num
    If TypeName(elm) = "String" Then
        num = underBarCnt(elm)
        vz = getVarStr(CStr(elm))
        If num = 1 Or num = 2 Then
            ret = vz
        Else
            ret = """" & vz & """"
        End If
    Else
        ret = CStr(elm)
    End If
    elmToStr = ret
End Function

Private Function fnAryToExp(fnAry0) As String
    Dim ret As String, fn As String, symbol As String
    Dim tmp
    fn = getAryAt(fnAry0, 1)
    tmp = mapA("elmToStr", dropAry(fnAry0, 1))
    If fn = "id_" Then
        ret = getAryAt(tmp, 1)
    ElseIf fn = "l_" Then
        ret = "Array" & mcJoin(tmp, ",", "(", ")")
    ElseIf fn = "calc_" Or fn = "comp_" Then
        symbol = Replace(getAryAt(tmp, 3), """", " ")
        If symbol = " % " Then symbol = " mod "
        ret = getAryAt(tmp, 1) & symbol & getAryAt(tmp, 2)
    ElseIf fn = "math_" Or fn = "info_" Then
        ret = Replace(getAryAt(tmp, 2), """", "") & "(" & getAryAt(tmp, 1) & ")"
    Else
        ret = fn & mcJoin(tmp, ",", "(", ")")
    End If
    fnAryToExp = ret
End Function

Private Function mkStatement(fnAry0, vz, retIsObj, Optional assertassert, Optional expected, Optional dp As debugPrint = debugPrint.faNone) As String
    Dim ret As String
    Dim fnStr As String
    fnStr = fnAryToExp(fnAry0)
    If vz = "" Then
        Select Case dp
            Case faNone: ret = ""
            Case faDebugPrint: ret = "Debug.Print " & fnStr
            Case faPrintAry: ret = "printAry " & fnStr
            Case Else
        End Select
    ElseIf retIsObj Then
        ret = "Set " & vz & " = " & fnStr
    Else
        ret = vz & " = " & fnStr
    End If
    If Not IsEmpty(assertassert) Then
        Select Case assertassert
            Case "="
                ret = ret & vbLf & "Assert " & vz & " , " & expected
            Case "string"
                ret = ret & vbLf & "Assert toString(" & vz & ")," & getStrExp(expected)
            Case True
                ret = ret & vbLf & "AssertTrue " & vz
            Case False
                ret = ret & vbLf & "AssertFalse " & vz
            Case Else
        End Select
    End If
    mkStatement = ret
End Function

Private Function underBarCnt(str, Optional chr = "_")
    Dim cnt As Long
    cnt = 0
    For i = 1 To Len(str)
        If Mid(str, i, 1) = chr Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next i
    underBarCnt = cnt
End Function

Private Function getVarStr(str As String) As String
    Dim ret As String
    Dim num As Long
    num = underBarCnt(str)
    If num > 2 Then
        ret = Right(str, Len(str) - 2)
    Else
        ret = Right(str, Len(str) - num)
    End If
    getVarStr = ret
End Function

Private Function getClmNum(clmnn, tbln, Optional bn = "")
    abn = ActiveWorkbook.Name
    If bn = "" Then bn = abn
    Workbooks(bn).Activate
    Dim ret
    With Application.WorksheetFunction
        ret = .Match(clmnn, Range(tbln & "[#headers]"), False)
    End With
    getClmNum = ret
    Workbooks(abn).Activate
End Function

Private Function tblRowToFnAry(rw, nfnc, nnext)
    Dim ary0, ary1, ret
    Dim n0 As Long
    n0 = lenAry(rw) - nnext + 1
    ary0 = dropAry(rw, nfnc - 1)
    ary1 = dropAry(ary0, n0, "r")
    ret = dropWhile("info_", ary1, "r", "isEmpty")
    tblRowToFnAry = ret
End Function

Private Function evalFnAry(fnAry, Optional retIsObj As Boolean = False)
    If retIsObj Then
        Set evalFnAry = evalObjA(fnAry)
    Else
        evalFnAry = evalA(fnAry)
    End If
End Function

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

Private Function getStrExp(str)
    ary = Split(str, vbCrLf)
    ary1 = mapA("addStr", ary, """", """")
    ret = Join(ary1, " & vbCrLf & ")
    ret = Replace(ret, vbLf, "")
    getStrExp = ret
End Function
