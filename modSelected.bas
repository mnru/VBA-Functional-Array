Attribute VB_Name = "modSelected"
'''''''''''''''''''''''''''''''''''''
' selected function from all modules
'''''''''''''''''''''''''''''''''''''


'''''''''''''''''
'from modAry
'''''''''''''''''

Option Base 0

Function lenAry(ary As Variant, Optional dm = 1) As Long
    lenAry = UBound(ary, dm) - LBound(ary, dm) + 1
End Function

Function getAryAt(ary, pos, Optional base = 1)
    Dim ret
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    ret = ary(idx)
    getAryAt = ret
End Function

Sub setAryAt(ByRef ary, pos, vl, Optional base = 1)
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    ary(idx) = vl
End Sub

Function getMAryAt(ary, pos, Optional base = 1)
    lsp = getAryShape(ary, "L")
    n = lenAry(lsp)
    bs = mkSameAry(base, n)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    ret = getElm(ary, idx2)
    getMAryAt = ret
End Function

Sub setMAryAt(ByRef ary, pos, vl, Optional base = 1)
    lsp = getAryShape(ary, "L")
    n = lenAry(lsp)
    bs = mkSameAry(base, n)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    Call setElm(vl, ary, idx2)
End Sub

Public Function dimAry(ByVal ary As Variant) As Long
    On Error GoTo Catch
    Dim idx     As Long
    idx = 0
    Do
        idx = idx + 1
        Dim tmp As Long
        tmp = UBound(ary, idx)
    Loop
Catch:
    dimAry = idx - 1
End Function

Function getAryShape(ary, Optional typ = "N")
    num = dimAry(ary)
    ReDim ret(0 To num - 1)
    For i = 1 To num
        Select Case UCase(typ)
            Case "N"
                tmp = lenAry(ary, i)
            Case "L"
                tmp = LBound(ary, i)
            Case "U"
                tmp = UBound(ary, i)
            Case Else
        End Select
        Call setAryAt(ret, i, tmp)
    Next i
    getAryShape = ret
End Function

Function getAryNum(ary)
    Dim ret
    sp = getAryShape(ary)
    'ret = reduceA("calc", sp, "*")
    ret = 1
    For Each elm In sp
        ret = ret * elm
    Next elm
    getAryNum = ret
End Function


Function conArys(ParamArray argArys())
    arys = argArys
    num = 0
    For Each ary In arys
        If IsArray(ary) Then
            num = num + lenAry(ary)
        Else
            num = num + 1
        End If
    Next ary
    ReDim ret(0 To num - 1)
    idx = 0
    For Each ary In arys
        If IsArray(ary) Then
            For Each elm In ary
                ret(idx) = elm
                idx = idx + 1
            Next elm
        Else
            ret(idx) = ary
            idx = idx + 1
        End If
    Next ary
    conArys = ret
End Function

Function mkSameAry(vl, num)
    ReDim ret(0 To num - 1)
    For i = 0 To num - 1
        ret(i) = vl
    Next i
    mkSameAry = ret
End Function

Function mkSeq(num, Optional first = 1, Optional step = 1)
    ReDim ret(0 To num - 1)
    For i = 0 To num - 1
        ret(i) = first + step * i
    Next i
    mkSeq = ret
End Function

Function dropAry(ary, num)
    lng = lenAry(ary)
    sz = lng - Abs(num)
    Dim ret
    If sz < 0 Then
        Call Err.Raise(1001, "dropAry", "num is larger than array length")
    ElseIf sz = 0 Then
        ret = Array()
    ElseIf num > 0 Then
        ReDim ret(0 To sz - 1)
        lb = LBound(ary)
        For i = 0 To sz - 1
            ret(i) = getAryAt(ary, i + num, 0)
        Next i
    Else
        ReDim ret(0 To sz - 1)
        ub = UBound(ary)
        For i = 0 To sz - 1
            ret(i) = getAryAt(ary, i, 0)
        Next i
    End If
    dropAry = ret
End Function

Function takeAry(ary, num)
    lng = lenAry(ary)
    sz = Abs(num)
    Dim ret
    If sz < 0 Then
        Call Err.Raise(1001, "takeAry", "num is larger than array length")
    End If
    If num > 0 Then
        ReDim ret(0 To sz - 1)
        lb = LBound(ary)
        For i = 0 To sz - 1
            ret(i) = getAryAt(ary, i, 0)
        Next i
    ElseIf num < 0 Then
        ReDim ret(0 To sz - 1)
        ' ub = UBound(ary)
        For i = 0 To sz - 1
            ret(i) = getAryAt(ary, lng - sz + i, 0)
        Next i
    Else
        ret = Array()
    End If
    takeAry = ret
End Function

Function revAry(ary)
    num = lenAry(ary)
    ReDim ret(0 To num - 1)
    lb = LBound(ary)
    For i = 0 To num - 1
        ret(i) = getAryAt(ary, num - i - 1, 0)
    Next i
    revAry = ret
End Function

Function zip(ParamArray argArys())
    arys = argArys
    ret = zipAry(arys)
    zip = ret
End Function

Function zipAry(arys)
    rnum = lenAry(arys)
    cnum = lenAry(arys(LBound(arys)))
    ReDim ret(0 To cnum - 1)
    lb = LBound(arys)
    For c = 0 To cnum - 1
        ReDim v(0 To rnum - 1)
        For r = 0 To rnum - 1
            v(r) = getAryAt(arys(lb + r), c, 0)
        Next r
        ret(c) = v
    Next c
    zipAry = ret
End Function

Function prmAry(ParamArray argAry())
    'flatten last elm
    ary = argAry
    ary1 = dropAry(ary, -1)
    ary2 = getAryAt(ary, -1)
    ret = conArys(ary1, ary2)
    prmAry = ret
End Function

Function inAry(ary As Variant, elm As Variant) As Boolean
    Dim ret   As Boolean
    ret = False
    For Each x In ary
        If x = elm Then
            ret = True
            Exit For
        End If
    Next x
    inAry = ret
End Function


Function mkIndex(num, shape, Optional lshape = Null)
    n = lenAry(shape)
    ReDim ret(0 To n - 1)
    r = num
    For i = n To 1 Step -1
        p = getAryAt(shape, i)
        Call setAryAt(ret, i, r Mod p)
        r = r \ p
    Next i
    If Not IsNull(lshape) Then
        For i = 1 To n
            Call setAryAt(ret, i, getAryAt(ret, i) + getAryAt(lshape, i))
        Next i
    End If
    mkIndex = ret
End Function

Function getElm(ary, idx)
    Dim ret
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: ret = ary(idx(lb))
        Case 2: ret = ary(idx(lb), idx(lb + 1))
        Case 3: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2))
        Case 4: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3))
        Case 5: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4))
        Case Else:
    End Select
    getElm = ret
End Function

Sub setElm(vl, ary, idx)
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: ary(idx(lb)) = vl
        Case 2: ary(idx(lb), idx(lb + 1)) = vl
        Case 3: ary(idx(lb), idx(lb + 1), idx(lb + 2)) = vl
        Case 4: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3)) = vl
        Case 5: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4)) = vl
        Case Else:
    End Select
End Sub

Sub setAryMbyS(mAry, sAry)
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, "L")
    n = getAryNum(mAry)
    For i = 0 To n - 1
        idx = mkIndex(i, sp, lsp)
        vl = getAryAt(sAry, i, 0)
        Call setElm(vl, mAry, idx)
    Next i
End Sub

Function getArySbyM(mAry, Optional bs = 0)
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, "L")
    n = getAryNum(mAry)
    ReDim ret(bs To bs + n - 1)
    For i = 0 To n - 1
        idx = mkIndex(i, sp, lsp)
        vl = getElm(mAry, idx)
        Call setAryAt(ret, i, vl, 0)
    Next i
    getArySbyM = ret
End Function

Function reshapeAry(ary, sp, Optional bs = 0)
    n = lenAry(sp)
    ret = mkMAry(sp, bs)
    Call setAryMbyS(ret, ary)
    reshapeAry = ret
    
End Function

Function calc(num1, num2, symbol As String)
    Dim ret
    Select Case symbol
        Case "+": ret = num1 + num2
        Case "-": ret = num1 - num2
        Case "*": ret = num1 * num2
        Case "/": ret = num1 / num2
        Case "\": ret = num1 \ num2
        Case "%": ret = num1 Mod num2
        Case "^": ret = num1 ^ num2
        Case Else
    End Select
    calc = ret
End Function

Function calcAry(ary1, ary2, symbol As String)
    n = lenAry(ary1)
    ReDim ret(0 To n - 1)
    For i = 0 To n - 1
        ret(i) = calc(getAryAt(ary1, i, 0), getAryAt(ary2, i, 0), symbol)
    Next i
    calcAry = ret
End Function

Function calcMAry(ary1, ary2, symbol As String, Optional bs = 0)
    sp1 = getAryShape(ary1)
    sp2 = getAryShape(ary2)
    lsp1 = getAryShape(ary1, "L")
    lsp2 = getAryShape(ary2, "L")
    n = getAryNum(ary1)
    dm = lenAry(sp1)
    ret = mkMAry(sp1, bs)
    lsp0 = mkSameAry(bs, dm)
    For i = 0 To n - 1
        idx = mkIndex(i, sp1)
        idx1 = calcAry(idx, lsp1, "+")
        idx2 = calcAry(idx, lsp2, "+")
        idx0 = calcAry(idx, lsp0, "+")
        
        vl = calc(getElm(ary1, idx1), getElm(ary2, idx2), symbol)
        Call setElm(vl, ret, idx0)
    Next i
    calcMAry = ret
End Function


Function mkMAry(sp, Optional bs = 0)
    n = lenAry(sp)
    ub = calcAry(sp, mkSameAry(bs - 1, n), "+")
    lb = LBound(ub)
    Dim ret
    Select Case n
        Case 1: ReDim ret(bs To ub(lb))
        Case 2: ReDim ret(bs To ub(lb), bs To ub(lb + 1))
        Case 3: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2))
        Case 4: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3))
        Case 5: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4))
        Case Else:
    End Select
    mkMAry = ret
End Function

Function l_(ParamArray argAry() As Variant)
    'works like function array()
    Dim ary As Variant
    ary = argAry
    l_ = ary
End Function

Sub setMArySeq(ary, Optional first = 1, Optional step = 1)
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    num = getAryNum(ary)
    vl = first
    For i0 = 0 To num - 1
        idx = mkIndex(i0, sp, lsp)
        'vl = first + i0 * step
        Call setElm(vl, ary, idx)
        vl = vl + step
    Next i0
End Sub


Function mkMArySeq(sp, Optional first = 1, Optional step = 1, Optional bs = 0)
    ret = mkMAry(sp, bs)
    Call setMArySeq(ret, first, step)
    mkMArySeq = ret
End Function

Function uniqueAry(ary)
    Set dic = CreateObject("Scripting.Dictionary")
    For Each elm In ary
        dic(elm) = Null
    Next
    ret = dic.keys
    uniqueAry = ret
End Function

'''''''''''''''''
'from modFnc
'''''''''''''''''

Public Function evalA(argAry As Variant) As Variant
    ary = argAry
    Dim ret As Variant
    lb = LBound(ary)
    Select Case lenAry(ary)
        Case 1: ret = Application.Run(ary(lb))
        Case 2: ret = Application.Run(ary(lb), ary(lb + 1))
        Case 3: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2))
        Case 4: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3))
        Case 5: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4))
        Case 6: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5))
        Case 7: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6))
        Case 8: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7))
        Case 9: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8))
        Case 10: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9))
        Case 11: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10))
        Case 12: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11))
        Case 13: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12))
        Case 14: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13))
        Case 15: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14))
        Case Else:
    End Select
    evalA = ret
End Function

Public Function mapA(fnc As String, seq As Variant, ParamArray argAry() As Variant) As Variant
    ary = argAry
    fnAry = prmAry(fnc, Null, ary)
    num = lenAry(seq)
    ReDim ret(1 To num)
    Dim i As Long
    For i = 1 To num
        Call setAryAt(fnAry, 2, getAryAt(seq, i))
        ret(i) = evalA(fnAry)
    Next i
    mapA = ret
End Function

Public Function mMapA(fnc As String, mAry As Variant, ParamArray argAry() As Variant) As Variant
    ary = argAry
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, "L")
    ret = mkMAry(sp)
    fnAry = prmAry(fnc, Null, ary)
    num = getAryNum(mAry)
    Dim i As Long
    For i = 1 To num
        idx0 = mkIndex(i, sp)
        idx = calcAry(idx0, lsp, "+")
        Call setAryAt(fnAry, 2, getElm(mAry, idx))
        vl = evalA(fnAry)
        Call setElm(vl, ret, idx0)
    Next i
    mMapA = ret
End Function

Public Function filterA(fnc As String, seq As Variant, ParamArray argAry() As Variant) As Variant
    ary = argAry
    num = lenAry(seq)
    fnAry = prmAry(fnc, Null, ary)
    idx = 0
    ReDim ret(1 To num)
    For Each elm In seq
        Call setAryAt(fnAry, 2, elm)
        If evalA(fnAry) Then
            idx = idx + 1
            ret(idx) = elm
        End If
    Next elm
    ReDim Preserve ret(1 To idx)
    filterA = ret
End Function

Public Function foldA(fnc As String, seq As Variant, init As Variant, ParamArray argAry() As Variant) As Variant
    ary = argAry
    ret = foldAryA(fnc, seq, init, ary)
    foldA = ret
End Function

Public Function foldAryA(fnc As String, seq As Variant, init As Variant, ary) As Variant
    fnAry = prmAry(fnc, init, Null, ary)
    n = lenAry(seq)
    ret = init
    For Each elm In seq
        Call setAryAt(fnAry, 1, ret, 0)
        Call setAryAt(fnAry, 2, elm, 0)
        ret = evalA(fnAry)
        
    Next elm
    foldAryA = ret
End Function

Public Function reduceA(fnc As String, seq As Variant, ParamArray argAry() As Variant) As Variant
    ary = argAry
    init = getAryAt(seq, 1)
    seq1 = dropAry(seq, 1)
    ret = foldAryA(fnc, seq1, init, ary)
    
    reduceA = ret
End Function

Public Function foldF(fnObj, seq As Variant, init As Variant) As Variant
    ret = init
    For Each elm In seq
        ret = applyF(Array(ret, elm), fnObj, True)
    Next elm
    foldF = ret
End Function

Public Function reduceF(fnObj, seq As Variant) As Variant
    init = getAryAt(seq, 1)
    seq1 = dropAry(seq, 1)
    ret = foldF(fnObj, seq1, init)
    reduceF = ret
End Function

Function applyF(vl, fnObj, Optional argAsAry = False)
    Dim ret
    fnAry = getAryAt(fnObj, 2)
    arity = getAryAt(fnObj, 1)
    If Not argAsAry Then
        Call setAryAt(fnAry, getAryAt(arity, 1), vl, 0)
    Else
        n = lenAry(arity)
        For i = 1 To n
            Call setAryAt(fnAry, getAryAt(arity, i), getAryAt(vl, i), 0)
        Next i
    End If
    ret = evalA(fnAry)
    applyF = ret
End Function

Function applyFs(vl, fnObjs, Optional argAsAry = False)
    Dim ret
    ret = vl
    For Each fnObj In fnObjs
        ret = applyF(ret, fnObj, argAsAry)
    Next fnObj
    applyFs = ret
End Function
Function getArity(ary)
    Dim ret
    ret = 0
    For Each elm In ary
        If IsNumeric(elm) Then
            ret = ret + 1
        Else
            Exit For
        End If
    Next elm
    getArity = ret
End Function

Function mkF(ParamArray argArys())
    ary = argArys
    n = getArity(ary)
    arity = takeAry(ary, n)
    fnAry = dropAry(ary, n)
    mkF = Array(arity, fnAry)
End Function

Function zipApplyF(fnObj, ParamArray argAry())
    arys = argAry
    x = zipAry(arys)
    ret = mapA("applyF", x, fnObj, True)
    zipApplyF = ret
End Function

Sub setAryMByF(ary, fnObj)
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    n = getAryNum(ary)
    For i = 0 To n - 1
        idx = mkIndex(i, sp, lsp)
        vl = applyF(i, fnObj)
        Call setElm(vl, ary, idx)
    Next i
End Sub

Function negate(fnc As String, x, ParamArray argAry())
    ary = argAry
    Dim ret
    fnAry = prmAry(fnc, x, ary)
    ret = Not evalA(fnAry)
    negate = ret
End Function

Function takeWhile(fnc, ary, direction, ParamArray argAry())
    prm = argAry
    fnAry = prmAry(fnc, Null, prm)
    
    n = lenAry(ary)
    
    sn = Sgn(direction)
    num = 0
    For i = 1 To n
        v = getAryAt(ary, sn * i)
        Call setAryAt(fnAry, 1, v, 0)
        If evalA(fnAry) Then
            num = num + 1
        Else
            Exit For
        End If
    Next
    ret = takeAry(ary, sn * num)
    takeWhile = ret
End Function


Function dropWhile(fnc, ary, direction, ParamArray argAry())
    prm = argAry
    fnAry = prmAry(fnc, Null, prm)
    
    n = lenAry(ary)
    
    sn = Sgn(direction)
    num = 0
    For i = 1 To n
        v = getAryAt(ary, sn * i)
        Call setAryAt(fnAry, 1, v, 0)
        If evalA(fnAry) Then
            num = num + 1
        Else
            Exit For
        End If
    Next
    ret = dropAry(ary, sn * num)
    dropWhile = ret
End Function

'''''''''''''''''
'from modUtil
'''''''''''''''''

Function toString(elm, Optional qt = True, Optional fm = "", Optional lcr = "r", Optional width = 0, _
    Optional insheet As Boolean = False) As String
    
    Dim ret
    ret = ""
    If IsArray(elm) Then
        d = dimAry(elm)
        ret = ret & "["
        sp = getAryShape(elm)
        lsp = getAryShape(elm, "L")
        aryNum = getAryNum(elm)
        If aryNum = 0 Then
            ret = ret & "]"
        Else
            For i = 0 To aryNum - 1
                idx0 = mkIndex(i, sp)
                idx = calcAry(idx0, lsp, "+")
                vl = getElm(elm, idx)
                dlm = getDlm(sp, idx0)
                ret = ret & toString(vl, qt, fm, lcr, width) & dlm
            Next i
        End If
    Else
        If IsObject(elm) Then
            ret = ret & "<" & TypeName(elm) & ">"
        ElseIf IsNull(elm) Then
            ret = ret & "Null"
            
        Else
            If TypeName(elm) = "String" Then
                If qt Then
                    tmp = "'" & elm & "'"
                Else
                    tmp = elm
                End If
            Else
                tmp = fmt(elm, fm)
            End If
            tmp = align(tmp, lcr, width)
            ret = ret & tmp
            
        End If
    End If
    toString = ret
End Function

Function getDlm(shape, idx, Optional insheet As Boolean = False)
    Dim ret
    Dim nl
    nl = IIf(insheet, vbLf, vbCrLf)
    n = lenAry(shape)
    m = 0
    For i = n To 1 Step -1
        If getAryAt(shape, i) - 1 > getAryAt(idx, i) Then
            m = i
            Exit For
        End If
    Next i
    Select Case m
        Case 0
            ret = "]"
        Case n
            ret = ","
        Case n - 1
            ret = ";" & nl & " "
        Case Else
            ret = String(n - m, ";") & nl & nl & " "
    End Select
    getDlm = ret
End Function

Function secToHMS(vl As Double)
    'Dim x2 As Double
    x0 = vl
    x1 = Int(x0)
    x2 = x0 - x1
    x3 = mkIndex(x1, Array(24, 60, 60))
    x4 = getAryAt(x3, 3) + x2
    ret = Format(getAryAt(x3, 1), "00") & ":" & Format(getAryAt(x3, 2), "00") & ":" & Format(x4, "00.000")
    secToHMS = ret
End Function

Function clcToAry(clc As Collection)
    cnt = clc.Count
    ReDim ret(1 To cnt)
    For i = 1 To cnt
        ret(i) = clc.item(i)
    Next i
    clcToAry = ret
End Function

Function flattenAry(ary)
    Dim clc As Collection
    Set clc = New Collection
    
    For Each elm In ary
        If IsArray(elm) Then
            For Each el In flattenAry(elm)
                clc.Add el
            Next el
        Else
            clc.Add elm
        End If
    Next elm
    
    ret = clcToAry(clc)
    flattenAry = ret
    
End Function

Function mcLike(word As String, wildcard As String, Optional include As Boolean = True) As Boolean
    Dim bol As Boolean
    bol = word Like wildcard
    mcLike = IIf(include, bol, Not bol)
End Function

Function mcJoin(ary, Optional dlm As String = "", Optional pre As String = "", Optional suf As String = "") As String
    Dim ret As String
    ret = pre & Join(ary, dlm) & suf
    mcJoin = ret
End Function

Function addStr(body As String, Optional prefix As String = "", Optional suffix As String = "")
    addStr = prefix & body & suffix
End Function

Function poly(x, polyAry)
    lb = LBound(polyAry)
    ub = UBound(polyAry)
    ret = polyAry(lb)
    For i = lb + 1 To ub
        ret = ret * x + polyAry(i)
    Next
    poly = ret
End Function

Function polyStr(polyAry)
    ret = ""
    n = lenAry(polyAry)
    For i = 1 To n
        c = getAryAt(polyAry, i)
        If c <> 0 Then
            If ret <> "" Then ret = ret & " "
            If c > 0 Then ret = ret & "+"
            If c <> 1 Or i = n Then ret = ret & c
            If i < n Then ret = ret & "X"
            If i < n - 1 Then ret = ret & "^" & n - i
        End If
    Next i
    If ret = "" Then ret = getAryAt(polyAry, -1)
    If Left(ret, 1) = "+" Then ret = Right(ret, Len(ret) - 1)
    polyStr = ret
    
End Function

Function fmt(expr, Optional fm = "", Optional lcr = "r", Optional width = 0)
    ret = Format(expr, fm)
    ret = align(ret, lcr, width)
    fmt = ret
End Function

Function align(str, Optional lcr = "r", Optional width = 0)
    Dim ret
    ret = CStr(str)
    d = width - Len(ret)
    If d > 0 Then
        Select Case LCase(lcr)
            Case "r": ret = space(d) & ret
            Case "l": ret = ret & space(d)
            Case "c": ret = space(d \ 2) & ret & space(d - d \ 2)
            Case Else:
        End Select
    End If
    
    align = ret
    
    
End Function


Function math(x, symbol)
    Dim ret
    Select Case LCase(symbol)
        Case "sin": ret = Sin(x)
        Case "cos": ret = Cos(x)
        Case "tan": ret = Tan(x)
        Case "atn": ret = Atn(x)
        Case "log": ret = Log(x)
        Case "exp": ret = Exp(x)
        Case "sqr": ret = Sqr(x)
        Case "abs": ret = Abs(x)
        Case "sgn": ret = Sgn(x)
        Case Else:
    End Select
    math = ret
    
End Function

Function comp(x, y, symbol)
    Dim ret
    Select Case symbol
        Case "=": ret = x = y 'caution assign and eqaul is same symbol
        Case "<>": ret = x <> y
        Case "<": ret = x < y
        Case ">": ret = x > y
        Case "<=": ret = x <= y
        Case ">=": ret = x >= y
        Case "<": ret = x < y
        Case Else:
    End Select
    comp = ret
End Function

Function info(x, symbol)
    Dim ret
    Select Case LCase(symbol)
        Case "isarray": ret = IsArray(x)
        Case "isdate": ret = IsDate(x)
        Case "isempty": ret = IsEmpty(x)
        Case "iserror": ret = IsError(x)
        Case "ismissing": ret = IsMissing(x)
        Case "isnull": ret = IsNull(x)
        Case "isnumeric": ret = IsNumeric(x)
        Case "isobject": ret = IsObject(x)
        Case "typename": ret = TypeName(x)
        Case "vartype": ret = VarType(x)
        Case Else:
    End Select
    info = ret
    
End Function

Function id_(x)
    id_ = x
End Function

'''''''''''''''''
'from modLog
'''''''''''''''''

Sub outPut(Optional msg = "", Optional crlf As Boolean = True)
    If crlf Then
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
End Sub

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim etime  As Double
    Dim stime  As Double
    Dim secs   As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call outPut(fnc & " - " & secToHMS(secs), True)
End Function

Sub printAry(ary, Optional qt = True, Optional fm = "", Optional lcr = "r", Optional width = 0)
    Call outPut(toString(ary, qt, fm, lcr, width), True)
End Sub

