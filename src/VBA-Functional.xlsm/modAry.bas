Attribute VB_Name = "modAry"

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
    ReDim ret(1 To num)
    idx = 1
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
    ReDim ret(1 To num)
    For i = 1 To num
        ret(i) = vl
    Next i
    mkSameAry = ret
End Function

Function mkSeq(ParamArray argAry())
    ary = argAry
    Dim first
    Dim last
    Dim step
    argn = lenAry(ary)
    Select Case argn
        Case 1
            first = 1
            last = getAryAt(ary, 1)
            step = IIf(first <= last, 1, -1)
        Case 2
            first = getAryAt(ary, 1)
            last = getAryAt(ary, 2)
            step = IIf(first <= last, 1, -1)
        Case 3
            step = Abs(getAryAt(ary, 3))
            first = getAryAt(ary, 1)
            last = getAryAt(ary, 2)
            step = IIf(first <= last, step, -1 * step)
        Case Else
    End Select
    n = Int((last - first) / step) + 1
    ReDim ret(1 To n)
    For i = 1 To n
        ret(i) = first + step * (i - 1)
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
        ReDim ret(1 To sz)
        lb = LBound(ary)
        For i = 1 To sz
            ret(i) = getAryAt(ary, i + num)
        Next i
    Else
        ReDim ret(1 To sz)
        ub = UBound(ary)
        For i = 1 To sz
            ret(i) = getAryAt(ary, i)
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
        ReDim ret(1 To sz)
        lb = LBound(ary)
        For i = 1 To sz
            ret(i) = getAryAt(ary, i)
        Next i
    ElseIf num < 0 Then
        ReDim ret(1 To sz)
        ub = UBound(ary)
        For i = 1 To sz
            ret(i) = getAryAt(ary, l - num + i)
        Next i
    Else
        ret = Array()
    End If
    takeAry = ret
End Function

Function revAry(ary)
    num = lenAry(ary)
    ReDim ret(1 To num)
    lb = LBound(ary)
    For i = 1 To num
        ret(i) = getAryAt(num - i + 1)
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
    ReDim ret(1 To cnum)
    lb = LBound(arys)
    For c = 1 To cnum
        ReDim v(1 To rnum)
        For r = 1 To rnum
            v(r) = getAryAt(arys(lb + r - 1), c)
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
    Dim ret     As Boolean
    ret = False
    For Each x In ary
        If x = elm Then
            ret = True
            Exit For
        End If
    Next x
    inAry = ret
End Function
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
    ReDim ret(1 To num)
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
    ret = reduceA("calc", sp, "*")
    getAryNum = ret
End Function

Function mkIndex(num, shape, Optional lshape = Null)
    n = lenAry(shape)
    ReDim ret(1 To n)
    r = num
    For i = n To 1 Step -1
        p = getAryAt(shape, i)
        Call setAryAt(ret, i, r Mod p)
        r = r \ p
    Next i
    If Not IsNull(lshape) Then
        For i = 1 To n
            Call setAryAt(ret, i, ret(i) + getAryAt(lshape, i))
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
        Case 6: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5))
        Case 7: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6))
        Case 8: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7))
        Case 9: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8))
        Case 10: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9))
        Case 11: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10))
        Case 12: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11))
        Case 13: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12))
        Case 14: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13))
        Case 15: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14))
        Case 16: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15))
        Case 17: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16))
        Case 18: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17))
        Case 19: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18))
        Case 20: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19))
        Case 21: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20))
        Case 22: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21))
        Case 23: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22))
        Case 24: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23))
        Case 25: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24))
        Case 26: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25))
        Case 27: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26))
        Case 28: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27))
        Case 29: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28))
        Case 30: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29))
        Case 31: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30))
        Case 32: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31))
        Case 33: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32))
        Case 34: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33))
        Case 35: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34))
        Case 36: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35))
        Case 37: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36))
        Case 38: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37))
        Case 39: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38))
        Case 40: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39))
        Case 41: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40))
        Case 42: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41))
        Case 43: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42))
        Case 44: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43))
        Case 45: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44))
        Case 46: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45))
        Case 47: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46))
        Case 48: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47))
        Case 49: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48))
        Case 50: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49))
        Case 51: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50))
        Case 52: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51))
        Case 53: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52))
        Case 54: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53))
        Case 55: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54))
        Case 56: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55))
        Case 57: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56))
        Case 58: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57))
        Case 59: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58))
        Case 60: ret = ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58), idx(lb + 59))
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
        Case 6: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5)) = vl
        Case 7: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6)) = vl
        Case 8: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7)) = vl
        Case 9: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8)) = vl
        Case 10: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9)) = vl
        Case 11: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10)) = vl
        Case 12: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11)) = vl
        Case 13: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12)) = vl
        Case 14: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13)) = vl
        Case 15: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14)) = vl
        Case 16: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15)) = vl
        Case 17: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16)) = vl
        Case 18: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17)) = vl
        Case 19: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18)) = vl
        Case 20: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19)) = vl
        Case 21: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20)) = vl
        Case 22: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21)) = vl
        Case 23: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22)) = vl
        Case 24: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23)) = vl
        Case 25: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24)) = vl
        Case 26: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25)) = vl
        Case 27: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26)) = vl
        Case 28: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27)) = vl
        Case 29: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28)) = vl
        Case 30: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29)) = vl
        Case 31: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30)) = vl
        Case 32: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31)) = vl
        Case 33: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32)) = vl
        Case 34: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33)) = vl
        Case 35: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34)) = vl
        Case 36: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35)) = vl
        Case 37: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36)) = vl
        Case 38: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37)) = vl
        Case 39: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38)) = vl
        Case 40: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39)) = vl
        Case 41: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40)) = vl
        Case 42: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41)) = vl
        Case 43: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42)) = vl
        Case 44: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43)) = vl
        Case 45: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44)) = vl
        Case 46: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45)) = vl
        Case 47: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46)) = vl
        Case 48: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47)) = vl
        Case 49: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48)) = vl
        Case 50: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49)) = vl
        Case 51: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50)) = vl
        Case 52: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51)) = vl
        Case 53: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52)) = vl
        Case 54: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53)) = vl
        Case 55: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54)) = vl
        Case 56: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55)) = vl
        Case 57: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56)) = vl
        Case 58: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57)) = vl
        Case 59: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58)) = vl
        Case 60: ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58), idx(lb + 59)) = vl
        Case Else:
    End Select
End Sub
