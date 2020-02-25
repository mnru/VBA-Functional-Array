Attribute VB_Name = "modFnc"
Option Base 0
Option Explicit

Public Function evalA(argAry As Variant) As Variant
    Dim lb As Long
    Dim ary
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
        Case 16: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15))
        Case 17: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16))
        Case 18: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17))
        Case 19: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18))
        Case 20: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19))
        Case 21: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20))
        Case 22: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21))
        Case 23: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22))
        Case 24: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23))
        Case 25: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24))
        Case 26: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25))
        Case 27: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26))
        Case 28: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27))
        Case 29: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28))
        Case 30: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28), ary(lb + 29))
        Case 31: ret = Application.Run(ary(lb), ary(lb + 1), ary(lb + 2), ary(lb + 3), ary(lb + 4), ary(lb + 5), ary(lb + 6), ary(lb + 7), ary(lb + 8), ary(lb + 9), ary(lb + 10), ary(lb + 11), ary(lb + 12), ary(lb + 13), ary(lb + 14), ary(lb + 15), ary(lb + 16), ary(lb + 17), ary(lb + 18), ary(lb + 19), ary(lb + 20), ary(lb + 21), ary(lb + 22), ary(lb + 23), ary(lb + 24), ary(lb + 25), ary(lb + 26), ary(lb + 27), ary(lb + 28), ary(lb + 29), ary(lb + 30))
        Case Else:
    End Select
    evalA = ret
End Function

Public Function mapA(fnc As String, seq As Variant, ParamArray argAry() As Variant) As Variant
    Dim ary, fnAry
    Dim lNum As Long
    ary = argAry
    fnAry = prmAry(fnc, Empty, ary)
    lNum = lenAry(seq)
    ReDim ret(1 To lNum)
    Dim i As Long
    For i = 1 To lNum
        Call setAryAt(fnAry, 2, getAryAt(seq, i))
        ret(i) = evalA(fnAry)
    Next i
    mapA = ret
End Function

Public Function mapMA(fnc As String, mAry As Variant, ParamArray argAry() As Variant) As Variant
    Dim ary, sp, lsp, fnAry, ret, idx, idx0, vl
    Dim aNum As Long
    ary = argAry
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, "L")
    ret = mkMAry(sp)
    fnAry = prmAry(fnc, Empty, ary)
    aNum = getAryNum(mAry)
    Dim i As Long
    For i = 1 To aNum
        idx0 = mkIndex(i, sp)
        idx = calcAry(idx0, lsp, "+")
        Call setAryAt(fnAry, 2, getElm(mAry, idx))
        vl = evalA(fnAry)
        Call setElm(vl, ret, idx0)
    Next i
    mapMA = ret
End Function

Public Function filterA(fnc As String, seq As Variant, affirmative As Boolean, ParamArray argAry() As Variant) As Variant
    Dim lNum As Long, i As Long
    Dim ary, fnAry, elm
    Dim bol As Boolean
    ary = argAry
    lNum = lenAry(seq)
    fnAry = prmAry(fnc, Empty, ary)
    i = 0
    ReDim ret(1 To lNum)
    For Each elm In seq
        Call setAryAt(fnAry, 2, elm)
        bol = evalA(fnAry)
        If Not affirmative Then bol = Not bol
        If bol Then
            i = i + 1
            ret(i) = elm
        End If
    Next elm
    ReDim Preserve ret(1 To i)
    filterA = ret
End Function

Public Function foldA(fnc As String, seq As Variant, init As Variant, ParamArray argAry() As Variant) As Variant
    Dim ary, ret
    ary = argAry
    ret = foldAryA(fnc, seq, init, ary)
    foldA = ret
End Function

Public Function foldAryA(fnc As String, seq As Variant, init As Variant, ary) As Variant
    Dim fnAry, elm, ret
    Dim n As Long
    fnAry = prmAry(fnc, init, Empty, ary)
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
    Dim ary, seq1, ret, init
    ary = argAry
    init = getAryAt(seq, 1)
    seq1 = dropAry(seq, 1)
    ret = foldAryA(fnc, seq1, init, ary)
    reduceA = ret
End Function

Public Function foldF(fnObj, seq As Variant, init As Variant) As Variant
    Dim ret, elm
    ret = init
    For Each elm In seq
        ret = applyF(Array(ret, elm), fnObj, True)
    Next elm
    foldF = ret
End Function

Public Function reduceF(fnObj, seq As Variant) As Variant
    Dim init, seq1, ret
    init = getAryAt(seq, 1)
    seq1 = dropAry(seq, 1)
    ret = foldF(fnObj, seq1, init)
    reduceF = ret
End Function

Function applyF(vl, fnObj, Optional argAsAry = False)
    Dim ret, fnAry, arity
    Dim n As Long, i As Long
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
    Dim ret, fnObj
    ret = vl
    For Each fnObj In fnObjs
        ret = applyF(ret, fnObj, argAsAry)
    Next fnObj
    applyFs = ret
End Function

Function getArity(ary)
    Dim ret, elm
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
    Dim n As Long
    Dim ary, fnAry, arity
    ary = argArys
    n = getArity(ary)
    arity = takeAry(ary, n)
    fnAry = dropAry(ary, n)
    mkF = Array(arity, fnAry)
End Function

Function zipApplyF(fnObj, ParamArray argAry())
    Dim arys, x, ret
    arys = argAry
    x = zipAry(arys)
    ret = mapA("applyF", x, fnObj, True)
    zipApplyF = ret
End Function

Sub setAryMByF(ary, fnObj)
    Dim aNum As Long
    Dim i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aNum = getAryNum(ary)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = applyF(i, fnObj)
        Call setElm(vl, ary, idx)
    Next i
End Sub

Function takeWhile(fnc As String, ary, dir As Direction, ParamArray argAry())
    Dim lNum As Long, sn As Long, i As Long, num As Long
    Dim prm, v, ret, fnAry
    prm = argAry
    fnAry = prmAry(fnc, Empty, prm)
    lNum = lenAry(ary)
    sn = Sgn(dir)
    num = 0
    For i = 1 To lNum
        v = getAryAt(ary, sn * i)
        Call setAryAt(fnAry, 1, v, 0)
        If evalA(fnAry) Then
            num = num + 1
        Else
            Exit For
        End If
    Next
    ret = takeAry(ary, num, dir)
    takeWhile = ret
End Function

Function dropWhile(fnc As String, ary, dir As Direction, ParamArray argAry())
    Dim lNum As Long, sn As Long, i As Long, num As Long
    Dim prm, v, ret, fnAry
    prm = argAry
    fnAry = prmAry(fnc, Empty, prm)
    lNum = lenAry(ary)
    sn = Sgn(dir)
    num = 0
    For i = 1 To lNum
        v = getAryAt(ary, sn * i)
        Call setAryAt(fnAry, 1, v, 0)
        If evalA(fnAry) Then
            num = num + 1
        Else
            Exit For
        End If
    Next
    ret = dropAry(ary, num, dir)
    dropWhile = ret
End Function
