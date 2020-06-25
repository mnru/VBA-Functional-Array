Attribute VB_Name = "FunctionalArraySelected"

Option Base 0
Option Explicit

''''''''''''''''''''
' enum
''''''''''''''''''''
Enum Direction
    faDirect = 1
    faReverse = -1
End Enum

Enum shapeType
    faNormal = 0
    faLower = 1
    faUpper = 2
End Enum

Enum Aligned
    faLeft = 1
    faRight = -1
    faCenter = 0
End Enum

Enum rowColumn
    faRow = 1
    faColumn = 2
End Enum


''''''''''''''''''''
'from modAry
''''''''''''''''''''
Function lenAry(ary As Variant, Optional dm = 1) As Long
    lenAry = UBound(ary, dm) - LBound(ary, dm) + 1
End Function

Function getAryAt(ary As Variant, pos As Long, Optional base As Long = 1)
    Dim idx As Long
    Dim ret As Variant
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    assign_ getAryAt, ary(idx)
End Function

Sub setAryAt(ByRef ary As Variant, pos As Long, vl As Variant, Optional base As Long = 1)
    Dim idx As Long
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    assign_ ary(idx), vl
End Sub

Public Function dimAry(ByVal ary As Variant) As Long
    On Error GoTo Catch
    Dim ret As Long
    ret = 0
    Do
        ret = ret + 1
        Dim tmp As Long
        tmp = UBound(ary, ret)
    Loop
Catch:
    dimAry = ret - 1
End Function

Function getAryShape(ary, Optional spt As shapeType = faNormal)
    Dim i As Long, num As Long
    Dim tmp
    num = dimAry(ary)
    ReDim ret(0 To num - 1)
    For i = 1 To num
        Select Case spt
            Case faNormal
                tmp = lenAry(ary, i)
            Case faLower
                tmp = LBound(ary, i)
            Case faUpper
                tmp = UBound(ary, i)
            Case Else
        End Select
        Call setAryAt(ret, i, tmp)
    Next i
    getAryShape = ret
End Function

Function getAryNum(ary) As Long
    Dim ret As Long
    Dim sp, elm
    sp = getAryShape(ary)
    'ret = reduceA("calc_", sp, "*")
    ret = 1
    For Each elm In sp
        ret = ret * elm
    Next elm
    getAryNum = ret
End Function

Function conArys(ParamArray argArys())
    Dim num As Long, i As Long
    Dim arys, ret, elm, ary
    arys = argArys
    num = 0
    For Each ary In arys
        If IsArray(ary) Then
            num = num + getAryNum(ary)
        Else
            num = num + 1
        End If
    Next ary
    ReDim ret(0 To num - 1)
    i = 0
    For Each ary In arys
        If IsArray(ary) Then
            For Each elm In ary
                ret(i) = elm
                i = i + 1
            Next elm
        Else
            ret(i) = ary
            i = i + 1
        End If
    Next ary
    conArys = ret
End Function

Function mkSameAry(vl, lNum As Long)
    Dim i As Long
    ReDim ret(0 To lNum - 1)
    For i = 0 To lNum - 1
        ret(i) = vl
    Next i
    mkSameAry = ret
End Function

Function mkSeq(lNum As Long, Optional first = 1, Optional step = 1)
    ReDim ret(0 To lNum - 1)
    Dim i As Long
    For i = 0 To lNum - 1
        ret(i) = first + step * i
    Next i
    mkSeq = ret
End Function

Function takeAry(ary, num As Long, Optional dir As Direction = faDirect)
    Dim lNum As Long, i As Long, lb As Long
    Dim ret
    lNum = lenAry(ary)
    If lNum < num Then
        Call Err.Raise(1001, "takeAry", "num is larger than array length")
    ElseIf dir = 0 Then
        Call Err.Raise(1001, "takeAry", "faCenter is not valid")
    End If
    If num <= 0 Then
        ret = Array()
    Else
        Select Case dir
            Case faDirect
                ReDim ret(0 To num - 1)
                lb = LBound(ary)
                For i = 0 To num - 1
                    ret(i) = getAryAt(ary, i, 0)
                Next i
            Case faReverse
                ReDim ret(0 To num - 1)
                ' ub = UBound(ary)
                For i = 0 To num - 1
                    ret(i) = getAryAt(ary, lNum - num + i, 0)
                Next i
        End Select
    End If
    takeAry = ret
End Function

Function dropAry(ary, num As Long, Optional dir As Direction = faDirect)
    Dim lNum As Long, i As Long, lb As Long, ub As Long
    Dim ret
    lNum = lenAry(ary)
    If lNum < num Then
        Call Err.Raise(1001, "dropAry", "num is larger than array length")
    ElseIf dir = 0 Then
        Call Err.Raise(1001, "takeAry", "faCenter is not valid")
    Else
        ret = takeAry(ary, lNum - num, -1 * dir)
    End If
    dropAry = ret
End Function

Function revAry(ary)
    Dim lNum As Long, i As Long, lb As Long
    lNum = lenAry(ary)
    ReDim ret(0 To lNum - 1)
    lb = LBound(ary)
    For i = 0 To lNum - 1
        ret(i) = getAryAt(ary, lNum - i - 1, 0)
    Next i
    revAry = ret
End Function

Function zip(ParamArray argArys())
    Dim arys, ret
    arys = argArys
    ret = zipAry(arys)
    zip = ret
End Function

Function zipAry(arys)
    Dim rNum As Long, cNum As Long, lb As Long, c As Long, r As Long
    Dim ret, v
    rNum = lenAry(arys)
    cNum = lenAry(arys(LBound(arys)))
    ReDim ret(0 To cNum - 1)
    lb = LBound(arys)
    For c = 0 To cNum - 1
        ReDim v(0 To rNum - 1)
        For r = 0 To rNum - 1
            v(r) = getAryAt(arys(lb + r), c, 0)
        Next r
        ret(c) = v
    Next c
    zipAry = ret
End Function

Function zipWithIndex(ary, Optional first As Long = 1, Optional step As Long = 1)
    Dim ret, aryI
    Dim lNum As Long
    lNum = lenAry(ary)
    aryI = mkSeq(lNum, first, step)
    ret = zip(ary, aryI)
    zipWithIndex = ret
End Function

Function prmAry(ParamArray argAry())
    Dim ret
    'flatten last elm
    Dim ary, ary1, ary2
    ary = argAry
    ary1 = dropAry(ary, 1, faRight)
    ary2 = getAryAt(ary, -1)
    ret = conArys(ary1, ary2)
    prmAry = ret
End Function

Function inAry(ary As Variant, elm As Variant) As Boolean
    Dim ret As Boolean
    ret = False
    For Each x In ary
        If x = elm Then
            ret = True
            Exit For
        End If
    Next x
    inAry = ret
End Function

Function calc_(num1 As Variant, num2 As Variant, symbol As String)
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
    calc_ = ret
End Function

Function calcAry(ary1, ary2, symbol As String)
    Dim lNum As Long, i As Long
    lNum = lenAry(ary1)
    ReDim ret(0 To lNum - 1)
    For i = 0 To lNum - 1
        ret(i) = calc_(getAryAt(ary1, i, 0), getAryAt(ary2, i, 0), symbol)
    Next i
    calcAry = ret
End Function

Function uniqueAry(ary)
    Dim dic As Dictionary
    Set dic = CreateObject("Scripting.Dictionary")
    For Each elm In ary
        If Not dic.exists(elm) Then
            dic.Add elm, Empty
        End If
    Next
    ret = dic.keys
    uniqueAry = ret
End Function

Sub assign_(ByRef Variable As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Variable = Value
    Else
        Variable = Value
    End If
End Sub

''''''''''''''''''''
'from modFnc
''''''''''''''''''''
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

Private Function foldAryA(fnc As String, seq As Variant, init As Variant, ary) As Variant
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

''''''''''''''''''''
'from modMulti
''''''''''''''''''''
Function getMAryAt(ary As Variant, pos As Variant, Optional base As Long = 1)
    Dim lNum As Long
    Dim lsp, bs, idx1, idx2, ret
    lsp = getAryShape(ary, faLower)
    lNum = lenAry(lsp)
    bs = mkSameAry(base, lNum)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    ret = getElm(ary, idx2)
    getMAryAt = ret
End Function

Sub setMAryAt(ByRef ary As Variant, pos As Variant, vl As Variant, Optional base As Long = 1)
    Dim lNum As Long
    Dim lsp, bs, idx1, idx2
    lsp = getAryShape(ary, faLower)
    lNum = lenAry(lsp)
    bs = mkSameAry(base, lNum)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    Call setElm(vl, ary, idx2)
End Sub

Function mkIndex(num As Long, shape, Optional lshape = Empty)
    Dim lNum As Long, i As Long, p As Long, r As Long
    lNum = lenAry(shape)
    ReDim ret(0 To lNum - 1)
    r = num
    For i = lNum To 1 Step -1
        p = getAryAt(shape, i)
        Call setAryAt(ret, i, r Mod p)
        r = r \ p
    Next i
    If Not IsEmpty(lshape) Then
        For i = 1 To lNum
            Call setAryAt(ret, i, getAryAt(ret, i) + getAryAt(lshape, i))
        Next i
    End If
    mkIndex = ret
End Function

Function getElm(ByRef ary, idx)
    Dim ret
    Dim lb As Long
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: assign_ ret, ary(idx(lb))
        Case 2: assign_ ret, ary(idx(lb), idx(lb + 1))
        Case 3: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2))
        Case 4: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3))
        Case 5: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4))
        Case Else:
    End Select
    assign_ getElm, ret
End Function

Sub setElm(vl, ary, idx)
    Dim lb As Long
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: assign_ ary(idx(lb)), vl
        Case 2: assign_ ary(idx(lb), idx(lb + 1)), vl
        Case 3: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2)), vl
        Case 4: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3)), vl
        Case 5: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4)), vl
        Case Else:
    End Select
End Sub

Sub setAryMbyS(mAry, sAry)
    Dim i As Long, aNum As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
    aNum = getAryNum(mAry)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = getAryAt(sAry, i, 0)
        Call setElm(vl, mAry, idx)
    Next i
End Sub

Function getArySbyM(mAry, Optional bs As Long = 0)
    Dim aNum As Long, i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
    aNum = getAryNum(mAry)
    ReDim ret(bs To bs + aNum - 1)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = getElm(mAry, idx)
        Call setAryAt(ret, i, vl, 0)
    Next i
    getArySbyM = ret
End Function

Function reshapeAry(ary, sp, Optional bs As Long = 0)
    Dim ret
    ret = mkMAry(sp, bs)
    Call setAryMbyS(ret, ary)
    reshapeAry = ret
End Function

Function calcMAry(ary1, ary2, symbol As String, Optional bs As Long = 0)
    Dim aNum As Long, dm As Long, i As Long
    Dim ret, vl, sp1, sp2, lsp1, lsp2, lsp0, idx, idx0, idx1, idx2
    sp1 = getAryShape(ary1)
    sp2 = getAryShape(ary2)
    lsp1 = getAryShape(ary1, faLower)
    lsp2 = getAryShape(ary2, faLower)
    aNum = getAryNum(ary1)
    dm = lenAry(sp1)
    ret = mkMAry(sp1, bs)
    lsp0 = mkSameAry(bs, dm)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp1)
        idx1 = calcAry(idx, lsp1, "+")
        idx2 = calcAry(idx, lsp2, "+")
        idx0 = calcAry(idx, lsp0, "+")
        vl = calc_(getElm(ary1, idx1), getElm(ary2, idx2), symbol)
        Call setElm(vl, ret, idx0)
    Next i
    calcMAry = ret
End Function

Function mkMAry(sp, Optional bs As Long = 0)
    Dim lNum As Long
    Dim ub, lb
    lNum = lenAry(sp)
    ub = calcAry(sp, mkSameAry(bs - 1, lNum), "+")
    lb = LBound(ub)
    Dim ret
    Select Case lNum
        Case 1: ReDim ret(bs To ub(lb))
        Case 2: ReDim ret(bs To ub(lb), bs To ub(lb + 1))
        Case 3: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2))
        Case 4: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3))
        Case 5: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4))
        Case Else:
    End Select
    mkMAry = ret
End Function

Sub setMArySeq(ary, Optional first = 1, Optional step = 1)
    Dim aNum As Long, i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, faLower)
    aNum = getAryNum(ary)
    vl = first
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        'vl = first + i0 * step
        Call setElm(vl, ary, idx)
        vl = vl + step
    Next i
End Sub

Function mkMSameAry(vl, sp, Optional bs As Long = 0)
    Dim lNum As Long, aNum As Long, i As Long
    Dim ret, lsp, idx
    ret = mkMAry(sp, bs)
    lNum = lenAry(sp)
    sp = getAryShape(ret)
    lsp = mkSameAry(bs, lNum)
    aNum = getAryNum(ret)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        Call setElm(vl, ret, idx)
    Next i
    mkMSameAry = ret
End Function

Function mkMSeq(sp, Optional first = 1, Optional step = 1, Optional bs As Long = 0)
    Dim ret
    ret = mkMAry(sp, bs)
    Call setMArySeq(ret, first, step)
    mkMSeq = ret
End Function

Public Function mapMA(fnc As String, mAry As Variant, ParamArray argAry() As Variant) As Variant
    Dim ary, sp, lsp, fnAry, ret, idx, idx0, vl
    Dim aNum As Long
    ary = argAry
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
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

Sub setAryMByF(ary, fnObj)
    Dim aNum As Long
    Dim i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, faLower)
    aNum = getAryNum(ary)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = applyF(i, fnObj)
        Call setElm(vl, ary, idx)
    Next i
End Sub

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

''''''''''''''''''''
'from modUtil
''''''''''''''''''''
Function toString(elm, Optional qt As String = "'", Optional fm As String = "", _
    Optional lcr As Aligned = faRight, Optional width As Long = 0, _
    Optional insheet As Boolean = False) As String
    Dim ret As String, tmp As String
    Dim i As Long, aNum As Long
    Dim sp, lsp, idx, idx0, vl, dlm
    ret = ""
    If IsArray(elm) Then
        ret = ret & "["
        sp = getAryShape(elm)
        lsp = getAryShape(elm, faLower)
        aNum = getAryNum(elm)
        If aNum = 0 Then
            ret = ret & "]"
        Else
            For i = 0 To aNum - 1
                idx0 = mkIndex(i, sp)
                idx = calcAry(idx0, lsp, "+")
                vl = getElm(elm, idx)
                dlm = getDlm(sp, idx0, insheet)
                ret = ret & toString(vl, qt, fm, lcr, width) & dlm
            Next i
        End If
    ElseIf IsObject(elm) Then
        If TypeName(elm) = "Dictionary" Then
            ret = ret & dicToStr(elm)
        ElseIf TypeName(elm) = "Collection" Then
            ret = ret & clcToStr(elm)
        Else
            ret = ret & "<" & TypeName(elm) & ">"
        End If
    ElseIf IsNull(elm) Then
        ret = ret & "Null"
    ElseIf IsEmpty(elm) Then
        ret = ret & "Empty"
    Else
        If TypeName(elm) = "String" Then
            tmp = qt & elm & qt
        Else
            tmp = fmt(elm, fm)
        End If
        tmp = align(tmp, lcr, width)
        ret = ret & tmp
    End If
    toString = ret
End Function

Function dicToStr(dic) As String
    Dim tmp1, tmp2
    Dim ret As String
    tmp1 = zip(mapA("tostring", dic.keys), mapA("tostring", dic.items))
    tmp2 = mapA("mcJoin", tmp1, "=>")
    ret = mcJoin(tmp2, ",", "Dic(", ")")
    dicToStr = ret
End Function

Function clcToStr(clc) As String
    Dim tmp1, tmp2
    Dim ret As String
    tmp1 = clcToAry(clc)
    tmp2 = mapA("toString", tmp1)
    ret = mcJoin(tmp2, ",", "Clc(", ")")
    clcToStr = ret
End Function

Function getDlm(shape, idx, Optional insheet As Boolean = False) As String
    Dim ret As String
    Dim i As Long, lNum As Long, m As Long
    Dim nl As String
    nl = IIf(insheet, vbLf, vbCrLf)
    lNum = lenAry(shape)
    m = 0
    For i = lNum To 1 Step -1
        If getAryAt(shape, i) - 1 > getAryAt(idx, i) Then
            m = i
            Exit For
        End If
    Next i
    Select Case m
        Case 0
            ret = "]"
        Case lNum
            ret = ","
        Case lNum - 1
            ret = ";" & nl & " "
        Case Else
            ret = String(lNum - m, ";") & nl & nl & " "
    End Select
    getDlm = ret
End Function

Function secToHMS(vl As Double) As String
    Dim ret As String
    Dim x1 As Long
    Dim x0 As Double, x2 As Double, x3 As Double
    Dim tmp
    x0 = vl
    x1 = Int(x0)
    x2 = x0 - x1
    tmp = mkIndex(x1, Array(24, 60, 60))
    x3 = getAryAt(tmp, 3) + x2
    ret = Format(getAryAt(tmp, 1), "00") & ":" & Format(getAryAt(tmp, 2), "00") & ":" & Format(x3, "00.000")
    secToHMS = ret
End Function

Function clcToAry(clc)
    Dim cnt As Long, i As Long
    cnt = clc.Count
    ReDim ret(1 To cnt)
    For i = 1 To cnt
        assign_ ret(i), clc.Item(i)
    Next i
    clcToAry = ret
End Function

Function flattenAry(ary)
    Dim elm, el, ret
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
    Dim lb As Long, ub As Long, i As Long
    Dim ret
    lb = LBound(polyAry)
    ub = UBound(polyAry)
    ret = polyAry(lb)
    For i = lb + 1 To ub
        ret = ret * x + polyAry(i)
    Next
    poly = ret
End Function

Function polyStr(polyAry) As String
    Dim i As Long, lNum As Long
    Dim ret As String
    Dim c
    ret = ""
    lNum = lenAry(polyAry)
    For i = 1 To lNum
        c = getAryAt(polyAry, i)
        If c <> 0 Then
            If ret <> "" Then ret = ret & " "
            If c > 0 Then ret = ret & "+"
            If c <> 1 Or i = lNum Then ret = ret & c
            If i < lNum Then ret = ret & "X"
            If i < lNum - 1 Then ret = ret & "^" & lNum - i
        End If
    Next i
    If ret = "" Then ret = getAryAt(polyAry, -1)
    If Left(ret, 1) = "+" Then ret = Right(ret, Len(ret) - 1)
    polyStr = ret
End Function

Function fmt(expr, Optional fm As String = "", Optional lcr As Aligned = faRight, Optional width As Long = 0) As String
    Dim ret As String
    ret = Format(expr, fm)
    ret = align(ret, lcr, width)
    fmt = ret
End Function

Function align(str As String, Optional lcr As Aligned = faRight, Optional width As Long = 0) As String
    Dim ret As String
    Dim d As Long
    ret = CStr(str)
    d = width - Len(ret)
    If d > 0 Then
        Select Case lcr
            Case faRight: ret = space(d) & ret
            Case faLeft: ret = ret & space(d)
            Case faCenter: ret = space(d \ 2) & ret & space(d - d \ 2)
            Case Else:
        End Select
    End If
    align = ret
End Function

Function math_(x, symbol As String)
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
    math_ = ret
End Function

Function comp_(x, y, symbol As String)
    Dim ret
    Select Case symbol
        Case "=": ret = x = y                        'caution Assign_ and eqaul is same symbol
        Case "<>": ret = x <> y
        Case "<": ret = x < y
        Case ">": ret = x > y
        Case "<=": ret = x <= y
        Case ">=": ret = x >= y
        Case "<": ret = x < y
        Case Else:
    End Select
    comp_ = ret
End Function

Function info_(x, symbol As String)
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
    info_ = ret
End Function

Function id_(x)
    If IsObject(x) Then
        Set id_ = x
    Else
        id_ = x
    End If
End Function

Function mkDic(ParamArray argAry())
    Dim ary, ret
    ary = argAry
    Set ret = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = LBound(ary)
    Do While i < UBound(ary)
        ret.Add ary(i), ary(i + 1)
        i = i + 2
    Loop
    Set mkDic = ret
End Function

Function lookupDic(x, dic, Optional default = Empty)
    Dim ret
    If dic.exists(x) Then
        ret = dic(x)
    Else
        ret = default
    End If
    lookupDic = ret
End Function

Function mkClc(ParamArray argAry())
    Dim ary, elm, clc
    ary = argAry
    Set clc = New Collection
    For Each elm In ary
        clc.Add elm
    Next elm
    Set mkClc = clc
End Function

Function l_(ParamArray argAry() As Variant)
    'works like function array()
    Dim ary As Variant
    ary = argAry
    l_ = ary
End Function

''''''''''''''''''''
'from modRng
''''''''''''''''''''
Public Function TLookup(key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional otherwise = Empty) As Variant
    bkn = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    num = WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0)
    If num = 0 Then
        ret = otherwise
    Else
        ret = Range(tbl & "[" & targetCol & "]")(num, 1)
    End If
    TLookup = ret
    Workbooks(bkn).Activate
    Exit Function
lnError:
    Debug.Print Err.Description
    TLookup = Empty
    Workbooks(bkn).Activate
End Function

Public Sub TSetUp(vl, key, tbl As String, targetCol As String, Optional sourceCol As String = "")
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    Range(tbl & "[" & targetCol & "]")(WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0), 1).Value = vl
    Exit Sub
lnError:
    Debug.Print Err.Description
End Sub

Sub layAryAt(ary, r, c, Optional rc As rowColumn = rowColumn.faRow, Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ActiveWorkbook.Name
    n = lenAry(ary)
    Select Case rc
        Case rowColumn.faRow
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(1, n) = ary
        Case rowColumn.faColumn
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(n, 1) = Application.WorksheetFunction.Transpose(ary)
        Case Else
    End Select
End Sub

Function rangeToAry(rg, Optional rc As rowColumn = rowColumn.faRow, Optional num = 1)
    Dim ret, tmp
    tmp = rg
    If Not IsArray(tmp) Then
        ret = Array(tmp)
    Else
        With Application.WorksheetFunction
            Select Case rc
                Case rowColumn.faRow
                    ret = .Index(tmp, num, 0)
                Case rowColumn.faColumn
                    ret = .Transpose(.Index(tmp, 0, num))
                Case Else
            End Select
        End With
        If dimAry(ret) = 0 Then
            ret = Array(tmp)
        End If
    End If
    rangeToAry = ret
End Function

Function rangeToArys(rg, Optional rc As rowColumn = rowColumn.faRow)
    Dim ret, tmp
    Dim num As Long, i As Long
    tmp = rg
    If Not IsArray(tmp) Then
        ret = Array(tmp)
    Else
        If rc = rowColumn.faColumn Then
            tmp = Application.WorksheetFunction.Transpose(tmp)
        End If
        If dimAry(tmp) <= 1 Then
            ret = Array(tmp)
        Else
            num = lenAry(tmp)
            ReDim ret(1 To num)
            For i = 1 To num
                ret(i) = Application.WorksheetFunction.Index(tmp, i, 0)
            Next i
        End If
    End If
    rangeToArys = ret
End Function

''''''''''''''''''''
'from modLog
''''''''''''''''''''
Sub outPut(Optional msg As Variant = "", Optional crlf As Boolean = True)
    If crlf Then
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
End Sub

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim ary, fnAry
    Dim etime As Double
    Dim stime As Double
    Dim secs As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call outPut(fnc & " - " & secToHMS(secs), True)
End Function

Sub printAry(ary, Optional qt As String = "'", Optional fm As String = "", Optional lcr As Aligned = faRight, Optional width As Long = 0)
    Call outPut(toString(ary, qt, fm, lcr, width), True)
End Sub
