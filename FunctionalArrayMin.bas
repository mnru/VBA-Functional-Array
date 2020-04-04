Attribute VB_Name = "FunctionalArrayMin"

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
    'ret = reduceA("calc", sp, "*")
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

Function calc(num1 As Variant, num2 As Variant, symbol As String)
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
    Dim lNum As Long, i As Long
    lNum = lenAry(ary1)
    ReDim ret(0 To lNum - 1)
    For i = 0 To lNum - 1
        ret(i) = calc(getAryAt(ary1, i, 0), getAryAt(ary2, i, 0), symbol)
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

