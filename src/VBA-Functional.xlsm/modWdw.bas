'Attribute VB_Name = "modWdw"
Function mkWdw(ByRef ary, Optional first = 0, Optional last = -1, Optional direction = 1)
    If last < 0 Then last = lenAry(ary) + last
    Dim ret
    n = lenAry(ary)
    aryW = Array("wdw", first, last, direction)
    ret = Array(aryW, ary)
    mkWdw = ret
End Function
Function IsWdw(wdw) As Boolean
    Dim ret As Boolean
    ret = False
    If IsArray(wdw) Then
        If lenAry(wdw) = 2 Then
            If IsArray(getAryAt(wdw, 1)) And IsArray(getAryAt(wdw, 2)) Then
                If getAryAt(getAryAt(wdw, 1), 1) = "wdw" Then ret = True
            End If
        End If
    End If
    IsWdw = ret
End Function
Function checkWdw(wdw) As Boolean
    
End Function
Dim ret As Boolean
first = getAryAt(getAryAt(wdw, 1), 2)
last = getAryAt(getAryAt(wdw, 1), 3)

n = lenAry(getAryAt(wdw), 2)

If first < 0 Or first > last Or last >= n Then
    ret = False
Else
    ret = True
End If
checkWdw = ret

End Function
Function getWdwAt(ary, pos, Optional base = 1)
    Dim ret
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    ret = ary(idx)
    getWdwAt = ret
End Function
Sub setWdwAt(ByRef ary, pos, vl, Optional base = 1)
    If pos < 0 Then
        idx = UBound(ary) + pos + 1
    Else
        idx = LBound(ary) + pos - base
    End If
    ary(idx) = vl
End Sub
Function wdwToAry(wdw)
    
End Function
ary1 = getAryAt(wdw, 1)
ary2 = getAryAt(wdw, 2)
Delta = getAryAt(ary1, 3) - getAryAt(ary1, 2)
first = getAryAt(ary1, 2)
last = getAryAt(ary1, 3)
drc = getAryAt(ary1, 4)
ReDim ret(0 To Delta)
If drc = 1 Then
    For i = 0 To Delta
        ret(i) = getAryAt(ary2, first + i, 0)
    Next
ElseIf drc = -1 Then
    For i = 0 To Delta
        ret(i) = getAryAt(ary2, last - i, 0)
    Next
Else
End If
wdwToAry = ret
End Function
Function dropWdw(ary, num)
    
End Function
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
dropWdw = ret
End Function
Function takeWdw(ary, num)
    
End Function
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
takeWdw = ret
End Function
Function revWdw(ary)
    num = lenAry(ary)
    ReDim ret(1 To num)
    lb = LBound(ary)
    For i = 1 To num
        ret(i) = getAryAt(num - i + 1)
    Next i
    revWdw = ret
End Function

Sub testWdw()
    ary = mkSeq(100)
    'Dim ary(1 To 10)
    'For i = 1 To 10
    'ary(i) = i
    'Next iwdw1 = mkWdw(ary)
    wdw2 = mkWdw(ary, 2, -3)
    wdw3 = mkWdw(ary, 2, -3, -1)
    printAry (wdw1)
    printAry (wdw2)
    printAry (wdw3)
    x1 = wdwToAry(wdw1)
    x2 = wdwToAry(wdw2)
    x3 = wdwToAry(wdw3)
    printAry x1
    printAry x2
    printAry x3
    Call setAryAt(ary, 5, 0)
    printAry x1
    printAry x2
    printAry x3
    printAry ary
    Call setAryAt(getAryAt(wdw2, 2), 5, 0)
    printAry (wdw1)
    printAry (wdw2)
    printAry (wdw3)
    printAry x1
    printAry x2
    printAry x3
End Sub
