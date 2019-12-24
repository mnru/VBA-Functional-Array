Attribute VB_Name = "modWdw"
Function mkWdw(ary)
    Dim ret
    n = lenAry
    aryW = Array("wdw", 0, n - 1, 1)
    ret = Array(aryW, ary)
    mkWdw = ret

End Function

Function wdwToAry(wdw)

    

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

                    
