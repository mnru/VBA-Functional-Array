Attribute VB_Name = "modEq"
Sub testEqAry()
    n = 10
    x1 = mkMArySeq(Array(n), , , 0)
    x2 = mkMArySeq(Array(n), , , 0)
    x3 = mkMArySeq(Array(n), , , 1)
    x4 = mkMArySeq(Array(n), , 2, 0)
    x5 = mkMArySeq(Array(n, n), , 2, 0)
    x6 = 1
    printOut (eqAry(x1, x2))
    printOut (eqAry(x1, x3))
    printOut (eqAry(x1, x4))
    printOut (eqAry(x1, x5))
    printOut (eqAry(x1, x6))
End Sub

Function areBothAry(ary1, ary2) As Boolean
    Dim ret As Boolean
    ret = IIf(IsArray(ary1) And IsArray(ary2), True, False)
    areBothAry = ret
End Function

Function haveSameDimension(ary1, ary2) As Boolean
    Dim ret As Boolean
    ret = (dimAry(ary1) = dimAry(ary2))
    haveSameDimension = ret
End Function

Function hasDimensionOne(ary1) As Boolean
    Dim ret As Boolean
    ret = (dimAry(ary1) = 1)
    hasDimensionOne = ret
End Function

Function areSameLength(ary1, ary2) As Boolean
    Dim ret As Boolean
    ret = (lenAry(ary1) = lenAry(ary2))
    areSameLength = ret
End Function

Function eqAry(ary1, ary2) As Boolean
    Dim ret As Boolean
    ret = False
    If Not IsArray(ary1) Or Not IsArray(ary2) Then
        printOut ("at least one of two is not array")
    ElseIf dimAry(ary1) > 1 Or dimAry(ary2) > 1 Then
        printOut ("at least one of two has higher dimension than one")
    Else
        n1 = lenAry(ary1)
        n2 = lenAry(ary2)
        If n1 <> n2 Then
            printOut ("two arrays has different length")
        Else
            For i = 1 To n1
                If getAryAt(ary1, i) <> getAryAt(ary2, i) Then
                    printOut ("at least one element has different value")
                    Exit For
                End If
                If i = n1 Then
                    ret = True
                    If LBound(ary1) = LBound(ary2) Then
                        printOut ("two arrays have same value and same index")
                    Else
                        printOut ("two arrays have same value but different index")
                    End If
                End If
            Next i
        End If
    End If
    eqAry = ret
End Function

Function eqShape(ary1, ary2) As Boolean
    Dim ret As Boolean
    sp1 = getAryShape(ary1)
    sp2 = getAryShape(ary2)
    ret = eqAry(sp1, sp2)
    eqShape = ret
End Function
