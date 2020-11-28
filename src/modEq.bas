Attribute VB_Name = "modEq"
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
    Dim n1 As Long, n2 As Long, i As Long
    ret = False
    If Not IsArray(ary1) Or Not IsArray(ary2) Then
        outPut ("at least one of two is not array")
    ElseIf dimAry(ary1) > 1 Or dimAry(ary2) > 1 Then
        outPut ("at least one of two has higher dimension than one")
    Else
        n1 = lenAry(ary1)
        n2 = lenAry(ary2)
        If n1 <> n2 Then
            outPut ("two arrays has different length")
        Else
            For i = 1 To n1
                If getAryAt(ary1, i) <> getAryAt(ary2, i) Then
                    outPut ("at least one element has different value")
                    Exit For
                End If
                If i = n1 Then
                    ret = True
                    If LBound(ary1) = LBound(ary2) Then
                        outPut ("two arrays have same value and same index")
                    Else
                        outPut ("two arrays have same value but different index")
                    End If
                End If
            Next i
        End If
    End If
    eqAry = ret
End Function

Function eqShape(ary1, ary2) As Boolean
    Dim ret As Boolean
    Dim sp1, sp2
    sp1 = getAryShape(ary1)
    sp2 = getAryShape(ary2)
    ret = eqAry(sp1, sp2)
    eqShape = ret
End Function

Function getEqLevel(ary1, ary2) As Long
    Dim n1 As Long, aNum As Long, i As Long
    Dim ret As Long
    Dim l1, l2, idx, idx1, idx2
    Dim bol As Boolean
    ret = 0
    If areBothAry(ary1, ary2) Then
        ret = 1
        If haveSameDimension(ary1, ary2) Then
            ret = 2
            If eqShape(ary1, ary2) Then
                ret = 3
                aNum = getAryNum(ary1)
                sp = getAryShape(ary1)
                l1 = getAryShape(ary1, "l")
                l2 = getAryShape(ary2, "l")
                bol = True
                For i = 0 To aNum - 1
                    idx = mkIndex(i, sp)
                    idx1 = calcAry(idx, l1, "+")
                    idx2 = calcAry(idx, l2, "+")
                    If getElm(ary1, idx1) <> getElm(ary2, idx2) Then
                        bol = False
                        Exit For
                    End If
                Next i
                If bol Then
                    If eqAry(l1, l2) Then
                        ret = 6
                    Else
                        ret = 5
                    End If
                Else
                    ret = 4
                End If
            End If
        End If
    End If
    getEqLevel = ret
End Function
