Attribute VB_Name = "modSub"
'''''''''''''''''''''''''''''''''''
'sub of modAry
'''''''''''''''''''''''''''''''''''

Sub set1DArySeq(ary, Optional first = 1, Optional step = 1)
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    vl = first
    For i1 = lb1 To ub1
        ary(i1) = vl
        vl = vl + step
    Next i1
End Sub

Sub set2DArySeq(ary, Optional first = 1, Optional step = 1)
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    vl = first
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            ary(i1, i2) = vl
            vl = vl + step
        Next i2
    Next i1
End Sub

Sub set3DArySeq(ary, Optional first = 1, Optional step = 1)
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    lb3 = LBound(ary, 3): ub3 = UBound(ary, 3)
    vl = first
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            For i3 = lb3 To ub3
                ary(i1, i2, i3) = vl
                vl = vl + step
            Next i3
        Next i2
    Next i1
End Sub

Function mk2DSeq(r, c, Optional first = 1, Optional step = 1, Optional bs = 0)
    ReDim ret(bs To bs + r - 1, bs To bs + c - 1)
    vl = first
    For i1 = bs To bs + r - 1
        For i2 = bs To bs + c - 1
            ret(i1, i2) = vl
            vl = vl + step
        Next i2
    Next i1
    mk2DSeq = ret
End Function

Function mk3DSeq(r, c, h, Optional first = 1, Optional step = 1, Optional bs = 0)
    ReDim ret(bs To bs + r - 1, bs To bs + c - 1, bs To bs + h - 1)
    vl = first
    For i1 = bs To bs + r - 1
        For i2 = bs To bs + c - 1
            For i3 = bs To bs + h - 1
                ret(i1, i2, i3) = vl
                vl = vl + step
            Next i3
        Next i2
    Next i1
    mk3DSeq = ret
End Function
