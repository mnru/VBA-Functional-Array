Attribute VB_Name = "modSort"
Function min2(x, y)
    min2 = IIf(x < y, x, y)
End Function

Function max2(x, y)
    max2 = IIf(x > y, x, y)
End Function

Function mid3(x, y, z)
    mid3 = IIf(max2(x, y) < z, max2(x, y), max2(min2(x, y), z))
End Function

Function qsortAry(ary, Optional l = -1, Optional r = -1)
    Dim l   As Long
    Dim r   As Long
    If l = -1 Then
        l = LBound(ary)
    End If
    If r = -1 Then
        r = UBound(ary)
    End If
    ret = ary
    Call qsortAry0(ret, l, r)
    qsortAry = ret
End Function

Sub qsortAry0(ary, l, r)
    Dim pivot, i, j
    pivot = ary(l + (r - l) \ 2)
    i = l: j = r
    Do
        Do While ary(i) < pivot
            i = i + 1
        Loop
        Do While ary(j) > pivot
            j = j - 1
        Loop
        If i >= j Then Exit Do
        tmp = ary(i): ary(i) = ary(j): ary(j) = tmp
    Loop
    If l < i - 1 Then Call qsortAry0(ary, l, i - 1)
    If j + 1 < r Then Call qsortAry0(ary, j + 1, r)
End Sub

Sub testSort()
    x = Array(4, 3, 9, 1, 2, 10, 7, 5, 6, 8)
    printAry (x)
    Call qsortAry0(x, 0, 9)
    printAry (x)
    x = Array(4, 3, 9, 1, 2, 10, 7, 5, 6, 8)
    y = qsortAry(x)
    printAry (x)
    printAry (y)
End Sub
