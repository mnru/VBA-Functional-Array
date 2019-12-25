Attribute VB_Name = "modTest"
'"Attribute VB_Name = "modTest"
Sub testconAry()
    ary1 = Array(1, 2, 3)
    ary2 = Array(4, 5, 6, 7)
    ary3 = Array(8, 9, 10)
    x1 = conArys(ary1, ary2)
    x2 = conArys(ary1, ary2, ary3)
    printAry x1
    printAry x2
    Stop
    
End Sub
Function conStr(a, b, dlm)
    conStr = a & dlm & b
End Function
Sub testReduce()
    x = reduceA("conStr", Array("a", "b", "c"), "-")
    Debug.Print x
End Sub
Sub testFold()
    x = foldA("calc", mkSeq(5), 100, "-")
    Debug.Print x
End Sub
Sub testCollection()
    Dim cll               As Collection
    Set cll = New Collection
    ary1 = Array(1, 2, 3)
    ary2 = Array("a", "b", "c")
    For Each elm In ary1
        cll.Add elm
    Next
    For Each elm In ary2
        cll.Add elm
    Next
    x = clcToAry(cll)
    printAry x
End Sub
Sub testSeq()
    printAry mkSameAry(12, 5)
    printAry mkSeq(1, 5)
    printAry mkSeq(5, 3)
    printAry mkSeq(15, 3, 2)
    printAry mkSeq(5, 9, 2)
    printAry mkSeq(3, -3, -2)
End Sub
Sub testToString()
    a = "abc"
    b = Time
    c = "123"
    x = Array(1, Array(1 / 3, 2.5, 3), Array(2, Array(3)), Array(1, 2))
    y = Array(Array(True, False, True), Array(4, 5, 6))
    Z = Application.WorksheetFunction.Transpose(y)
    w = Range("A1:C2")
    Debug.Print toString(a)
    Debug.Print toString(b)
    Debug.Print toString(c)
    Debug.Print toString(x)
    Debug.Print toString(y)
    Debug.Print toString(Z)
    Debug.Print toString(w)
End Sub
Sub testDrop()
    ary = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim ary1(1 To 9)
    For i = 1 To 9
        ary1(i) = i
    Next
    x = dropAry(ary, 3)
    y = dropAry(ary, 0)
    Z = dropAry(ary, -3)
    w = dropAry(ary, 9)
    x1 = dropAry(ary1, 3)
    y1 = dropAry(ary1, 0)
    z1 = dropAry(ary1, -3)
    w1 = dropAry(ary1, 9)
    printAry (x)
    printAry (y)
    printAry (Z)
    printAry (w)
    printAry (x1)
    printAry (y1)
    printAry (z1)
    printAry (w1)
End Sub
Sub testTake()
    ary = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim ary1(1 To 9)
    For i = 1 To 9
        ary1(i) = i
    Next
    y = takeAry(ary, 0)
    Z = takeAry(ary, 3)
    w = takeAry(ary, -3)
    y1 = takeAry(ary1, 0)
    z1 = takeAry(ary1, 3)
    w1 = takeAry(ary1, -3)
    printAry (y)
    printAry (Z)
    printAry (w)
    printAry (y1)
    printAry (z1)
    printAry (w1)
End Sub
Sub testCon()
    a = mkSeq(10000)
    b = mkSeq(20000, 0, 2)
    c = mkSeq(30000, 0, 3)
    Call printTime("conarys", a, b, c)
End Sub
Sub testMapA()
    a = mkSeq(10)
    b = mkSeq(0, 30000, 3)
    x = printTime("mapA", "calc", a, 1, "+")
    printAry (x)
    y = printTime("mapA", "calc", a, 2, "*")
    printAry (y)
    t1 = Time
    z0 = mapA("calc", b, 1, "+")
    t2 = Time
    Debug.Print "mapA: -" & Format(t2 - t1, "hh:nn:ss")
End Sub
Sub testRgt()
    rg = Range("A1:A2")
    Debug.Print TypeName(rg)
    Debug.Print IsArray(rg)
End Sub
Sub testRangeToArys()
    Dim rg                As Range
    Set rg = Range("A1:C2")
    dary = rg
    Dim dr(0 To 1, 0 To 2)
    dr(0, 0) = "a"
    dr(0, 1) = "b"
    dr(0, 2) = "c"
    dr(1, 0) = 1
    dr(1, 1) = 2
    dr(1, 2) = 3
    a = rangeToArys(rg)
    b = rangeToArys(rg, "c")
    Ad = rangeToArys(dary)
    bd = rangeToArys(dary, "c")
    Adr = rangeToArys(dr)
    bdr = rangeToArys(dr, "c")
    printAry (a)
    printAry (b)
    printAry (Ad)
    printAry (bd)
    printAry (Adr)
    printAry (bdr)
End Sub
Sub testElm()
    Dim a(0 To 1, 0 To 2, 0 To 3, 0 To 4)
    vl = 1
    For i = 0 To 1
        For j = 0 To 2
            For k = 0 To 3
                For l = 0 To 4
                    a(i, j, k, l) = vl
                    vl = vl + 1
                Next l
            Next k
        Next j
    Next i
    x = getElm(a, Array(0, 1, 2, 3))
    Debug.Print x
    sp = getAryShape(a)
    lsp = getAryShape(a, "L")
    n = reduceA("calc", sp, "*")
    For i = 0 To n - 1
        idx = mkIndex(i, sp, lsp)
        y = getElm(a, idx)
        Debug.Print y & ","
    Next i
End Sub
Sub testRangeToAry()
    Dim rg                As Range
    Set rg = Range("A1:C2")
    dary = rg
    Dim dr(0 To 1, 0 To 2)
    dr(0, 0) = "a"
    dr(0, 1) = "b"
    dr(0, 2) = "c"
    dr(1, 0) = 1
    dr(1, 1) = 2
    dr(1, 2) = 3
    a = rangeToAry(rg, "r", 2)
    b = rangeToAry(rg, "c", 2)
    Ad = rangeToAry(dary, "r", 2)
    bd = rangeToAry(dary, "c", 2)
    Adr = rangeToAry(dr, "r", 2)
    bdr = rangeToAry(dr, "c", 2)
    printAry (a)
    printAry (b)
    printAry (Ad)
    printAry (bd)
    printAry (Adr)
    printAry (bdr)
End Sub
Sub testAt()
    a = Array(1, 2, 3, 4, 5, 6)
    Dim b(1 To 6)
    For i = 1 To 6
        b(i) = i
    Next
    Debug.Print getAryAt(a, 2)
    Debug.Print getAryAt(b, 2)
    Debug.Print getAryAt(a, -2)
    Debug.Print getAryAt(b, -2)
    Call setAryAt(a, 2, -3)
    Call setAryAt(b, 2, -3)
    Call setAryAt(a, -2, -5)
    Call setAryAt(b, -2, -5)
    printAry (a)
    printAry (b)
    a = Array(1, 2, 3, 4, 5, 6)
    For i = 1 To 6
        b(i) = i
    Next
    Debug.Print getAryAt(a, 2, 0)
    Debug.Print getAryAt(b, 2, 0)
    Debug.Print getAryAt(a, -2, 0)
    Debug.Print getAryAt(b, -2, 0)
    Call setAryAt(a, 2, -3, 0)
    Call setAryAt(b, 2, -3, 0)
    Call setAryAt(a, -2, -5, 0)
    Call setAryAt(b, -2, -5, 0)
    printAry (a)
    printAry (b)
End Sub
Sub testadd()
    x = Array(Null, Null)
    Debug.Print lenAry(x)
    printAry (x)
End Sub
Sub testShape()
    Dim a(1 To 3, 1 To 4, 1 To 5)
    ' Dim a(1 To 3, 1 To 4)
    Dim b(0 To 3, 0 To 4, 0 To 5)
    Dim c(1 To 5)
    vl = 1
    fob = mkF(1, "calc", Null, 1, "+")
    Call setAryByF(a, fob)
    Call setAryByF(b, fob)
    Call setAryByF(c, fob)
    x = getAryShape(a)
    y = getAryShape(b)
    Z = getAryShape(c)
    printAry (x)
    printAry (y)
    printAry (Z)
    x = getAryShape(a, "U")
    y = getAryShape(b, "U")
    Z = getAryShape(c, "U")
    printAry (x)
    printAry (y)
    printAry (Z)
    x = getAryShape(a, "L")
    y = getAryShape(b, "L")
    Z = getAryShape(c, "L")
    printAry (x)
    printAry (y)
    printAry (Z)
    Call printTime("printAry", a)
    Call printTime("printAry", b)
    Call printTime("printAry", c)
    Stop
End Sub
Sub testApply()
    a = mkSeq(30)
    e = mapA("applyF", a, mkF(2, "calc", 2, Null, "^"))
    printAry (e)
    fob0 = Array(Array(Array(1), Array("calc", Null, 3, "+")), Array(Array(2), Array("calc", 100, Null, "/")))
    fob1 = Array(mkF(1, "calc", Null, 3, "+"), mkF(2, "calc", 100, Null, "/"))
    b0 = mapA("applyFs", a, fob0)
    b1 = mapA("applyFs", a, fob1)
    printAry (b0)
    printAry (b1)
End Sub
Sub testmkF()
    a = mkF(1, "calc", Null, 3, "%")
    b = mkF(2, 1, "calc", Null, Null, "-")
    printAry (a)
    printAry (b)
End Sub
Sub testPrmAry()
    a = Array(1, 2, 3, Array(4, 5, 6), Array(7, 8, 9))
    b = prmAry(a)
    printAry (a)
    printAry (b)
End Sub
Sub testFoldF()
    fo = mkF(2, 1, "calc", Null, Null, "-")
    sq = mkSeq(5)
    a = foldF(fo, sq, 1)
    Debug.Print a
End Sub
Sub testZipApply()
    fob = mkF(1, 2, "calc", Null, Null, "+")
    Z = zipApplyF(fob, mkSeq(5), mkSeq(10, 2, 2))
    printAry (Z)
End Sub
Sub testZip()
    x = zip(Array(1, 2, 3, 4), Array(2, 3, 4, 5), Array(3, 4, 5, 6))
    printAry (x)
    a = mkSeq(5)
    b = mkSeq(10, 2, 2)
    c = zip(a, b)
    printAry (c)
    d = Array(Array(1, 2), Array(3, 4), Array(5, 6))
    e = Array(7, 8, 9)
    f = Array("a", "b", "c")
    y = zip(d, e, f)
    printAry (y)
    Stop
End Sub
Sub testAry()
    Dim x(1 To 3, 1 To 3) As String
    For i = 1 To 3
        For j = 1 To 3
            x(i, j) = Chr(65 + (i - 1) + (j - 1) * 3)
        Next j
    Next i
    Set y = Range("a1:c2")
    Z = Range("a1:c2")
    Debug.Print TypeName(x)
    printAry (x)
    Debug.Print TypeName(y)
    printAry (y)
    Debug.Print TypeName(Z)
    printAry (Z)
End Sub
Sub testZipArrayTime()
    a = mkSeq(10)
    b = mkSeq(20, 2, 2)
    c = mkSeq(100, 10, 10)
    a = mkSeq(100000)
    b = mkSeq(200000, 2, 2)
    c = mkSeq(1000000, 10, 10)
    x = Array(a, b, c)
    y = printTime("zipary", x)
    Z = printTime("zip", a, b, c)
    Call printTime("conarys", x)
    ' printAry x
    ' printAry y
    ' printAry Z
    'Stop
End Sub
