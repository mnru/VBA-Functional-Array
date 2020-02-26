Attribute VB_Name = "modCheck"
Sub checkconAry()
    ' DebugLog.setLog
    ary1 = Array(1, 2, 3)
    ary2 = Array(4, 5, 6, 7)
    ary3 = Array(8, 9, 10)
    x1 = conArys(ary1, ary2)
    x2 = conArys(ary1, ary2, ary3)
    printAry x1
    printAry x2
    Stop
End Sub

Function conStr(a, b, dlm): conStr = a & dlm & b: End Function

Sub checkReduce() 'comment for check
    x = reduceA("conStr", Array("a", "b", "c"), "-")
    outPut x
End Sub

Sub checkFold()
    x = foldA("calc", mkSeq(5), 100, "-")
    outPut x
End Sub

Sub checkCollection()
    Dim clc As Collection
    Set clc = New Collection
    ary1 = Array(1, 2, 3)
    ary2 = Array("a", "b", "c")
    For Each elm In ary1
        clc.Add elm
    Next
    For Each elm In ary2
        clc.Add elm
    Next
    x = clcToAry(clc)
    printAry x
End Sub

Sub checkSeq()
    printAry mkSameAry(12, 5)
    printAry mkSeq(5)
    printAry mkSeq(3, 5, -1)
    printAry mkSeq(15, 3, 2)
    printAry mkSeq(5, 9, 2)
    printAry mkSeq(3, -3, -2)
End Sub

Sub checkToString()
    a = "abc"
    b = Time
    c = "123"
    x = Array(1, Array(1 / 3, 2.5, 3), Array(2, Array(3)), Array(1, 2))
    y = Array(Array(True, False, True), Array(4, 5, 6))
    z = Application.WorksheetFunction.Transpose(y)
    w = Range("A1:C2")
    outPut toString(a)
    outPut toString(b)
    outPut toString(c)
    outPut toString(x)
    outPut toString(y)
    outPut toString(z)
    outPut toString(w)
End Sub

Sub checkDrop()
    ary = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim ary1(1 To 9)
    For i = 1 To 9
        ary1(i) = i
    Next
    x = dropAry(ary, 3)
    y = dropAry(ary, 0)
    z = dropAry(ary, 3, faReverse)
    w = dropAry(ary, 9)
    x1 = dropAry(ary1, 3)
    y1 = dropAry(ary1, 0)
    z1 = dropAry(ary1, 3, faReverse)
    w1 = dropAry(ary1, 9)
    printAry (x)
    printAry (y)
    printAry (z)
    printAry (w)
    printAry (x1)
    printAry (y1)
    printAry (z1)
    printAry (w1)
End Sub

Sub checkTake()
    ary = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim ary1(1 To 9)
    For i = 1 To 9
        ary1(i) = i
    Next
    y = takeAry(ary, 0)
    z = takeAry(ary, 3)
    w = takeAry(ary, 3, faReverse)
    y1 = takeAry(ary1, 0)
    z1 = takeAry(ary1, 3)
    w1 = takeAry(ary1, 3, faReverse)
    printAry (y)
    printAry (z)
    printAry (w)
    printAry (y1)
    printAry (z1)
    printAry (w1)
End Sub

Sub checkCon()
    a = mkSeq(10000)
    b = mkSeq(10000, 2, 2)
    c = mkSeq(10000, 3, 3)
    Call printTime("conarys", a, b, c)
    printAry (conArys(a, b, c))
End Sub

Sub checkMapA()
    a = mkSeq(10)
    b = mkSeq(10001, 0, 3)
    x = printTime("mapA", "calc", a, 1, "+")
    printAry (x)
    y = printTime("mapA", "calc", a, 2, "*")
    printAry (y)
    t1 = Time
    z0 = mapA("calc", b, 1, "+")
    t2 = Time
    printAry z0
    outPut "mapA: -" & Format(t2 - t1, "hh:nn:ss")
End Sub

Sub checkRgt()
    rg = Range("A1:A2")
    outPut TypeName(rg)
    outPut IsArray(rg)
End Sub

Sub checkRangeToArys()
    Dim rg As Range
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
    b = rangeToArys(rg, rowColumn.faColumn)
    Ad = rangeToArys(dary)
    bd = rangeToArys(dary, rowColumn.faColumn)
    Adr = rangeToArys(dr)
    bdr = rangeToArys(dr, rowColumn.faColumn)
    printAry (a)
    printAry (b)
    printAry (Ad)
    printAry (bd)
    printAry (Adr)
    printAry (bdr)
End Sub

Sub checkElm()
    Dim a(0 To 1, 0 To 2, 0 To 3, 0 To 4)
    Dim i As Long
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
    outPut x
    sp = getAryShape(a)
    lsp = getAryShape(a, faLower)
    n = reduceA("calc", sp, "*")
    For i = 0 To n - 1
        idx = mkIndex(i, sp, lsp)
        y = getElm(a, idx)
        outPut y & ","
    Next i
End Sub

Sub checkRangeToAry()
    Dim rg As Range
    Set rg = Range("A1:C2")
    dary = rg
    Dim dr(0 To 1, 0 To 2)
    dr(0, 0) = "a"
    dr(0, 1) = "b"
    dr(0, 2) = rowColumn.faColumn
    dr(1, 0) = 1
    dr(1, 1) = 2
    dr(1, 2) = 3
    a = rangeToAry(rg, rowColumn.faRow, 2)
    b = rangeToAry(rg, rowColumn.faColumn, 2)
    Ad = rangeToAry(dary, rowColumn.faRow, 2)
    bd = rangeToAry(dary, rowColumn.faColumn, 2)
    Adr = rangeToAry(dr, rowColumn.faRow, 2)
    bdr = rangeToAry(dr, rowColumn.faColumn, 2)
    printAry (a)
    printAry (b)
    printAry (Ad)
    printAry (bd)
    printAry (Adr)
    printAry (bdr)
End Sub

Sub checkAt()
    a = Array(1, 2, 3, 4, 5, 6)
    Dim b(1 To 6)
    For i = 1 To 6
        b(i) = i
    Next
    outPut getAryAt(a, 2)
    outPut getAryAt(b, 2)
    outPut getAryAt(a, -2)
    outPut getAryAt(b, -2)
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
    outPut getAryAt(a, 2, 0)
    outPut getAryAt(b, 2, 0)
    outPut getAryAt(a, -2, 0)
    outPut getAryAt(b, -2, 0)
    Call setAryAt(a, 2, -3, 0)
    Call setAryAt(b, 2, -3, 0)
    Call setAryAt(a, -2, -5, 0)
    Call setAryAt(b, -2, -5, 0)
    printAry (a)
    printAry (b)
End Sub

Sub checkadd()
    x = Array(Empty, Empty)
    outPut lenAry(x)
    printAry (x)
End Sub

Sub checkShape()
    Dim a(1 To 3, 1 To 4, 1 To 5)
    ' Dim a(1 To 3, 1 To 4)
    Dim b(0 To 3, 0 To 4, 0 To 5)
    Dim c(1 To 5)
    vl = 1
    fob = mkF(1, "calc", Empty, 1, "+")
    Call setAryMByF(a, fob)
    Call setAryMByF(b, fob)
    Call setAryMByF(c, fob)
    x = getAryShape(a)
    y = getAryShape(b)
    z = getAryShape(c)
    printAry (x)
    printAry (y)
    printAry (z)
    x = getAryShape(a, faUpper)
    y = getAryShape(b, faUpper)
    z = getAryShape(c, faUpper)
    printAry (x)
    printAry (y)
    printAry (z)
    x = getAryShape(a, faLower)
    y = getAryShape(b, faLower)
    z = getAryShape(c, faLower)
    printAry (x)
    printAry (y)
    printAry (z)
    Call printTime("printAry", a)
    Call printTime("printAry", b)
    Call printTime("printAry", c)
    Stop
End Sub

Sub checkApply()
    a = mkSeq(30)
    e = mapA("applyF", a, mkF(2, "calc", 2, Empty, "^"))
    printAry (e)
    fob0 = Array(Array(Array(1), Array("calc", Empty, 3, "+")), Array(Array(2), Array("calc", 100, Empty, "/")))
    fob1 = Array(mkF(1, "calc", Empty, 3, "+"), mkF(2, "calc", 100, Empty, "/"))
    b0 = mapA("applyFs", a, fob0)
    b1 = mapA("applyFs", a, fob1)
    printAry (b0)
    printAry (b1)
End Sub

Sub checkmkF()
    a = mkF(1, "calc", Empty, 3, "%")
    b = mkF(2, 1, "calc", Empty, Empty, "-")
    printAry (a)
    printAry (b)
End Sub

Sub checkPrmAry()
    a = Array(1, 2, 3, Array(4, 5, 6), Array(7, 8, 9))
    b = prmAry(a)
    printAry (a)
    printAry (b)
End Sub

Sub checkFoldF()
    fo = mkF(2, 1, "calc", Empty, Empty, "-")
    sq = mkSeq(5)
    a = foldF(fo, sq, 1)
    outPut a
End Sub

Sub checkZipApply()
    fob = mkF(1, 2, "calc", Empty, Empty, "+")
    z = zipApplyF(fob, mkSeq(5), mkSeq(5, 10, -2))
    printAry (z)
End Sub

Sub checkZip()
    x = zip(Array(1, 2, 3, 4), Array(2, 3, 4, 5), Array(3, 4, 5, 6))
    printAry (x)
    a = mkSeq(5)
    b = mkSeq(5, 10, -2)
    c = zip(a, b)
    printAry (c)
    d = Array(Array(1, 2), Array(3, 4), Array(5, 6))
    e = Array(7, 8, 9)
    f = Array("a", "b", "c")
    y = zip(d, e, f)
    printAry (y)
    Stop
End Sub

Sub checkAry()
    Dim x(1 To 3, 1 To 3) As String
    For i = 1 To 3
        For j = 1 To 3
            x(i, j) = chr(65 + (i - 1) + (j - 1) * 3)
        Next j
    Next i
    Set y = Range("a1:c2")
    z = Range("a1:c2")
    outPut TypeName(x)
    printAry (x)
    outPut TypeName(y)
    printAry (y)
    outPut TypeName(z)
    printAry (z)
End Sub

Sub checkZipArrayTime()
    a = mkSeq(10)
    b = mkSeq(10, 20, -2)
    c = mkSeq(10, 100, -10)
    ' a = mkSeq(100000)
    ' b = mkSeq(100000, 200000, -2)
    ' c = mkSeq(100000, 1000000, -10)
    x = Array(a, b, c)
    y = printTime("zipary", x)
    z = printTime("zip", a, b, c)
    Call printTime("conarys", x)
    printAry x
    printAry y
    printAry z
    'Stop
End Sub

Sub checkGetAryNum()
    Dim a(3, 4, 5)
    Dim b(1 To 3, 1 To 4, 1 To 5)
    x = getAryNum(a)
    y = getAryNum(b)
    outPut x
    outPut y
End Sub

Sub checkMAry()
    Dim a(3, 4)
    Dim b(1 To 3, 1 To 4)
    c = mkSeq(60, 1, 2)
    Call setAryMbyS(a, c)
    Call setAryMbyS(b, c)
    printAry (a)
    outPut
    printAry (b)
    Range("a1").Resize(4, 5) = a
    Range("a6").Resize(3, 4) = b
End Sub

Sub checkFlatten()
    Dim a(3, 4, 5)
    b = mkSeq(120)
    Call setAryMbyS(a, b)
    x = flattenAry(a)
    y = getArySbyM(a)
    printAry (a)
    printAry (b)
    printAry (x)
    printAry (y)
    outPut
    Dim f(2, 3)
    g = mkSeq(12, 11)
    Call setAryMbyS(f, g)
    d = Array(1, 2, Array(3, 4, Array(5, 6), 7, Array(8), f), 9, 1)
    w = flattenAry(d)
    printAry (w)
End Sub

Sub checkReshape()
    a = reshapeAry(mkSeq(720, 1, 2), Array(3, 4, 5, 6))
    b = reshapeAry(mkSeq(720, 1, 2), Array(3, 4, 5, 6), 1)
    'e = printTime("reshapeAry0", mkSeq(1000000), Array(100, 100, 100))
    c = printTime("reshapeAry", mkSeq(100000), Array(10, 100, 100))
    f = printTime("reshapeAry", mkSeq(100000), Array(10, 100, 100), 1)
    d = reshapeAry(mkSeq(27000), Array(30, 30, 30), 1)
    Stop
    printTime "printAry", a
    printTime "printAry", b
    printTime "printAry", c
    'Stop
    'printTime "printAry", e
    Stop
    printTime "printAry", d
End Sub

Sub checkSequence()
    t1 = Time
    r = 20
    c = 100
    y = reshapeAry(mkSeq(r * c), Array(r, c))
    t2 = Time
    outPut Format(t2 - t1, "hh:mm;ss")
    Call printTime("printAry", y)
    Stop
    Call printTime("print2DAry", y)
    Stop
    ' t3 = Time
    ' x = Application.WorksheetFunction.Sequence(500, 100)
    ' t4 = Time
    ' output Format(t4 - t3, "hh:mm;ss")
    ' Call printTime("printAry", x)
    ' Stop
End Sub

Sub checkMaryAccessor()
    x = reshapeAry(mkSeq(100), Array(2, 3, 4))
    printAry x
    y = getMAryAt(x, Array(1, 1, 1))
    outPut y
    z = getMAryAt(x, Array(1, 1, 1), 0)
    outPut z
    Call setMAryAt(x, Array(1, 1, 1), -1)
    Call setMAryAt(x, Array(1, 1, 1), -2, 0)
    printAry x
    ' x0 = Application.WorksheetFunction.Sequence(4, 5)
    ' printAry x0
    ' y0 = getMAryAt(x0, Array(1, 1))
    ' outPut y0
    ' z0 = getMAryAt(x0, Array(1, 1), 0)
    ' outPut z0
    ' Call setMAryAt(x0, Array(1, 1), -1)
    ' Call setMAryAt(x0, Array(1, 1), -2, 0)
    ' printAry x0
End Sub

Sub checkl_()
    Dim x As Variant
    Dim y As Variant
    Dim z As Variant
    x = Array(Array(Array(10, 11), Array(20, 21)), Array(Array(30, 31)), Array(Array(40, 41), Array(50, 51), Array(60, 61)))
    y = l_(l_(l_(10, 11), l_(20, 21)), l_(l_(30, 31)), l_(l_(40, 41), l_(50, 51), l_(60, 61)))
    z = l_()
    printAry (x)
    printAry (y)
    printAry (z)
    outPut TypeName(x)
    outPut TypeName(y)
    outPut TypeName(z)
End Sub

Sub checkmapMA()
    x1 = reshapeAry(mkSeq(24), Array(2, 3, 4))
    printAry (x1)
    outPut
    y1 = mapMA("calc", x1, 5, "*")
    printAry (y1)
    ReDim x2(1 To 4, 1 To 5)
    Call setAryMbyS(x2, mkSeq(20))
    printAry (x2)
    outPut
    y2 = mapMA("calc", x2, 5, "-")
    printAry (y2)
    fob = mkF(2, "calc", 3, Empty, "-")
    fobs = Array(fob, mkF(1, "calc", Empty, 3, "*"))
    y3 = mapMA("ApplyF", x2, fob)
    y4 = mapMA("ApplyFs", x2, fobs)
    printAry (y3)
    printAry (y4)
End Sub

Sub checkSpill()
    r = 20
    c = 100
    x = reshapeAry(mkSeq(r * c), Array(r, c))
    'x = Application.WorksheetFunction.Sequence(2000, 10000)
    'Call LogSetting.setAllFlg(True, True)
    printTime "print2DAry", x
End Sub

Sub checkSimpleAry()
    r = 500
    c = 100
    ' r = 1048576
    ' c = 16384
    x = reshapeAry(mkSeq(r * c), Array(r, c))
    printTime "print2DAry", x
    Stop
    printTime "printSimpleAry", x
    Stop
    printTime "printAry", x
    Stop
    ' Call LogSetting.setAllFileFlg(True)
    printTime "print2DAry", x
    Stop
    printTime "printSimpleAry", x
    Stop
    printTime "printAry", x
    Stop
    'Call LogSetting.setDic(False, True, "array")
    printTime "print2DAry", x
    Stop
    printTime "printSimpleAry", x
    Stop
    printTime "printAry", x
End Sub

Sub check3DArray()
    d = 10
    r = 10
    c = 10
    x = reshapeAry(mkSeq(d * r * c), Array(d, r, c))
    ' LogSetting.setAllFileFlg (True)
    ' Call LogSetting.setDic(False, True, "array")
    printTime "print3DAry", x
    printTime "print3DAry", x
    printTime "print3DAry", x
End Sub

Sub check1DArray()
    x = mkSeq(1000000)
    ' Call LogSetting.setAllFlg(True, True)
    printTime "print1DAry", x
End Sub

Sub checkPoly()
    outPut poly(-2, Array(1, 2, 3))
    outPut (polyStr(Array(2, -3, 4, 5)))
    outPut polyStr(Array(2, 3.2, 0, 5))
    outPut polyStr(Array(1, 3, 0, 0))
    outPut polyStr(Array(1, 1, 0, 1))
    outPut polyStr(Array(0, 0, 1, 0))
    outPut polyStr(Array(5))
    outPut polyStr(Array(1))
    outPut polyStr(Array(0))
    x = Array(1, 2, 1, -2, 0, 2, 0)
    outPut polyStr(x)
    outPut poly(3, x)
    y = revAry(x)
    outPut polyStr(y)
    outPut poly(3, y)
End Sub

Function mk2DSeq1(r, c, Optional first As Long = 1, Optional step As Long = 1, Optional bs As Long = 0)
    sp = Array(r, c)
    ret = mkMAry(sp, bs)
    Call set2DArySeq(ret, first, step)
    mk2DSeq1 = ret
End Function

Function mkSequence(r, n, Optional first = 1, Optional step = 1) As Variant()
    ret = Application.WorksheetFunction.Sequence(r, n, first, step)
    mkSequence = ret
End Function

Sub checkmkSeq()
    r = 1000
    c = 10000
    first = -100
    step = 7
    ' x1 = printTime("mkSequence", r, c, first, step)
    x2 = printTime("mk2DSeq", r, c, first, step)
    x3 = printTime("mk2DSeq1", r, c, first, step)
    x4 = printTime("mkMSeq", Array(r, c), first, step)
    ' t1 = Timer
    ' x5 = Application.WorksheetFunction.Sequence(r, c, first, step)
    ' t2 = Timer
    ' outPut ("worksheetfunction" & " - " & secToHMS(t2 - t1))
    Stop
End Sub

Sub checkCalcMary()
    x = mk2DSeq(4, 5, -10)
    y = mk2DSeq(4, 5, 5, -1, 1)
    z = calcMAry(x, y, "*")
    w = mapMA("fmt", z, "0000")
    printAry (w)
End Sub

Sub checkfmt()
    x = mkMSeq(Array(4, 5), 0.5, -1 / 3)
    Call setElm("a", x, Array(1, 1))
    Call setElm("2", x, Array(2, 2))
    printAry (x)
    outPut
    Call printAry(x, "")
    outPut
    Call printAry(x, , "0.000", Alignedirection.faRight, 7)
    outPut
    Call printAry(x)
End Sub

Sub checkMath()
    Pi = Atn(1) * 4
    x = mapA("calc", mkSeq(101, 0, 1), 2 * Pi / 100, "*")
    y0 = mapA("math_", x, "sin")
    z0 = mapA("math_", x, "cos")
    y1 = mapA("calc", y0, 2, "^")
    z1 = mapA("calc", z0, 2, "^")
    w = calcAry(y1, z1, "+")
    Call printAry(y0, , "0.000", , 6)
    Call printAry(z0, , "0.000", , 6)
    Call printAry(w, , "0.000", , 6)
End Sub

Sub checkWhile()
    x = mkSeq(10)
    y1 = takeWhile("comp_", x, faDirect, 6, "<")
    y2 = takeWhile("comp_", x, faDirect, 6, ">")
    y3 = takeWhile("comp_", x, faReverse, 6, "<=")
    y4 = takeWhile("comp_", x, faReverse, 6, ">=")
    y5 = dropWhile("comp_", x, faDirect, 6, "<")
    y6 = dropWhile("comp_", x, faDirect, 6, ">")
    y7 = dropWhile("comp_", x, faReverse, 6, "<=")
    y8 = dropWhile("comp_", x, faReverse, 6, ">=")
    printAry y1
    printAry y2
    printAry y3
    printAry y4
    printAry y5
    printAry y6
    printAry y7
    printAry y8
End Sub

Sub checkInfo()
    x = Sheets("check").Range("e3:j3")
    x0 = rangeToAry(x)
    printAry x0
    y = dropWhile("info_", x0, -1, "isEmpty")
    printAry y
    z = evalA(y)
    outPut z
End Sub

Sub checkDicStr()
    Set dic = CreateObject("Scripting.Dictionary")
    dic.Add "right", 1
    dic.Add "left", 2
    a = dic.keys
    b = dic.items
    c = zip(mapA("addStr", a, "'", "'"), b)
    d = mapA("mcjoin", c, ":")
    outPut TypeName(dic)
    str1 = mcJoin(d, ",", "Dic(", ")")
    Debug.Print str1
    printAry dic
End Sub

Sub checkClcStr()
    Set clc = New Collection
    clc.Add 1
    clc.Add 2
    clc.Add "abc"
    printAry clc
    Set clc1 = mkClc("a", "b", 3, clc)
    Call printAry(clc1)
End Sub

Sub checkDic()
    Set dic = mkDic("datetime", "date", "time", "date")
    outPut lookupDic("a", dic, "double")
    outPut lookupDic("time", dic, "double")
End Sub

Sub checkEqAry()
    n = 10
    x1 = mkMSeq(Array(n), , , 0)
    x2 = mkMSeq(Array(n), , , 0)
    x3 = mkMSeq(Array(n), , , 1)
    x4 = mkMSeq(Array(n), , 2, 0)
    x5 = mkMSeq(Array(n, n), , 2, 0)
    x6 = 1
    outPut (eqAry(x1, x2))
    outPut (eqAry(x1, x3))
    outPut (eqAry(x1, x4))
    outPut (eqAry(x1, x5))
    outPut (eqAry(x1, x6))
End Sub

Sub checkZipWithIndex()
    x = Array("a", "b", "c", "d", "e")
    y = zipWithIndex(x)
    z = zipWithIndex(x, 0)
    w = zipWithIndex(x, -1, 2)
    printAry y
    printAry z
    printAry w
End Sub
