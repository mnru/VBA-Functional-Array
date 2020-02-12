Attribute VB_Name = "modUtil"
Option Base 0

Enum AlignDirection
    faLeft = 1
    faCenter = 0
    faRight = -1
End Enum

Function toString(elm, Optional qt = True, Optional fm As String = "", Optional lcr As AlignDirection = AlignDirection.faRight, Optional width As Long = 0, _
    Optional insheet As Boolean = False) As String
    Dim ret As String, tmp As String
    Dim i As Long, aryNum As Long
    Dim sp, lsp
    ret = ""
    If IsArray(elm) Then
        d = dimAry(elm)
        ret = ret & "["
        sp = getAryShape(elm)
        lsp = getAryShape(elm, "L")
        aryNum = getAryNum(elm)
        If aryNum = 0 Then
            ret = ret & "]"
        Else
            For i = 0 To aryNum - 1
                idx0 = mkIndex(i, sp)
                idx = calcAry(idx0, lsp, "+")
                vl = getElm(elm, idx)
                dlm = getDlm(sp, idx0)
                ret = ret & toString(vl, qt, fm, lcr, width) & dlm
            Next i
        End If
    Else
        If TypeName(elm) = "Dictionary" Then
            ret = ret & dicToStr(elm)
        ElseIf TypeName(elm) = "Collection" Then
            ret = ret & clcToStr(elm)
        ElseIf IsObject(elm) Then
            ret = ret & "<" & TypeName(elm) & ">"
        ElseIf IsNull(elm) Then
            ret = ret & "Null"
        ElseIf IsEmpty(elm) Then
            ret = ret & "Empty"
        Else
            If TypeName(elm) = "String" Then
                If qt Then
                    tmp = "'" & elm & "'"
                    'tmp = """" & elm & """"
                Else
                    tmp = elm
                End If
            Else
                tmp = fmt(elm, fm)
            End If
            tmp = align(tmp, lcr, width)
            ret = ret & tmp
        End If
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
    ret = mcJoin(tmp2, ",", "Collection(", ")")
    clcToStr = ret
End Function

Function getDlm(shape, idx, Optional insheet As Boolean = False) As String
    Dim ret As String
    Dim i As Long
    Dim nl As String
    nl = IIf(insheet, vbLf, vbCrLf)
    n = lenAry(shape)
    m = 0
    For i = n To 1 Step -1
        If getAryAt(shape, i) - 1 > getAryAt(idx, i) Then
            m = i
            Exit For
        End If
    Next i
    Select Case m
        Case 0
            ret = "]"
        Case n
            ret = ","
        Case n - 1
            ret = ";" & nl & " "
        Case Else
            ret = String(n - m, ";") & nl & nl & " "
    End Select
    getDlm = ret
End Function

Function secToHMS(vl As Double)
    Dim x1 As Long
    Dim x2 As Double
    x0 = vl
    x1 = Int(x0)
    x2 = x0 - x1
    x3 = mkIndex(x1, Array(24, 60, 60))
    x4 = getAryAt(x3, 3) + x2
    ret = Format(getAryAt(x3, 1), "00") & ":" & Format(getAryAt(x3, 2), "00") & ":" & Format(x4, "00.000")
    secToHMS = ret
End Function

Function clcToAry(clc)
    cnt = clc.Count
    ReDim ret(1 To cnt)
    For i = 1 To cnt
        ret(i) = clc.item(i)
    Next i
    clcToAry = ret
End Function

Function flattenAry(ary)
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
    lb = LBound(polyAry)
    ub = UBound(polyAry)
    ret = polyAry(lb)
    For i = lb + 1 To ub
        ret = ret * x + polyAry(i)
    Next
    poly = ret
End Function

Function polyStr(polyAry) As String
    Dim i As Long
    Dim ret As String
    ret = ""
    n = lenAry(polyAry)
    For i = 1 To n
        c = getAryAt(polyAry, i)
        If c <> 0 Then
            If ret <> "" Then ret = ret & " "
            If c > 0 Then ret = ret & "+"
            If c <> 1 Or i = n Then ret = ret & c
            If i < n Then ret = ret & "X"
            If i < n - 1 Then ret = ret & "^" & n - i
        End If
    Next i
    If ret = "" Then ret = getAryAt(polyAry, -1)
    If Left(ret, 1) = "+" Then ret = Right(ret, Len(ret) - 1)
    polyStr = ret
End Function

Function fmt(expr, Optional fm As String = "", Optional lcr As AlignDirection = AlignDirection.faRight, Optional width As Long = 0) As String
    Dim ret As String
    ret = Format(expr, fm)
    ret = align(ret, lcr, width)
    fmt = ret
End Function

Function align(str As String, Optional lcr As AlignDirection = AlignDirection.faRight, Optional width As Long = 0) As String
    Dim ret As String
    ret = CStr(str)
    d = width - Len(ret)
    If d > 0 Then
        Select Case LCase(lcr)
            Case AlignDirection.faRight: ret = space(d) & ret
            Case AlignDirection.faLeft: ret = ret & space(d)
            Case AlignDirection.faCenter: ret = space(d \ 2) & ret & space(d - d \ 2)
            Case Else:
        End Select
    End If
    align = ret
End Function

Function math(x, symbol As String)
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
    math = ret
End Function

Function comp(x, y, symbol As String)
    Dim ret
    Select Case symbol
        Case "=": ret = x = y 'caution assign and eqaul is same symbol
        Case "<>": ret = x <> y
        Case "<": ret = x < y
        Case ">": ret = x > y
        Case "<=": ret = x <= y
        Case ">=": ret = x >= y
        Case "<": ret = x < y
        Case Else:
    End Select
    comp = ret
End Function

Function info(x, symbol As String)
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
    info = ret
End Function

Function id_(x)
    If IsObject(x) Then
        Set id_ = x
    Else
        id_ = x
    End If
End Function

Function mkDic(ParamArray argAry())
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
    If dic.Exists(x) Then
        ret = dic(x)
    Else
        ret = default
    End If
    lookupDic = ret
End Function

Function mkClc(ParamArray argAry())
    Dim ary
    ary = argAry
    Dim clc
    Set clc = New Collection
    For Each elm In ary
        clc.Add elm
    Next elm
    Set mkClc = clc
End Function
