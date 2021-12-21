Attribute VB_Name = "modUtil"
Option Base 0
Option Explicit

Function toString(elm, Optional qt As String = "'", Optional fm As String = "", _
    Optional lcr = "r", Optional width As Long = 0, _
    Optional insheet As Boolean = False) As String
    Dim ret As String, tmp As String
    Dim i As Long, aNum As Long
    Dim sp, lsp, idx, idx0, vl, dlm
    ret = ""
    If IsArray(elm) Then
        ret = ret & "["
        sp = getAryShape(elm)
        lsp = getAryShape(elm, "l")
        aNum = getAryNum(elm)
        If aNum = 0 Then
            ret = ret & "]"
        Else
            For i = 0 To aNum - 1
                idx0 = mkIndex(i, sp)
                idx = calcAry(idx0, lsp, "+")
                vl = getElm(elm, idx)
                dlm = getDlm(sp, idx0, insheet)
                ret = ret & toString(vl, qt, fm, lcr, width) & dlm
            Next i
        End If
    ElseIf IsObject(elm) Then
        If TypeName(elm) = "Dictionary" Then
            ret = ret & toStringDic(elm)
        ElseIf TypeName(elm) = "Collection" Then
            ret = ret & toStringClc(elm)
        Else
            ret = ret & "<" & TypeName(elm) & ">"
        End If
    ElseIf IsNull(elm) Then
        ret = ret & "Null"
    ElseIf IsEmpty(elm) Then
        ret = ret & "Empty"
    Else
        If TypeName(elm) = "String" Then
            tmp = qt & elm & qt
        Else
            tmp = fmt(elm, fm)
        End If
        tmp = align(tmp, lcr, width)
        ret = ret & tmp
    End If
    toString = ret
End Function

Function toStringDic(dic, Optional dlm = ":") As String
    Dim tmp1, tmp2
    Dim ret As String
    tmp1 = zip(mapA("tostring", dic.keys), mapA("tostring", dic.items))
    tmp2 = mapA("mcJoin", tmp1, dlm)
    ret = mcJoin(tmp2, ",", "Dic(", ")")
    toStringDic = ret
End Function

Function toStringClc(clc) As String
    Dim tmp1, tmp2
    Dim ret As String
    tmp1 = clcToAry(clc)
    tmp2 = mapA("toString", tmp1)
    ret = mcJoin(tmp2, ",", "Clc(", ")")
    toStringClc = ret
End Function

Function getDlm(shape, idx, Optional insheet As Boolean = False) As String
    Dim ret As String
    Dim i As Long, lNum As Long, m As Long
    Dim nl As String
    nl = IIf(insheet, vbLf, vbCrLf)
    lNum = lenAry(shape)
    m = 0
    For i = lNum To 1 Step -1
        If getAryAt(shape, i) - 1 > getAryAt(idx, i) Then
            m = i
            Exit For
        End If
    Next i
    Select Case m
        Case 0
            ret = "]"
        Case lNum
            ret = ","
        Case lNum - 1
            ret = ";" & nl & " "
        Case Else
            ret = String(lNum - m, ";") & nl & nl & " "
    End Select
    getDlm = ret
End Function

Function secToHMS(vl As Double) As String
    Dim ret As String
    Dim x1 As Long
    Dim x0 As Double, x2 As Double, x3 As Double
    Dim tmp
    x0 = vl
    x1 = Int(x0)
    x2 = x0 - x1
    tmp = mkIndex(x1, Array(24, 60, 60))
    x3 = getAryAt(tmp, 3) + x2
    ret = Format(getAryAt(tmp, 1), "00") & ":" & Format(getAryAt(tmp, 2), "00") & ":" & Format(x3, "00.000")
    secToHMS = ret
End Function

Function clcToAry(clc)
    Dim cnt As Long, i As Long
    Dim ret
    cnt = clc.Count
    If cnt = 0 Then
        ret = Array()
    Else
        ReDim ret(0 To cnt - 1)
        For i = 1 To cnt
            assign_ ret(i - 1), clc.Item(i)
        Next i
    End If
    clcToAry = ret
End Function

Function flattenAry(ary)
    Dim elm, el, ret
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
    Dim lb As Long, ub As Long, i As Long
    Dim ret
    lb = LBound(polyAry)
    ub = UBound(polyAry)
    ret = polyAry(lb)
    For i = lb + 1 To ub
        ret = ret * x + polyAry(i)
    Next
    poly = ret
End Function

Function polyStr(polyAry) As String
    Dim i As Long, lNum As Long
    Dim ret As String
    Dim c
    ret = ""
    lNum = lenAry(polyAry)
    For i = 1 To lNum
        c = getAryAt(polyAry, i)
        If c <> 0 Then
            If ret <> "" Then ret = ret & " "
            If c > 0 Then ret = ret & "+"
            If i = lNum Then
                ret = ret & c
            Else
                If c = 1 Then
                    ret = ret & "X"
                ElseIf c = -1 Then
                    ret = ret & "-X"
                Else
                    ret = ret & c & "X"
                End If
                If i < lNum - 1 Then ret = ret & "^" & lNum - i
            End If
        End If
    Next i
    If ret = "" Then ret = getAryAt(polyAry, -1)
    If Left(ret, 1) = "+" Then ret = Right(ret, Len(ret) - 1)
    polyStr = ret
End Function

Function fmt(expr, Optional fm As String = "", Optional lcr = "r", Optional width As Long = 0) As String
    Dim ret As String
    ret = Format(expr, fm)
    ret = align(ret, lcr, width)
    fmt = ret
End Function

Function align(str As String, Optional lcr = "r", Optional width As Long = 0) As String
    Dim ret As String
    Dim d As Long
    ret = CStr(str)
    d = width - Len(ret)
    If d > 0 Then
        Select Case LCase(lcr)
            Case "r": ret = space(d) & ret
            Case "l": ret = ret & space(d)
            Case "c": ret = space(d \ 2) & ret & space(d - d \ 2)
            Case Else:
        End Select
    End If
    align = ret
End Function

Function math_(x, symbol As String)
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
    math_ = ret
End Function

Function comp_(x, y, symbol As String)
    Dim ret
    Select Case symbol
        Case "=": ret = x = y 'caution Assign_ and eqaul is same symbol
        Case "<>": ret = x <> y
        Case "<": ret = x < y
        Case ">": ret = x > y
        Case "<=": ret = x <= y
        Case ">=": ret = x >= y
        Case "<": ret = x < y
        Case Else:
    End Select
    comp_ = ret
End Function

Function info_(x, symbol As String)
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
    info_ = ret
End Function

Function id_(x)
    If IsObject(x) Then
        Set id_ = x
    Else
        id_ = x
    End If
End Function

Function mkDic(ParamArray argAry())
    Dim ary, ret
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
    If dic.exists(x) Then
        ret = dic(x)
    Else
        ret = default
    End If
    lookupDic = ret
End Function

Function mkClc(ParamArray argAry())
    Dim ary, elm, clc
    ary = argAry
    Set clc = New Collection
    For Each elm In ary
        clc.Add elm
    Next elm
    Set mkClc = clc
End Function

Function l_(ParamArray argAry() As Variant)
    'works like function array()
    Dim ary As Variant
    ary = argAry
    l_ = ary
End Function
