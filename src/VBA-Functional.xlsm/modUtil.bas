Attribute VB_Name = "modUtil"

Sub printAry(ary)
    Debug.Print toString(ary)
End Sub

Function toString(elm) As String
    Dim ret
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
                ret = ret & toString(vl) & dlm
            Next i
        End If
    Else
        If IsObject(elm) Then
            ret = ret & "<" & TypeName(elm) & ">"
        ElseIf TypeName(elm) = "String" Then
            ret = ret & "'" & CStr(elm) & "'"
        ElseIf IsNull(elm) Then
            ret = ret & "Null"
        Else
            ret = ret & CStr(elm)
        End If
    End If
    toString = ret
End Function

Function toString2DAry(ary)
    ret = "["
    lb1 = LBound(ary): ub1 = UBound(ary)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            
            elm = CStr(ary(i1, i2))
            If i2 < ub2 Then
                dlm = ","
            ElseIf i1 = ub1 Then
                dlm = "]"
            Else
                dlm = ";" & vbCrLf
            End If
            
            ret = ret & elm & dlm
            
        Next i2
    Next i1
    
    toString2DAry = ret
    
End Function

Function getDlm(shape, idx)
    Dim ret
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
            ret = ";" & vbCrLf
        Case Else
            ret = String(n - m, ";") & vbCrLf & vbCrLf
    End Select
    getDlm = ret
End Function

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim etime As Double
    Dim stime As Double
    Dim secs  As Double
    ary = argAry
    fnAry = conArys(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Debug.Print fnc & " - " & secToHMS(secs)
End Function

Function secToHMS(vl As Double)
    'Dim x2 As Double
    x0 = vl
    x1 = Int(x0)
    x2 = x0 - x1
    x3 = mkIndex(x1, Array(24, 60, 60))
    x4 = getAryAt(x3, 3) + x2
    ret = Format(getAryAt(x3, 1), "00") & ":" & Format(getAryAt(x3, 2), "00") & ":" & Format(x4, "00.000")
    secToHMS = ret
End Function


Function clcToAry(clc As Collection)
    cnt = clc.Count
    ReDim ret(1 To cnt)
    For i = 1 To cnt
        ret(i) = clc.Item(i)
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
