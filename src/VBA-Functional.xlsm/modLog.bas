Attribute VB_Name = "modLog"
Sub printAry(ary)
    Dim logtype As String
    logtype = "array"
    Call DebugLog.writeLog(toString(ary), True, logtype)
End Sub

Sub printSimpleAry(ary, Optional flush = 1000)
    Dim logtype As String
    logtype = "array"
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aryNum = getAryNum(ary)
    If aryNum = 0 Then
        Call DebugLog.writeLog("[]", False, logtype)
    Else
        ret = "["
        For i = 0 To aryNum - 1
            idx0 = mkIndex(i, sp)
            idx = calcAry(idx0, lsp, "+")
            vl = getElm(ary, idx)
            If TypeName(vl) = "String" Then vl = "'" & vl & "'"
            dlm = getDlm(sp, idx0)
            ret = ret & vl & dlm
            If i Mod flush = 0 Then
                Call DebugLog.writeLog(ret, False, logtype)
                ret = ""
            End If
        Next i
    End If
    Call DebugLog.writeLog(ret, True, logtype)
    
End Sub

Sub print1DAry(ary, Optional flush = 1000)
    Dim logtype As String
    logtype = "array"
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    cnt = 1
    For i1 = lb1 To ub1
        elm = CStr(ary(i1))
        If i1 < ub1 Then
            dlm = ","
        Else
            dlm = "]"
        End If
        ret = ret & elm & dlm
        If cnt Mod flush = 0 Then
            Call DebugLog.writeLog(ret, False, logtype)
            ret = ""
        End If
        cnt = cnt + 1
    Next i1
    Call DebugLog.writeLog(ret, True, logtype)
End Sub
Sub print2DAry(ary, Optional flush = 1000)
    Dim logtype As String
    logtype = "array"
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    cnt = 1
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            elm = CStr(ary(i1, i2))
            If i2 < ub2 Then
                dlm = ","
            ElseIf i1 < ub1 Then
                dlm = ";" & vbCrLf
            Else
                dlm = "]"
            End If
            ret = ret & elm & dlm
            If cnt Mod flush = 0 Then
                Call DebugLog.writeLog(ret, False, logtype)
                ret = ""
            End If
            cnt = cnt + 1
        Next i2
    Next i1
    Call DebugLog.writeLog(ret, True, logtype)
End Sub
Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim logtype As String
    logtype = "time"
    Dim etime   As Double
    Dim stime   As Double
    Dim secs    As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call DebugLog.writeLog(fnc & " - " & secToHMS(secs), True, logtype)
End Function

