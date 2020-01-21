Attribute VB_Name = "modLog"
Sub printOut(x, Optional crlf As Boolean = True)
     Set wr = New LogWriter
    wr.logType = "debug"
    Call wr.output(x, crlf)
End Sub

Sub printAry(ary)
    Set wr = New LogWriter
    wr.logType = "array"
    Call wr.output(toString(ary), True)
End Sub

Sub printSimpleAry(ary, Optional flush = 1000)
    Set wr = New LogWriter
    wr.logType = "array"
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aryNum = getAryNum(ary)
    If aryNum = 0 Then
        Call wr.output("[]", False)
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
                Call wr.output(ret, False)
                ret = ""
            End If
        Next i
    End If
    Call wr.output(ret, True)

End Sub

Sub print1DAry(ary, Optional flush = 1000)
    Set wr = New LogWriter
    wr.logType = "array"
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
            Call wr.output(ret, False)
            ret = ""
        End If
        cnt = cnt + 1
    Next i1
    Call wr.output(ret, True)
End Sub

Sub print2DAry(ary, Optional flush = 1000)
    Set wr = New LogWriter
    wr.logType = "array"
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
                Call wr.output(ret, False)
                ret = ""
            End If
            cnt = cnt + 1
        Next i2
    Next i1
    Call wr.output(ret, True)
End Sub

Sub print3DAry(ary, Optional flush = 1000)
    Set wr = New LogWriter
    wr.logType = "array"
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    lb3 = LBound(ary, 3): ub3 = UBound(ary, 3)

    cnt = 1
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            For i3 = lb3 To ub3

                elm = CStr(ary(i1, i2, i3))
                If i3 < ub3 Then
                    dlm = ","
                ElseIf i2 < ub2 Then
                    dlm = ";" & vbCrLf
                ElseIf i1 < ub1 Then
                    dlm = ";;" & vbCrLf & vbCrLf

                Else
                    dlm = "]"
                End If
                ret = ret & elm & dlm
                If cnt Mod flush = 0 Then
                    Call wr.output(ret, False)
                    ret = ""
                End If
                cnt = cnt + 1
            Next i3
        Next i2
    Next i1
    Call wr.output(ret, True)
End Sub

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Set wr = New LogWriter
    wr.logType = "time"
    Dim etime   As Double
    Dim stime   As Double
    Dim secs    As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call wr.output(fnc & " - " & secToHMS(secs), True)
End Function
