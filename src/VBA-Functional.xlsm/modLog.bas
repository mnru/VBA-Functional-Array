Attribute VB_Name = "modLog"
Private pn_
Private toFile_    As Boolean
Private toConsole_ As Boolean

Sub setLog(Optional toConsole As Boolean = True, Optional toFile As Boolean = False, Optional pn = "")
    toConsole_ = toConsole
    toFile_ = toFile
    If pn = "" Then pn = ThisWorkbook.Path & "\log.txt"
    pn_ = pn
    
End Sub

Function prepareLogFile(pn)
    On Error GoTo Catch
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pn) Then
        Set stm = fso.OpenTextFile(pn, IOMode:=8) 'ForAppending
    Else
        Set stm = fso.CreateTextFile(pn)
    End If
    Set prepareLogFile = stm
    
    Exit Function
Catch:
    MsgBox Err.Description
    Err.Clear
    
End Function

Sub writeToFile(msg, Optional crlf As Boolean = True)
    
    Set stm = prepareLogFile(pn_)
    If crlf Then
        stm.writeline (msg)
    Else
        stm.Write (msg)
    End If
    On Error Resume Next
    stm.Close
    On Error GoTo 0
End Sub

Sub writeToConsole(msg, Optional crlf As Boolean = True)
    
    If crlf Then
        
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
    
End Sub

Sub writeLog(msg, Optional crlf As Boolean = True)
    If toConsole_ Then Call writeToConsole(msg, crlf)
    If toFile_ Then Call writeToFile(msg, crlf)
    ' If toSheet_ Then Call writeToSheet(msg, crlf)
    
End Sub

Sub printAry(ary)
    writeLog toString(ary)
End Sub

Sub printSimpleAry(ary, Optional flush = 1000)
    
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aryNum = getAryNum(ary)
    If aryNum = 0 Then
        Call writeLog("[]", False)
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
                Call writeLog(ret, False)
                ret = ""
            End If
        Next i
    End If
    Call writeLog(ret, True)
    
End Sub

Sub print1DAry(ary, Optional flush = 1000)
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
            Call writeLog(ret, False)
            ret = ""
        End If
        cnt = cnt + 1
    Next i1
    Call writeLog(ret, True)
End Sub

Sub print2DAry(ary, Optional flush = 1000)
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
                Call writeLog(ret, False)
                ret = ""
            End If
            cnt = cnt + 1
        Next i2
    Next i1
    Call writeLog(ret, True)
End Sub

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim etime As Double
    Dim stime As Double
    Dim secs  As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call writeLog(fnc & " - " & secToHMS(secs))
End Function
