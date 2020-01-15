Attribute VB_Name = "modLog"
Private toFile_    As Boolean
Private toConsole_ As Boolean
Private pn_
Private stmLog     As Object

Sub setLog(Optional toConsole As Boolean = True, Optional toFile As Boolean = False, Optional pn = "")
    toConsole_ = toConsole
    toFile_ = toFile
    If pn = "" Then pn = ThisWorkbook.Path & "\log.txt"
    pn_ = pn
    
End Sub

Sub prepareLogFile()
    On Error GoTo Catch
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pn_) Then
        Set stmLog = fso.OpenTextFile(pn_, IOMode:=8) 'ForAppending
    Else
        Set stmLog = fso.CreateTextFile(pn_)
    End If
    Exit Sub
Catch:
    MsgBox Err.Description
    Err.Clear
    
End Sub

Sub closeLogFile()
    
    On Error Resume Next
    stmLog.Close
    On Error GoTo 0
    
End Sub

Sub writeToFile(msg, Optional crlf As Boolean = True)
    
    Call prepareLogFile
    If crlf Then
        stmLog.writeline (msg)
    Else
        stmLog.Write (msg)
    End If
    Call closeLogFile
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

