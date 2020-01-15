Attribute VB_Name = "modLog"
Private stmLog
Private toFile_    As Boolean
Private toConsole_ As Boolean
Private toSheet_   As Boolean

Sub setLog(Optional toConsole As Boolean = True, Optional toFile As Boolean = False, Optional pn = "", Optional toSheet As Boolean = False)
    
    toConsole_ = toConsole
    toFile_ = toFile
    toSheet_ = toSheet
    If toFile Then
        Call prepareLogFile(pn)
    End If
    
    
End Sub

Sub prepareLogFile(Optional pn = "")
    
    
    If pn = "" Then pn = ThisWorkbook.Path & "\log.txt"
    
    On Error GoTo Catch
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(pn) Then
        
        Set stmLog = fso.OpenTextFile(pn, IOMode:=ForAppending)
        
    Else
        
        Set stmLog = fso.CreateTextFile(pn)
        
    End If
    
Catch:
    Exit Sub
    MsgBox Err.Description
    Err.Clear
    
    
    
End Sub

Sub closeLogFile()
    
    On Error Resume Next
    stmLog.Close
    On Error GoTo 0
    
    
End Sub

Sub writeToFile(msg, withNewLine)
    If withNewLine Then
        stmLog.writeline (msg)
    Else
        stmLog.Write (msg)
    End If
    
End Sub

Sub writeToConsole(msg, withNewLine)
    
    If withNewLine Then
        
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
    
End Sub

Sub writeLog(msg, Optional withNewLine As Boolean = True)
    If toConsole_ Then Call writeToConsole(msg, withNewLine)
    If toFile_ Then Call writeToFile(msg, withNewLine)
    ' If toSheet_ Then Call writeToSheet(msg, withNewLine)
    
End Sub

Sub printAry(ary)
    writeLog toString(ary)
End Sub

Sub printSimpleAry(ary)
    
    Call writeLog("[", False)
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aryNum = getAryNum(ary)
    If aryNum = 0 Then
        Call writeLog("]", False)
    Else
        
        For i = 0 To aryNum - 1
            idx0 = mkIndex(i, sp)
            idx = calcAry(idx0, lsp, "+")
            vl = getElm(ary, idx)
            dlm = getDlm(sp, idx0)
            Call writeLog(vl & dlm, False)
            
        Next i
    End If
    Call writeLog(vbCrLf, True)
    
End Sub
