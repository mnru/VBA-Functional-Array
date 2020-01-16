VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pn_
Private toFile_    As Boolean
Private toConsole_ As BooleanPrivate

Sub Class_Initialize()
    Call setLog(True, False)
End Sub

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
        stm.write (msg)
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
End Sub
