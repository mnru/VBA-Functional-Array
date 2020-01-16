VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pn_
Private DicToFileFlg_
Private DicToConsoleFlg_

Private Sub Class_Initialize()
    createDic
    setDefault
End Sub

Sub setDefault()
    Call setDic("array", True, False)
    Call setDic("time", True, False)
    setLogPath
End Sub

Sub createDic()
    Set DicToFileFlg_ = CreateObject("Scripting.dictionary")
    Set DicToConsoleFlg_ = CreateObject("Scripting.dictionary")
End Sub

Sub setDic(lType As String, toConsole As Boolean, toFile As Boolean)
    DicToConsoleFlg_(LCase(lType)) = toConsole
    DicToFileFlg_(LCase(lType)) = toFile
End Sub

Property Get consoleFlg(lType As String) As Boolean
    consoleFlg = DicToConsoleFlg_(LCase(lType))
End Property

Property Get fileFlg(logtype As String) As Boolean
    fileFlg = DicToFileFlg_(LCase(logtype))
End Property

Sub setAllFileFlg(bol As Boolean)
    For Each key In DicToFileFlg_.keys
        DicToFileFlg_(key) = bol
    Next key
End Sub

Sub setAllConsoleFlg(bol As Boolean)
    For Each key In DicToConsoleFlg_.keys
        DicToConsoleFlg_(key) = bol
    Next key
End Sub

Sub setLogPath(Optional pn = "")
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

Sub writeLog(msg, Optional crlf As Boolean = True, Optional logtype As String)
    If Me.consoleFlg(logtype) Then
        Call writeToConsole(msg, crlf)
    End If
    If Me.fileFlg(logtype) Then
        Call writeToFile(msg, crlf)
    End If
End Sub
