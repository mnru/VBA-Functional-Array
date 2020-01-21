VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pn_           As String
Private toConsoleFlg_ As Boolean
Private toFileFlg_    As Boolean
Private logType_      As String

Property Let logType(lt As String)
    logType_ = LCase(lt)
    updateSetting
End Property

Sub updateSetting()
    
    toConsoleFlg_ = LogSetting.consoleFlg(logType_)
    toFileFlg_ = LogSetting.fileFlg(logType_)
    pn_ = LogSetting.logPath
    
End Sub

Private Function prepareFile(pn As String)
    On Error GoTo Catch
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pn) Then
        Set stm = fso.OpenTextFile(pn, IOMode:=8) 'ForAppending
    Else
        Set stm = fso.CreateTextFile(pn)
    End If
    Set prepareFile = stm
    
    Exit Function
Catch:
    MsgBox Err.Description
    Err.Clear
    
End Function

Private Sub toFile(msg, Optional crlf As Boolean = True)
    
    Set stm = prepareFile(pn_)
    If crlf Then
        stm.writeline (msg)
    Else
        stm.write (msg)
    End If
    On Error Resume Next
    stm.Close
    On Error GoTo 0
End Sub

Private Sub toConsole(msg, Optional crlf As Boolean = True)
    If crlf Then
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
    
End Sub

Sub output(msg, Optional crlf As Boolean = True)
    If toConsoleFlg_ Then
        Call toConsole(msg, crlf)
    End If
    If toFileFlg_ Then
        Call toFile(msg, crlf)
    End If
End Sub



