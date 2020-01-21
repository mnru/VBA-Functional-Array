VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pn_ As String
Private DicToFileFlg_
Private DicToConsoleFlg_

Property Get consoleFlg(lType As String) As Boolean
    consoleFlg = DicToConsoleFlg_(LCase(lType))
End Property

Property Get fileFlg(logType As String) As Boolean
    fileFlg = DicToFileFlg_(LCase(logType))
End Property

Property Let logPath(pn As String)
    pn_ = pn
End Property

Property Get logPath() As String
    logPath = pn_
End Property

Sub setDefaultLogPath()
    pn_ = ThisWorkbook.Path & "\log.txt"
End Sub

Private Sub Class_Initialize()
    createDic
    setDefault
End Sub

Private Sub createDic()
    
    Set DicToFileFlg_ = CreateObject("Scripting.dictionary")
    Set DicToConsoleFlg_ = CreateObject("Scripting.dictionary")
    
    Call setDic(True, False, "array", "time", "debug")
    
End Sub


Sub setDefault()
    Call setAllFlg(True, False)
    setDefaultLogPath
End Sub

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

Sub setAllFlg(toConsole As Boolean, toFile As Boolean)
    setAllConsoleFlg (toConsole)
    setAllFileFlg (toFile)
End Sub

Sub setDic(toConsole As Boolean, toFile As Boolean, ParamArray logTypes())
    
    For Each lType In logTypes
        DicToConsoleFlg_(LCase(lType)) = toConsole
        DicToFileFlg_(LCase(lType)) = toFile
    Next
    
End Sub



