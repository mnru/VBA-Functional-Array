Dim fso, sstm, tstm
Call main

Sub main()
  Dim sourcePath
  Set fso = CreateObject("Scripting.FileSystemObject")
  parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  'parentPath = ThisWorkbook.Path
  sAry0 = Array("modSelected.bas", 2, 18)
  sAry1 = Array(Array("modAry.bas", 5, -1), Array("modFnc.bas", 10, -1), Array("modUtil.bas", 11, -1), Array("modLog.bas", 5, 29))
  tAry = Array("modSelectedNoLog.bas", "modSelected_.bas")
  For Each targetFile In tAry
    targetPath = parentPath & "\" & targetFile
    Set tstm = fso.createtextfile(targetPath)
    str0 = "Attribute VB_Name = """ & Replace(fso.getbasename(targetPath), "_", "") & """"
    tstm.writeline (str0)
    tstm.Close
    sourcePath = parentPath & "\" & sAry0(0)
    Call cpFile(targetPath, sourcePath, sAry0(1), sAry0(2), False)
    For Each sElm In sAry1
      sourcePath = parentPath & "\src\" & sElm(0)
      If sElm(0) = "modAry.bas" Then
        Call cpFile1(targetPath, sourcePath, sElm(1), sElm(2), True)
      ElseIf Not (sElm(0) = "modLog.bas" And targetFile = "modSelectedNoLog.bas") Then
        Call cpFile(targetPath, sourcePath, sElm(1), sElm(2), True)
      End If
    Next
  Next
  MsgBox "finished"
End Sub

Sub cpFile(tf, sf, fromLine, toLine, bolFrom)
  Dim fso, tstm, sstm
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set tstm = fso.opentextfile(tf, 8) 'for appending
  Set sstm = fso.opentextfile(sf, 1) 'for reading
  If bolFrom Then
    tstm.writeblanklines (1)
    tstm.writeline (String(20, "'"))
    tstm.writeline ("'from " & fso.getbasename(sf))
    tstm.writeline (String(20, "'"))
  End If
  Do While Not sstm.atEndOfStream
    Line = sstm.Line
    If Line < fromLine Or (toLine > 0 And Line > toLine) Then
      sstm.skipline
    Else
      str0 = sstm.readline
      tstm.writeline (str0)
    End If
  Loop
  sstm.Close
  tstm.Close
End Sub

Sub cpFile1(tf, sf, fromLine, toLine, bolFrom)
  Dim fso, tstm, sstm
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set tstm = fso.opentextfile(tf, 8) 'for appending
  Set sstm = fso.opentextfile(sf, 1) 'for reading
  If bolFrom Then
    tstm.writeblanklines (1)
    tstm.writeline (String(20, "'"))
    tstm.writeline ("'from " & fso.getbasename(sf))
    tstm.writeline (String(20, "'"))
  End If
  Do While Not sstm.atEndOfStream
    Line = sstm.Line
    If Line < fromLine Or (toLine > 0 And Line > toLine) Then
      sstm.skipline
    Else
      str0 = sstm.readline
      If getCaseNum(str0) <= 5 Then
        tstm.writeline (str0)
      End If
    End If
  Loop
  sstm.Close
  tstm.Close
End Sub

Function getCaseNum(str)
  ret = -1
  str1 = Trim(str)
  If Left(str1, 4) = "Case" Then
    tmp = InStr(str1, ":")
    If tmp = 0 Then
      str2 = Trim(Right(str1, Len(str1) - 5))
    Else
      str2 = Trim(Mid(str1, 6, tmp - 6))
    End If
    If IsNumeric(str2) Then
      ret = CLng(str2)
    End If
  End If
  getCaseNum = ret
End Function