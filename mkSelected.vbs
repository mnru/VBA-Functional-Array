Dim fso, sstm, tstm
call main
Sub main()
  Dim sourcePath
  Set fso = CreateObject("Scripting.FileSystemObject")
  parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  targetFile = "modSelected_.bas"
  ary0 = Array("modSelected.bas", 0, 18)
  ary1 = Array(Array("modAry.bas", 5, -1), Array("modFnc.bas", 10, -1), Array("modUtil.bas", 11, -1), Array("modLog.bas", 5, 29))
  targetPath = parentPath & "\" & targetFile
  Set tstm = fso.createtextfile(targetPath)
  tstm.Close
  sourcePath = parentPath & "\" & ary0(0)
  Call cpFile(targetPath, sourcePath, ary0(1), ary0(2), False)
  For Each elm In ary1
    sourcePath = parentPath & "\src\" & elm(0)
    If elm(0) = "modAry.bas" Then
      Call cpFile1(targetPath, sourcePath, elm(1), elm(2), True)
    Else
      Call cpFile(targetPath, sourcePath, elm(1), elm(2), True)
    End If
  Next
End Sub

Sub cpFile(tf, sf, fromLine, toLine, bolFrom)
  Dim fso, tstm, sstm
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set tstm = fso.opentextfile(tf, 8) 'for appending
  Set sstm = fso.opentextfile(sf, 1) 'for reading
  If bolFrom Then
    tstm.writeblanklines (1)
    tstm.writeline (String(20, "'"))
    tstm.writeline ("from " & fso.getbasename(sf))
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
    tstm.writeline ("from " & fso.getbasename(sf))
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
  str1 = Trim(str)
  If Left(str1, 4) = "Case" Then
    tmp = InStr(str1, ":")
    If tmp = 0 Then
      ret = -1
    Else
      str2 = Mid(str1, 6, tmp - 6)
      If IsNumeric(str2) Then
        ret = CLng(str2)
      Else
        ret = -1
      End If
    End If
  Else
    ret = -1
  End If
  getCaseNum = ret
End Function