Const fixedMode = True
Const toCRLF = True
Const clean = False
Dim targetExt
targetExt = "xlsm"

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_Document = 100

Const xlExcel9795 = 43           ' //.xls 97-2003 format in Excel 2003 or prev
Const xlExcel8 = 56             ' //.xls 97-2003 format in Excel 2007
Const xlTemplate = 17            ' //.xlt
Const xlAddIn = 18             ' //.xla
Const xlExcel12 = 50            ' //.xlsb
Const xlOpenXMLWorkbookMacroEnabled = 52  ' //.xlsm
Const xlOpenXMLTemplateMacroEnabled = 53  ' //.xltm
Const xlOpenXMLAddIn = 55          ' //.xlam

function saveType(ext)

select case lcase(ext)
case "xls":ret=xlExcel9795
case "xlt":ret=xlTemplate
case "xla":ret=xlAddIn
case "xlsb":ret=xlExcel12
case "xlsm":ret=xlOpenXMLWorkbookMacroEnabled
case "xltm":ret=xlOpenXMLTemplateMacroEnabled
case "xlam":ret=xlOpenXMLAddIn
case else:
end select

saveType=ret
end function

Call composeAll

Sub composeAll()
  'import excel macro module
 
  On Error Resume Next
  Dim oApp
  Dim oFso
  Dim sArModule()
  Dim sModule
  Dim sExt
  Dim sourcePath
  Dim parentPath
  Dim targetName
  Dim targetPath
  Dim targetBook
 
  Set oApp = CreateObject("Excel.Application")
'  oApp.DisplayAlerts = False
'  oApp.EnableEvents = False
 
  Set oFso = CreateObject("Scripting.FileSystemObject")
 
  If fixedMode Then
    tmp = getFixedPath(targetExt)
   
    parentPath = tmp(0)
    sourcePath = tmp(1)
    targetPath = tmp(2)
  Else
    sourcePath = getFolderPath
    If sourcePath = "" Then
      MsgBox "exit this script"
      Exit Sub
    End If
    parentPath = oFso.getParentFolderName(sourcePath)
    targetName = oFso.getFilename(parentPath) & "." & targetExt
    targetPath = oFso.buildPath(parentPath, targetName)
  End If
  If not oFso.FolderExists(sourcePath) Then
	msgbox "there is no sorce folder"
	exit sub
	end if
  If oFso.FileExists(targetPath) Then
      If clean Then Call cleanAll(targetPath)
  Else
    Set targetBook = oApp.Workbooks.Add
    Call targetBook.SaveAs(targetPath, saveType(targetExt))
    targetBook.Close
  End If
  Call addAll(sourcePath, targetPath)
  MsgBox "complete!"
End Sub

Function getFolderPath()
  'folder picker dialog
  Dim ret
  Dim oShl
  Dim oBrw
  Dim strPath
  On Error Resume Next
  Set oShl = WScript.CreateObject("Shell.Application")
  Set oBrw = oShl.BrowseForFolder(0, "Select sorce folder", &H10)
  If (oBrw Is Nothing) Then
    Err.Clear
    ret = ""
  Else
    ret = oBrw.Items.Item.Path
  End If
  Set oShl = Nothing
  Set oBrw = Nothing
  Err.Clear
  On Error GoTo 0
  'msgbox "folderPath=" & ret
  getFolderPath = ret
End Function

Sub lfToCrlf(pn)
  'change LF to CRLF in the file pn
  Dim oFso
  Dim oStm
  Set oFso = CreateObject("Scripting.FileSystemObject")
  Set oStm = oFso.openTextfile(pn)
  str0 = oStm.readAll
  oStm.Close
  txts = Split(str0, Chr(10))
    lb = LBound(txts)
    ub = UBound(txts)
    For i = ub To lb Step -1
        txt = txts(i)
        If txt = "" Or txt = Chr(13) Then
            ub = ub - 1
        Else
            Exit For
        End If
    Next

  Set oStm = oFso.createtextfile(pn)
  For i = lb To ub
 txt = txts(i)
    If Right(txt, 1) = Chr(13) Then
      txt = Left(txt, Len(txt) - 1)
    End If
    oStm.writeline (txt)
  Next
  oStm.Close
End Sub


Sub cleanAll(targetPath)
  Dim oApp
  Set oApp = CreateObject("Excel.Application")
  oApp.DisplayAlerts = False
  oApp.EnableEvents = False
 
  Set targetBook = oApp.Workbooks.Open(targetPath)
 
  On Error Resume Next
  Set cmps = targetBook.VBProject.VBComponents
  For Each cmp In cmps
    cn = cmp.Name
    If cmp.Type = vbext_ct_Document Then
     
      Call cmp.CodeModule.DeleteLines(1, cmp.CodeModule.CountOfLines)
    Else
      cmps.Remove (cmp)
    End If
  Next
  targetBook.Save
  targetBook.Close
  targetBook = Nothing
  oApp.Quit
  On Error GoTo 0
End Sub

Sub addAll(sourcePath, targetPath)
  Dim oApp
  Set oApp = CreateObject("Excel.Application")
  Set oFso = CreateObject("Scripting.FileSystemObject")
  oApp.DisplayAlerts = False
  oApp.EnableEvents = False

  On Error Resume Next
  Set targetBook = oApp.Workbooks.Open(targetPath)
 
  Set cmps = targetBook.VBProject.VBComponents
  Set oFdr = oFso.getFolder(sourcePath)
 
  For Each fl In oFdr.Files
    sExt = LCase(oFso.GetExtensionName(fl))
   
    If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
      If toCRLF Then Call lfToCrlf(fl)
      Call targetBook.VBProject.VBComponents.Remove(targetBook.VBProject.VBComponents(oFso.GetBaseName(fl)))
      Call targetBook.VBProject.VBComponents.Import(fl)
    End If
  Next
  targetBook.Save
  targetBook.Close
  oApp.Quit
End Sub

Function getFixedPath(ext)
    Dim oFso
    Dim scriptPath
    Dim targetPath
    Dim sorcePath
    Dim parentPath
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    parentName = oFso.getFilename(parentPath)
    
    sourcePath = oFso.buildPath(parentPath, "src")
    targetPath = oFso.buildPath(parentPath, parentName & "." & ext)
    
    getFixedPath = Array(parentPath, sourcePath, targetPath)
End Function
