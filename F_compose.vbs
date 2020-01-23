Const isFixedMode = True
Const toCRLF = True

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_Document = 100

Const xlExcel9795 = 43                    ' //.xls 97-2003 format in Excel 2003 or prev
Const xlExcel8 = 56                       ' //.xls 97-2003 format in Excel 2007
Const xlTemplate = 17                     ' //.xlt
Const xlAddIn = 18                        ' //.xla
Const xlExcel12 = 50                      ' //.xlsb
Const xlOpenXMLWorkbookMacroEnabled = 52  ' //.xlsm
Const xlOpenXMLTemplateMacroEnabled = 53  ' //.xltm
Const xlOpenXMLAddIn = 55                 ' //.xlam

Call composeAll

Sub composeAll()
  'import excel macro module
    
    On Error Resume Next
    Dim oApp
    Dim oFso
    Dim sExt
    Dim sourcePath
    Dim parentPath
    Dim targetName
    Dim targetPath
    Dim targetBook
    
    Set oApp = CreateObject("Excel.Application")
    oApp.DisplayAlerts = False
    oApp.EnableEvents = False
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    If isFixedMode Then
        tmp = getFixedPath
        
        parentPath = tmp(0)
        sourcePath = tmp(1)
    'targetPath = tmp(2)
    Else
        sourcePath = getFolderPath
        If sourcePath = "" Then
            MsgBox "exit this script"
            Exit Sub
        End If
        parentPath = oFso.getParentFolderName(sourcePath)
        
    End If
    targetName = oFso.getFilename(parentPath) & ".xlsm"
    binPath = oFso.buildPath(parentPath, "bin")
    If Not oFso.FolderExists(binPath) Then oFso.createFolder (binPath)
    targetPath = oFso.buildPath(binPath, targetName)
    
    If oFso.FileExists(targetPath) Then
        Call cleanAll(targetPath)
    Else
        Set targetBook = oApp.Workbooks.Add
        Call targetBook.SaveAs(targetPath, xlOpenXMLWorkbookMacroEnabled)
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
    lb=lbound(txts)
    ub=ubound(txts)
    do while (txts(ub)="" or txts(ub)=chr(13)) and lb<=ub 
        ub=ub-1
    loop 

    Set oStm = oFso.createtextfile(pn)
    For i=lb to ub
        txt=txts(i)
        If Right(txt, 1) = Chr(13) Then
            txt = Left(txt, Len(txt) - 1)
        End If
        oStm.WriteLine (txt)
    Next
    oStm.Close
End Sub

Function getFixedPath()
    Dim oFso
    Dim scriptPath
    Dim targetPath
    Dim sorcePath
    Dim parentPath
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    parentName = oFso.getFilename(parentPath)
    
    sourcePath = oFso.buildPath(parentPath, "src")
    targetPath = oFso.buildPath(parentPath, "bin" & "\" & parentName & ".xlsm")
    
    getFixedPath = Array(parentPath, sourcePath, targetPath)
End Function

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
            Call targetBook.VBProject.VBComponents.Import(fl)
        End If
    Next
    targetBook.Save
    targetBook.Close
    oApp.Quit
End Sub