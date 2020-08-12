Const fixedMode = True
Const useExcelDialog = True
Dim targetExt
targetExt = "xlsm"

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_Document = 100

Const xlExcel9795 = 43                   ' //.xls 97-2003 format in Excel 2003 or prev
Const xlExcel8 = 56                      ' //.xls 97-2003 format in Excel 2007
Const xlTemplate = 17                    ' //.xlt
Const xlAddIn = 18                       ' //.xla
Const xlExcel12 = 50                     ' //.xlsb
Const xlOpenXMLWorkbookMacroEnabled = 52 ' //.xlsm
Const xlOpenXMLTemplateMacroEnabled = 53 ' //.xltm
Const xlOpenXMLAddIn = 55                ' //.xlam

Call decomposeAll

Sub decomposeAll()
    'export excel macro module
    
    Dim oApp
    Dim oFso
    
    Dim module
    Dim modules
    Dim ext
    
    Dim parentPath
    Dim sourcePath
    Dim targetPath
    Dim sFilePath
    Dim TargetBook
    
    Dim bn
    Dim xn
    
    Set oApp = CreateObject("Excel.Application")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oShl = CreateObject("Shell.Application")
    oApp.DisplayAlerts = False
    oApp.EnableEvents = False
    
    If fixedMode Then
        tmp = getFixedPath(targetExt)
        parentPath = tmp(0)
        sourcePath = tmp(1)
        targetPath = tmp(2)
    Else
        If useExcelDialog Then
            targetPath = getFilePathByExcel
        Else
            targetPath = getFilePath
        End If
        If targetPath = "" Then
            MsgBox "exit this script"
            Exit Sub
        End If
        
        prn = oFso.GetParentFolderName(targetPath)
        bn = oFso.GetBaseName(targetPath)
        xn = oFso.GetExtensionName(targetPath)
        
        If Left(xn, 2) <> "xl" Then
            MsgBox "this file is not Excel File"
            Exit Sub
        End If
        
        parentPath = oFso.buildPath(prn, bn)
        sourcePath = oFso.buildPath(parentPath, "src")
    End If
    
    If Not oFso.FolderExists(parentPath) Then oFso.createFolder (parentPath)
    If Not oFso.FolderExists(sourcePath) Then oFso.createFolder (sourcePath)
    
    Set TargetBook = oApp.Workbooks.Open(targetPath)
    
    Set modules = TargetBook.VBProject.VBComponents
    
    For Each module In modules
        mExt = ""
        If (module.Type = vbext_ct_ClassModule) Then
            mExt = "cls"
        ElseIf (module.Type = vbext_ct_MSForm) Then
            mExt = "frm"
        ElseIf (module.Type = vbext_ct_StdModule) Then
            mExt = "bas"
        End If
        
        If mExt <> "" Then
            sFilePath = oFso.buildPath(sourcePath, module.Name & "." & mExt)
            Call module.Export(sFilePath)
            
        End If
    Next
    TargetBook.Close
    oApp.Quit
    MsgBox "Complete!"
End Sub

Function getFilePathByExcel()
    On Error Resume Next
    Set oApp = CreateObject("Excel.Application")
    oApp.Visible = True
    
    ret = oApp.GetOpenFilename("All files,*.*", 1, "select file")
    
    If ret = False Then ret = ""
    getFilePathByExcel = ret
    oApp.Quit
    
    Set oApp = Nothing
    
End Function


Function getFilePath()
    
    Dim oShl
    Dim oBrw
    Dim strPath
    On Error Resume Next
    Set oShl = WScript.CreateObject("Shell.Application")
    Set oBrw = oShl.BrowseForFolder(0, "Select Excel macro file", &H4000)
    
    If (oBrw Is Nothing) Then
        Err.Clear
        getFilePath = ""
    Else
        getFilePath = oBrw.Items.Item.Path
    End If
    
    Set oShl = Nothing
    Set oBrw = Nothing
    Err.Clear
    On Error GoTo 0
    
End Function

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
