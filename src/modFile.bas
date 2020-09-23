Attribute VB_Name = "modFile"
'Enum MsoFileDialogType
' msoFileDialogOpen = 1
' msoFileDialogSaveAs = 2
' msoFileDialogFilePicker = 3
' msoFileDialogFolderPicker = 4
'End Enum

Enum fileSelectType
    singleFile = 1
    multiFiles = 2
    singleFolder = 3
End Enum

Function getFileByDialog(Optional dialogType As fileSelectType = multiFiles, Optional title As String = "", _
    Optional initFolder As String = "", Optional initialFile As String = "", Optional extentions As String = "All Files,*.*")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ret
    Dim tmp
    Dim mstype
    Set dlg = Application.FileDialog(dialogType)
    ' If initFolder = "" Then
    '  initFolder = ThisWorkbook.Path
    'initFolder = CurDir
    ' End If
    If title = "" Then
        Select Case dialogType
            Case singleFile
                title = "Select file"
                mstype = 3
            Case multiFiles
                title = "Select files"
                mstype = 3
            Case singleFolder
                title = "Select folder"
                mstype = 4
            Case Else
        End Select
    End If
    MultiSelect = dialogType = fileSelectType.multiFiles
    With Application.FileDialog(mstype)
        .title = title
        .AllowMultiSelect = MultiSelect
        .InitialFileName = fso.buildpath(initFolder, initFile)
        If extentions <> "" And dialogType <> fileSelectType.singleFolder Then
            exts = Split(extentions, ",")
            For i = LBound(exts) To UBound(exts) Step 2
                .Filters.Add exts(i), exts(i + 1)
            Next i
        End If
        If .Show = True Then
            n = .SelectedItems.Count
            ReDim tmp(1 To n)
            For i = 1 To n
                tmp(i) = .SelectedItems(i)
            Next i
            If MultiSelect Then
                ret = tmp
            Else
                ret = tmp(LBound(tmp))
            End If
        Else
            ret = False
        End If
    End With
    getFileByDialog = ret
End Function

Function getFilePart(pn, prm) As String
    Dim ret As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Select Case LCase(prm)
        Case "parent": ret = fso.GetParentFolderName(pn)
        Case "file": ret = fso.GetFileName(pn)
        Case "base": ret = fso.GetBaseName(pn)
        Case "ext": ret = fso.GetExtensionName(pn)
        Case "drive": ret = fso.GetDriveName(pn)
        Case "abs": ret = fso.GetAbsolutePathName(pn)
        Case Else:
    End Select
    getFilePart = ret
End Function

Function joinOneDelm(a, b, delm)
    If Right(a, 1) = delm Then a = Left(a, Len(a) - 1)
    If Left(b, 1) = delm Then b = Right(b, Len(b) - 1)
    ret = a & delm & b
    joinOneDelm = ret
End Function

Function buildPaths(ParamArray prms())
    ary = prms
    ret = reduceA("joinOneDelm", ary, "\")
    buildPaths = ret
End Function

Function getFileAry(sFolder, Optional fileFolder = "file")
    Dim obj, ret, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fdr = fso.GetFolder(sFolder)
    Select Case LCase(fileFolder)
        Case "file"
            Set obj = fdr.Files
        Case "folder"
            Set obj = fdr.SubFolders
        Case Else
    End Select
    num = obj.Count
    ReDim ret(1 To num)
    i = 1
    For Each elm In obj
        ret(i) = elm.Path
        i = i + 1
    Next elm
    getFileAry = ret
End Function

Sub testFileAry()
    Path = ThisWorkbook.Path
    x = getFileAry(Path)
    y = getFileAry(Path, "folder")
    printAry (x)
    printAry (y)
End Sub

Sub testDialog()
    x = getFileByDialog(singleFolder)
    printAry (x)
    Stop
    y = getFileByDialog(multiFiles, , , , "All files,*.*,Excel files,*.xls*,Text files,*.txt;*.csv")
    printAry y
    z = mapA("getFilePart", y, "file")
    printAry z
End Sub

Sub testbuild()
    x = buildPaths("c:\", "\windows", "system")
    outPut (x)
End Sub
