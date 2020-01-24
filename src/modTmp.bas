Attribute VB_Name = "modTmp"
Sub addBtn(rn, mn, Optional cn = "run", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ThisWorkbook.Name
    
    Set rg = Workbooks(bn).Sheets(sn).Range(rn)
    Set btn = Workbooks(bn).Sheets(sn).Buttons.Add(rg.Left, rg.Top, rg.width, rg.Height)
    btn.OnAction = mn
    btn.Caption = cn
End Sub

'Enum MsoFileDialogType
'  msoFileDialogOpen = 1
'  msoFileDialogSaveAs = 2
'  msoFileDialogFilePicker = 3
'  msoFileDialogFolderPicker = 4
'End Enum

Function getFileByDialog(Optional dialogType As MsoFileDialogType = 4, Optional title As String = "", _
    Optional initFolder As String = "", Optional initialFile As String = "", Optional extentions As String = "all files,*.*", _
    Optional multiSelect As Boolean = False)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ret
    Dim tmp
    Set dlg = Application.FileDialog(dialogType)
    
'  If initFolder = "" Then
'    initFolder = ThisWorkbook.Path
    'initFolder = CurDir
    Print '  End If
    
    If title = "" Then
        
        Select Case dialogType
            Case msoFileDialogFolderPicker: title = "select folder"
            Case Else: title = "select file"
        End Select
    End If
    
    With Application.FileDialog(dialogType)
        If title <> "" Then .title = title
        .AllowMultiSelect = multiSelect
        .InitialFileName = fso.BuildPath(initFolder, initFile)
        
        If extentions <> "" And dialogType <> msoFileDialogFolderPicker Then
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
            If multiSelect Then
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
    Dim fso As FileSystemObject
    
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

Sub testDialog()
    
    x = getFileByDialog(msoFileDialogFolderPicker, , , , , True)
    printAry x
    Stop
    y = getFileByDialog(msoFileDialogFilePicker, , , , "all files,*.*,exel files,*.xls*,text files,*.txt;*.csv", True)
    printAry y
    Stop
    
    z = mapA("getfilepart", y, "file")
    
    printAry z
    
End Sub

