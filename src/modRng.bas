Attribute VB_Name = "modRng"

Function rangeToAry(rg, Optional rc As String = "r", Optional num = 1)
    Dim ret
    With Application.WorksheetFunction
        Select Case LCase(rc)
            Case "r"
                ret = .Index(rg, num, 0)
            Case "c"
                ret = .Transpose(.Index(rg, 0, num))
            Case Else
        End Select
    End With
    If dimAry(ret) = 0 Then
        ret = Array(rg)
    End If
    rangeToAry = ret
End Function

Function rangeToArys(rg, Optional rc As String = "r")
    Dim ret
    Select Case LCase(rc)
        Case "r"
            tmp = rg
        Case "c"
            tmp = Application.WorksheetFunction.Transpose(rg)
        Case Else
    End Select
    If dimAry(tmp) <= 1 Then
        ret = Array(tmp)
    Else
        num = lenAry(tmp)
        ReDim ret(1 To num)
        For i = 1 To num
            ret(i) = Application.WorksheetFunction.Index(tmp, i, 0)
        Next i
    End If
    rangeToArys = ret
End Function

Public Function TLookup(key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional otherwise = Null) As Variant
    bkn = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    num = WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0)
    If num = 0 Then
        ret = otherwise
    Else
        ret = Range(tbl & "[" & targetCol & "]")(num, 1)
    End If
    TLookup = ret
    Workbooks(bkn).Activate
    Exit Function
lnError:
    Debug.Print Err.Description
    TLookup = Null
    Workbooks(bkn).Activate
End Function

Public Sub TSetUp(vl, key, tbl As String, targetCol As String, Optional sourceCol As String = "")
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    Range(tbl & "[" & targetCol & "]")(WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0), 1).Value = vl
    Exit Sub
lnError:
    Debug.Print Err.Description
End Sub

Sub layAryAt(ary, r, c, Optional rc = "r", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ActiveWorkbook.Name
    n = lenAry(ary)
    Select Case LCase(rc)
        Case "r"
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(1, n) = ary
        Case "c"
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(n, 1) = Application.WorksheetFunction.Transpose(ary)
        Case Else
    End Select
End Sub



