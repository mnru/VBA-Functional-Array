Attribute VB_Name = "modRng"
Option Base 0
Option Explicit

Public Function TLookup(key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional otherwise = Empty) As Variant
    bkn = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    Num = WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0)
    If Num = 0 Then
        ret = otherwise
    Else
        ret = Range(tbl & "[" & targetCol & "]")(Num, 1)
    End If
    TLookup = ret
    Workbooks(bkn).Activate
    Exit Function
lnError:
    Debug.Print Err.Description
    TLookup = Empty
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

Function rangeToAry(rg, Optional rc As String = "r", Optional Num = 1)
    Dim ret, tmp
    tmp = rg
    With Application.WorksheetFunction
        Select Case LCase(rc)
            Case "r"
                ret = .Index(tmp, Num, 0)
            Case "c"
                ret = .Transpose(.Index(tmp, 0, Num))
            Case Else
        End Select
    End With
    If dimAry(ret) = 0 Then
        ret = Array(tmp)
    End If
    rangeToAry = ret
End Function

Function rangeToArys(rg, Optional rc As String = "r")
    Dim ret, tmp
    Dim Num As Long, i As Long
    tmp = rg
    Select Case LCase(rc)
        Case "r"
            tmp = rg
        Case "c"
            tmp = Application.WorksheetFunction.Transpose(tmp)
        Case Else
    End Select
    If dimAry(tmp) <= 1 Then
        ret = Array(tmp)
    Else
        Num = lenAry(tmp)
        ReDim ret(1 To Num)
        For i = 1 To Num
            ret(i) = Application.WorksheetFunction.Index(tmp, i, 0)
        Next i
    End If
    rangeToArys = ret
End Function
