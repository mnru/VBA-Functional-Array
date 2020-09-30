Attribute VB_Name = "modRng"
Option Base 0
Option Explicit

Enum rowColumn
    faRow = 1
    faColumn = 2
End Enum

Public Function TLookup(key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional otherwise = Empty, Optional bkn = "") As Variant
    Dim num, ret, bkn0
    bkn0 = ActiveWorkbook.Name
    If bkn = "" Then bkn = ThisWorkbook.Name
    Workbooks(bkn).Activate
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
    Workbooks(bkn0).Activate
    Exit Function
lnError:
    Debug.Print Err.Description
    TLookup = Empty
    Workbooks(bkn).Activate
End Function


Public Sub TSetUp(vl, key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional bkn = "")
    bkn0 = ActiveWorkbook.Name
    If bkn = "" Then bkn = ThisWorkbook.Name
    Workbooks(bkn).Activate
    
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    Range(tbl & "[" & targetCol & "]")(WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0), 1).Value = vl
    Workbooks(bkn0).Activate
    Exit Sub
lnError:
    Debug.Print Err.Description
End Sub

Sub layAryAt(ary, r, c, Optional rc As rowColumn = rowColumn.faRow, Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ActiveWorkbook.Name
    n = lenAry(ary)
    Select Case rc
        Case rowColumn.faRow
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(1, n) = ary
        Case rowColumn.faColumn
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(n, 1) = Application.WorksheetFunction.Transpose(ary)
        Case Else
    End Select
End Sub

Function rangeToAry(rg, Optional rc As rowColumn = rowColumn.faRow, Optional num = 1)
    Dim ret, tmp
    tmp = rg
    If Not IsArray(tmp) Then
        ret = Array(tmp)
    Else
        With Application.WorksheetFunction
            Select Case rc
                Case rowColumn.faRow
                    ret = .Index(tmp, num, 0)
                Case rowColumn.faColumn
                    ret = .Transpose(.Index(tmp, 0, num))
                Case Else
            End Select
        End With
        If dimAry(ret) = 0 Then
            ret = Array(tmp)
        End If
    End If
    rangeToAry = ret
End Function

Function rangeToArys(rg, Optional rc As rowColumn = rowColumn.faRow)
    Dim ret, tmp
    Dim num As Long, i As Long
    tmp = rg
    If Not IsArray(tmp) Then
        ret = Array(tmp)
    Else
        If rc = rowColumn.faColumn Then
            tmp = Application.WorksheetFunction.Transpose(tmp)
        End If
        If dimAry(tmp) <= 1 Then
            ret = Array(tmp)
        Else
            num = lenAry(tmp)
            ReDim ret(1 To num)
            For i = 1 To num
                ret(i) = Application.WorksheetFunction.Index(tmp, i, 0)
            Next i
        End If
    End If
    rangeToArys = ret
End Function
