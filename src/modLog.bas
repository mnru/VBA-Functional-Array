Attribute VB_Name = "modLog"
Option Base 0
Option Explicit

Sub outPut(Optional msg As Variant = "", Optional crlf As Boolean = True)
    If crlf Then
        Debug.Print msg
    Else
        Debug.Print msg;
    End If
End Sub

Function printTime(fnc As String, ParamArray argAry() As Variant)
    Dim ary, fnAry
    Dim etime As Double
    Dim stime As Double
    Dim secs As Double
    ary = argAry
    fnAry = prmAry(fnc, ary)
    stime = Timer
    printTime = evalA(fnAry)
    etime = Timer
    secs = etime - stime
    Call outPut(fnc & " - " & secToHMS(secs), True)
End Function

Sub printAry(ary, Optional qt As String = "'", Optional fm As String = "", Optional lcr As AlignDirection = AlignDirection.faRight, Optional width As Long = 0)
    Call outPut(toString(ary, qt, fm, lcr, width), True)
End Sub

Sub printSimpleAry(ary, Optional flush As Long = 1000, Optional qt As String = "'")
    Dim i As Long, aryNum As Long
    Dim sp, lsp
    Dim ret As String, dlm As String
    Dim idx, idx0, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, "L")
    aryNum = getAryNum(ary)
    If aryNum = 0 Then
        Call outPut("[]", False)
    Else
        ret = "["
        For i = 0 To aryNum - 1
            idx0 = mkIndex(i, sp)
            idx = calcAry(idx0, lsp, "+")
            vl = getElm(ary, idx)
            If TypeName(vl) = "String" Then vl = qt & vl & qt
            dlm = getDlm(sp, idx0)
            ret = ret & vl & dlm
            If i Mod flush = 0 Then
                Call outPut(ret, False)
                ret = ""
            End If
        Next i
    End If
    Call outPut(ret, True)
End Sub

Sub print1DAry(ary, Optional flush As Long = 1000)
    Dim ret As String, dlm As String, elm As String
    Dim lb1 As Long, ub1 As Long
    Dim cnt As Long, i1 As Long
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    cnt = 1
    For i1 = lb1 To ub1
        elm = CStr(ary(i1))
        If i1 < ub1 Then
            dlm = ","
        Else
            dlm = "]"
        End If
        ret = ret & elm & dlm
        If cnt Mod flush = 0 Then
            Call outPut(ret, False)
            ret = ""
        End If
        cnt = cnt + 1
    Next i1
    Call outPut(ret, True)
End Sub

Sub print2DAry(ary, Optional flush As Long = 1000)
    Dim ret As String, dlm As String, elm As String
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim cnt As Long, i1 As Long, i2 As Long
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    cnt = 1
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            elm = CStr(ary(i1, i2))
            If i2 < ub2 Then
                dlm = ","
            ElseIf i1 < ub1 Then
                dlm = ";" & vbCrLf
            Else
                dlm = "]"
            End If
            ret = ret & elm & dlm
            If cnt Mod flush = 0 Then
                Call outPut(ret, False)
                ret = ""
            End If
            cnt = cnt + 1
        Next i2
    Next i1
    Call outPut(ret, True)
End Sub

Sub print3DAry(ary, Optional flush As Long = 1000)
    Dim ret As String, dlm As String, elm As String
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, lb3 As Long, ub3 As Long
    Dim cnt As Long, i1 As Long, i2 As Long, i3 As Long
    ret = "["
    lb1 = LBound(ary, 1): ub1 = UBound(ary, 1)
    lb2 = LBound(ary, 2): ub2 = UBound(ary, 2)
    lb3 = LBound(ary, 3): ub3 = UBound(ary, 3)
    cnt = 1
    For i1 = lb1 To ub1
        For i2 = lb2 To ub2
            For i3 = lb3 To ub3
                elm = CStr(ary(i1, i2, i3))
                If i3 < ub3 Then
                    dlm = ","
                ElseIf i2 < ub2 Then
                    dlm = ";" & vbCrLf
                ElseIf i1 < ub1 Then
                    dlm = ";;" & vbCrLf & vbCrLf
                Else
                    dlm = "]"
                End If
                ret = ret & elm & dlm
                If cnt Mod flush = 0 Then
                    Call outPut(ret, False)
                    ret = ""
                End If
                cnt = cnt + 1
            Next i3
        Next i2
    Next i1
    Call outPut(ret, True)
End Sub
