Attribute VB_Name = "modMulti"
Function getMAryAt(ary As Variant, pos As Variant, Optional base As Long = 1)
    Dim lNum As Long
    Dim lsp, bs, idx1, idx2, ret
    lsp = getAryShape(ary, faLower)
    lNum = lenAry(lsp)
    bs = mkSameAry(base, lNum)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    ret = getElm(ary, idx2)
    getMAryAt = ret
End Function

Sub setMAryAt(ByRef ary As Variant, pos As Variant, vl As Variant, Optional base As Long = 1)
    Dim lNum As Long
    Dim lsp, bs, idx1, idx2
    lsp = getAryShape(ary, faLower)
    lNum = lenAry(lsp)
    bs = mkSameAry(base, lNum)
    idx1 = calcAry(pos, bs, "-")
    idx2 = calcAry(idx1, lsp, "+")
    Call setElm(vl, ary, idx2)
End Sub

Function mkIndex(num As Long, shape, Optional lshape = Empty)
    Dim lNum As Long, i As Long, p As Long, r As Long
    lNum = lenAry(shape)
    ReDim ret(0 To lNum - 1)
    r = num
    For i = lNum To 1 Step -1
        p = getAryAt(shape, i)
        Call setAryAt(ret, i, r Mod p)
        r = r \ p
    Next i
    If Not IsEmpty(lshape) Then
        For i = 1 To lNum
            Call setAryAt(ret, i, getAryAt(ret, i) + getAryAt(lshape, i))
        Next i
    End If
    mkIndex = ret
End Function

Function getElm(ByRef ary, idx)
    Dim ret
    Dim lb As Long
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: assign_ ret, ary(idx(lb))
        Case 2: assign_ ret, ary(idx(lb), idx(lb + 1))
        Case 3: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2))
        Case 4: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3))
        Case 5: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4))
        Case 6: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5))
        Case 7: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6))
        Case 8: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7))
        Case 9: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8))
        Case 10: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9))
        Case 11: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10))
        Case 12: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11))
        Case 13: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12))
        Case 14: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13))
        Case 15: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14))
        Case 16: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15))
        Case 17: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16))
        Case 18: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17))
        Case 19: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18))
        Case 20: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19))
        Case 21: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20))
        Case 22: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21))
        Case 23: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22))
        Case 24: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23))
        Case 25: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24))
        Case 26: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25))
        Case 27: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26))
        Case 28: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27))
        Case 29: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28))
        Case 30: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29))
        Case 31: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30))
        Case 32: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31))
        Case 33: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32))
        Case 34: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33))
        Case 35: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34))
        Case 36: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35))
        Case 37: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36))
        Case 38: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37))
        Case 39: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38))
        Case 40: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39))
        Case 41: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40))
        Case 42: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41))
        Case 43: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42))
        Case 44: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43))
        Case 45: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44))
        Case 46: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45))
        Case 47: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46))
        Case 48: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47))
        Case 49: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48))
        Case 50: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49))
        Case 51: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50))
        Case 52: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51))
        Case 53: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52))
        Case 54: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53))
        Case 55: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54))
        Case 56: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55))
        Case 57: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56))
        Case 58: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57))
        Case 59: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58))
        Case 60: assign_ ret, ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58), idx(lb + 59))
        Case Else:
    End Select
    assign_ getElm, ret
End Function

Sub setElm(vl, ary, idx)
    Dim lb As Long
    lb = LBound(idx)
    Select Case lenAry(idx)
        Case 1: assign_ ary(idx(lb)), vl
        Case 2: assign_ ary(idx(lb), idx(lb + 1)), vl
        Case 3: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2)), vl
        Case 4: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3)), vl
        Case 5: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4)), vl
        Case 6: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5)), vl
        Case 7: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6)), vl
        Case 8: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7)), vl
        Case 9: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8)), vl
        Case 10: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9)), vl
        Case 11: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10)), vl
        Case 12: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11)), vl
        Case 13: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12)), vl
        Case 14: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13)), vl
        Case 15: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14)), vl
        Case 16: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15)), vl
        Case 17: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16)), vl
        Case 18: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17)), vl
        Case 19: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18)), vl
        Case 20: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19)), vl
        Case 21: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20)), vl
        Case 22: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21)), vl
        Case 23: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22)), vl
        Case 24: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23)), vl
        Case 25: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24)), vl
        Case 26: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25)), vl
        Case 27: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26)), vl
        Case 28: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27)), vl
        Case 29: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28)), vl
        Case 30: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29)), vl
        Case 31: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30)), vl
        Case 32: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31)), vl
        Case 33: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32)), vl
        Case 34: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33)), vl
        Case 35: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34)), vl
        Case 36: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35)), vl
        Case 37: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36)), vl
        Case 38: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37)), vl
        Case 39: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38)), vl
        Case 40: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39)), vl
        Case 41: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40)), vl
        Case 42: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41)), vl
        Case 43: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42)), vl
        Case 44: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43)), vl
        Case 45: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44)), vl
        Case 46: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45)), vl
        Case 47: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46)), vl
        Case 48: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47)), vl
        Case 49: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48)), vl
        Case 50: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49)), vl
        Case 51: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50)), vl
        Case 52: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51)), vl
        Case 53: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52)), vl
        Case 54: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53)), vl
        Case 55: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54)), vl
        Case 56: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55)), vl
        Case 57: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56)), vl
        Case 58: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57)), vl
        Case 59: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58)), vl
        Case 60: assign_ ary(idx(lb), idx(lb + 1), idx(lb + 2), idx(lb + 3), idx(lb + 4), idx(lb + 5), idx(lb + 6), idx(lb + 7), idx(lb + 8), idx(lb + 9), idx(lb + 10), idx(lb + 11), idx(lb + 12), idx(lb + 13), idx(lb + 14), idx(lb + 15), idx(lb + 16), idx(lb + 17), idx(lb + 18), idx(lb + 19), idx(lb + 20), idx(lb + 21), idx(lb + 22), idx(lb + 23), idx(lb + 24), idx(lb + 25), idx(lb + 26), idx(lb + 27), idx(lb + 28), idx(lb + 29), idx(lb + 30), idx(lb + 31), idx(lb + 32), idx(lb + 33), idx(lb + 34), idx(lb + 35), idx(lb + 36), idx(lb + 37), idx(lb + 38), idx(lb + 39), idx(lb + 40), idx(lb + 41), idx(lb + 42), idx(lb + 43), idx(lb + 44), idx(lb + 45), idx(lb + 46), idx(lb + 47), idx(lb + 48), idx(lb + 49), idx(lb + 50), idx(lb + 51), idx(lb + 52), idx(lb + 53), idx(lb + 54), idx(lb + 55), idx(lb + 56), idx(lb + 57), idx(lb + 58), idx(lb + 59)), vl
        Case Else:
    End Select
End Sub

Sub setAryMbyS(mAry, sAry)
    Dim i As Long, aNum As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
    aNum = getAryNum(mAry)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = getAryAt(sAry, i, 0)
        Call setElm(vl, mAry, idx)
    Next i
End Sub

Function getArySbyM(mAry, Optional bs As Long = 0)
    Dim aNum As Long, i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
    aNum = getAryNum(mAry)
    ReDim ret(bs To bs + aNum - 1)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = getElm(mAry, idx)
        Call setAryAt(ret, i, vl, 0)
    Next i
    getArySbyM = ret
End Function

Function reshapeAry(ary, sp, Optional bs As Long = 0)
    Dim ret
    ret = mkMAry(sp, bs)
    Call setAryMbyS(ret, ary)
    reshapeAry = ret
End Function

Function calcMAry(ary1, ary2, symbol As String, Optional bs As Long = 0)
    Dim aNum As Long, dm As Long, i As Long
    Dim ret, vl, sp1, sp2, lsp1, lsp2, lsp0, idx, idx0, idx1, idx2
    sp1 = getAryShape(ary1)
    sp2 = getAryShape(ary2)
    lsp1 = getAryShape(ary1, faLower)
    lsp2 = getAryShape(ary2, faLower)
    aNum = getAryNum(ary1)
    dm = lenAry(sp1)
    ret = mkMAry(sp1, bs)
    lsp0 = mkSameAry(bs, dm)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp1)
        idx1 = calcAry(idx, lsp1, "+")
        idx2 = calcAry(idx, lsp2, "+")
        idx0 = calcAry(idx, lsp0, "+")
        vl = calc_(getElm(ary1, idx1), getElm(ary2, idx2), symbol)
        Call setElm(vl, ret, idx0)
    Next i
    calcMAry = ret
End Function

Function mkMAry(sp, Optional bs As Long = 0)
    Dim lNum As Long
    Dim ub, lb
    lNum = lenAry(sp)
    ub = calcAry(sp, mkSameAry(bs - 1, lNum), "+")
    lb = LBound(ub)
    Dim ret
    Select Case lNum
        Case 1: ReDim ret(bs To ub(lb))
        Case 2: ReDim ret(bs To ub(lb), bs To ub(lb + 1))
        Case 3: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2))
        Case 4: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3))
        Case 5: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4))
        Case 6: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5))
        Case 7: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6))
        Case 8: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7))
        Case 9: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8))
        Case 10: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9))
        Case 11: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10))
        Case 12: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11))
        Case 13: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12))
        Case 14: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13))
        Case 15: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14))
        Case 16: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15))
        Case 17: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16))
        Case 18: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17))
        Case 19: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18))
        Case 20: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19))
        Case 21: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20))
        Case 22: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21))
        Case 23: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22))
        Case 24: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23))
        Case 25: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24))
        Case 26: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25))
        Case 27: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26))
        Case 28: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27))
        Case 29: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28))
        Case 30: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29))
        Case 31: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30))
        Case 32: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31))
        Case 33: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32))
        Case 34: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33))
        Case 35: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34))
        Case 36: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35))
        Case 37: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36))
        Case 38: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37))
        Case 39: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38))
        Case 40: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39))
        Case 41: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40))
        Case 42: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41))
        Case 43: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42))
        Case 44: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43))
        Case 45: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44))
        Case 46: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44), bs To ub(lb + 45))
        Case 47: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44), bs To ub(lb + 45), bs To ub(lb + 46))
        Case 48: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44), bs To ub(lb + 45), bs To ub(lb + 46), bs To ub(lb + 47))
        Case 49: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44), bs To ub(lb + 45), bs To ub(lb + 46), bs To ub(lb + 47), bs To ub(lb + 48))
        Case 50: ReDim ret(bs To ub(lb), bs To ub(lb + 1), bs To ub(lb + 2), bs To ub(lb + 3), bs To ub(lb + 4), bs To ub(lb + 5), bs To ub(lb + 6), bs To ub(lb + 7), bs To ub(lb + 8), bs To ub(lb + 9), bs To ub(lb + 10), bs To ub(lb + 11), bs To ub(lb + 12), bs To ub(lb + 13), bs To ub(lb + 14), bs To ub(lb + 15), bs To ub(lb + 16), bs To ub(lb + 17), bs To ub(lb + 18), bs To ub(lb + 19), bs To ub(lb + 20), bs To ub(lb + 21), bs To ub(lb + 22), bs To ub(lb + 23), bs To ub(lb + 24), bs To ub(lb + 25), bs To ub(lb + 26), bs To ub(lb + 27), bs To ub(lb + 28), bs To ub(lb + 29), bs To ub(lb + 30), bs To ub(lb + 31), bs To ub(lb + 32), bs To ub(lb + 33), bs To ub(lb + 34), bs To ub(lb + 35), bs To ub(lb + 36), bs To ub(lb + 37), bs To ub(lb + 38), bs To ub(lb + 39), bs To ub(lb + 40), bs To ub(lb + 41), bs To ub(lb + 42), bs To ub(lb + 43), bs To ub(lb + 44), bs To ub(lb + 45), bs To ub(lb + 46), bs To ub(lb + 47), bs To ub(lb + 48), bs To ub(lb + 49))
        Case Else:
    End Select
    mkMAry = ret
End Function

Sub setMArySeq(ary, Optional first = 1, Optional step = 1)
    Dim aNum As Long, i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, faLower)
    aNum = getAryNum(ary)
    vl = first
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        'vl = first + i0 * step
        Call setElm(vl, ary, idx)
        vl = vl + step
    Next i
End Sub

Function mkMSameAry(vl, sp, Optional bs As Long = 0)
    Dim lNum As Long, aNum As Long, i As Long
    Dim ret, lsp, idx
    ret = mkMAry(sp, bs)
    lNum = lenAry(sp)
    sp = getAryShape(ret)
    lsp = mkSameAry(bs, lNum)
    aNum = getAryNum(ret)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        Call setElm(vl, ret, idx)
    Next i
    mkMSameAry = ret
End Function

Function mkMSeq(sp, Optional first = 1, Optional step = 1, Optional bs As Long = 0)
    Dim ret
    ret = mkMAry(sp, bs)
    Call setMArySeq(ret, first, step)
    mkMSeq = ret
End Function

Public Function mapMA(fnc As String, mAry As Variant, ParamArray argAry() As Variant) As Variant
    Dim ary, sp, lsp, fnAry, ret, idx, idx0, vl
    Dim aNum As Long
    ary = argAry
    sp = getAryShape(mAry)
    lsp = getAryShape(mAry, faLower)
    ret = mkMAry(sp)
    fnAry = prmAry(fnc, Empty, ary)
    aNum = getAryNum(mAry)
    Dim i As Long
    For i = 1 To aNum
        idx0 = mkIndex(i, sp)
        idx = calcAry(idx0, lsp, "+")
        Call setAryAt(fnAry, 2, getElm(mAry, idx))
        vl = evalA(fnAry)
        Call setElm(vl, ret, idx0)
    Next i
    mapMA = ret
End Function

Sub setAryMByF(ary, fnObj)
    Dim aNum As Long
    Dim i As Long
    Dim sp, lsp, idx, vl
    sp = getAryShape(ary)
    lsp = getAryShape(ary, faLower)
    aNum = getAryNum(ary)
    For i = 0 To aNum - 1
        idx = mkIndex(i, sp, lsp)
        vl = applyF(i, fnObj)
        Call setElm(vl, ary, idx)
    Next i
End Sub

Public Function foldF(fnObj, seq As Variant, init As Variant) As Variant
    Dim ret, elm
    ret = init
    For Each elm In seq
        ret = applyF(Array(ret, elm), fnObj, True)
    Next elm
    foldF = ret
End Function

Public Function reduceF(fnObj, seq As Variant) As Variant
    Dim init, seq1, ret
    init = getAryAt(seq, 1)
    seq1 = dropAry(seq, 1)
    ret = foldF(fnObj, seq1, init)
    reduceF = ret
End Function

Function applyF(vl, fnObj, Optional argAsAry = False)
    Dim ret, fnAry, arity
    Dim n As Long, i As Long
    fnAry = getAryAt(fnObj, 2)
    arity = getAryAt(fnObj, 1)
    If Not argAsAry Then
        Call setAryAt(fnAry, getAryAt(arity, 1), vl, 0)
    Else
        n = lenAry(arity)
        For i = 1 To n
            Call setAryAt(fnAry, getAryAt(arity, i), getAryAt(vl, i), 0)
        Next i
    End If
    ret = evalA(fnAry)
    applyF = ret
End Function

Function applyFs(vl, fnObjs, Optional argAsAry = False)
    Dim ret, fnObj
    ret = vl
    For Each fnObj In fnObjs
        ret = applyF(ret, fnObj, argAsAry)
    Next fnObj
    applyFs = ret
End Function

Function getArity(ary)
    Dim ret, elm
    ret = 0
    For Each elm In ary
        If IsNumeric(elm) Then
            ret = ret + 1
        Else
            Exit For
        End If
    Next elm
    getArity = ret
End Function

Function mkF(ParamArray argArys())
    Dim n As Long
    Dim ary, fnAry, arity
    ary = argArys
    n = getArity(ary)
    arity = takeAry(ary, n)
    fnAry = dropAry(ary, n)
    mkF = Array(arity, fnAry)
End Function

Function zipApplyF(fnObj, ParamArray argAry())
    Dim arys, x, ret
    arys = argAry
    x = zipAry(arys)
    ret = mapA("applyF", x, fnObj, True)
    zipApplyF = ret
End Function
