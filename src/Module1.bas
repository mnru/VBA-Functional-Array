Attribute VB_Name = "Module1"
Function min2(x, y)
min2 = IIf(x < y, x, y)
End Function
Function max2(x, y)
max2 = IIf(x > y, x, y)
End Function
Function mid3(x, y, z)
mid3 = IIf(max2(x, y) < z, max2(x, y), max2(min2(x, y), z))
End Function
Function qsortAry(ary)
ret = ary


End Function
