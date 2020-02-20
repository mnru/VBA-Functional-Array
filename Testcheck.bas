Attribute VB_Name = "Testcheck"

Sub testcheckA_
a = 5 + 4
Assert a , 9
b = 4
a = b
b = a
bol1 = 3 < 4
AssertTrue bol1
bol2 = 4 < 3
AssertFalse bol2
x = mkSeq(6,-5,3)
y = mapA("calc",x,5,"*")
Assert toString(y),"[-25,-10,5,20,35,50]"
pl = Array(2,0,3,1,0)
str1 = polyStr(pl)
Assert str1 , 2X^4 +3X^2 +X'
lp = revAry(pl)
sp = Array(2,3)
sp = Array(3,5)
x = mkMArySeq(sp,3,5)
sp = Array(2,3,4)
x = mkMAryseq(sp)
y = mapMA("calc",x,5,"*")
Assert toString(y),"[5,10,15,20;" & vbCrLf & " 25,30,35,40;" & vbCrLf & " 45,50,55,60;;" & vbCrLf & "" & vbCrLf & " 65,70,75,80;" & vbCrLf & " 85,90,95,100;" & vbCrLf & " 105,110,115,120]"
pi_quarter = atn(1)
z = pi_quarter * 4
root2 = sqr(2)
Assert root2 , 1.4142135623731
Set dic = mkdic("a",1,"b",2)
Set clc = mkclc(1,2,3)
Assert toString(clc),"Clc(1,2,3)"
Set tmp = clc
Assert toString(tmp),"Clc(1,2,3)"
Set tmp1 = mkclc(clc,4,sp)
Set tmp = dic
Set tmp2 = mkdic("a",dic,"b",clc)
Assert toString(tmp2),"Dic('a'=>Dic('a'=>1,'b'=>2),'b'=>Clc(1,2,3))"
End sub
