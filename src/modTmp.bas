Attribute VB_Name = "modTmp"
Sub addBtn(rn, mn, Optional cn = "run", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.Name
    If bn = "" Then bn = ThisWorkbook.Name
    
    Set rg = Workbooks(bn).Sheets(sn).Range(rn)
    Set btn = Workbooks(bn).Sheets(sn).Buttons.Add(rg.Left, rg.Top, rg.width, rg.Height)
    btn.OnAction = mn
    btn.Caption = cn
End Sub
