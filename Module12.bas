Attribute VB_Name = "Module12"
Sub NKtoNKC()
'
    Sheets("NK").Select
    
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(R3C1:R10000C1,"">0"")"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(NKC_sodongNK)"
    Range("K3").Select
    If Range("M1") > Range("N1") Then
 
 MsgBox ("NKC KHONG DU DONG")
    
Else
 Range("A3:E850").Select
    Selection.Copy
    Sheets("NKC").Select
    Range("A13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("NK").Select
 Range("F3:I850").Select
    Selection.Copy
    Sheets("NKC").Select
    Range("I13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

End If
    Sheets("NK").Select
    Range("K3").Select
    Sheets("NKC").Select
    Range("E10").Select
End Sub



