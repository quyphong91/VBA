Attribute VB_Name = "Module18"
Sub Auto_Open()
    
    Application.ScreenUpdating = False
' BAO SAO CHEP DU LIEU:
    Sheets("TTDN").Select
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(or(MID(CELL(""filename""),IF(ISERROR(FIND(""PHUCVN"",CELL(""filename"")))=TRUE,1,FIND(""PHUCVN"",CELL(""filename""))),6)=""PHUCVN"",MID(CELL(""filename""),IF(ISERROR(FIND(""TS-"",CELL(""filename"")))=TRUE,1,FIND(""TS-"",CELL(""filename""))),3)=""TS-""),0,1)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

If Selection.Value = 0 Then
    MsgBox ("Ban KHONG THE su dung Chuong Trinh Ke Toan Nay DO SAO CHEP KHONG DUNG QUY DINH")
    MsgBox ("Vui long lien he tac gia neu co nhu cau tiep tuc su dung!")
    DONG
Else
DIA_CHI

End If
End Sub
Sub DONG()
    
    Range("A65536").ClearContents
    ActiveWorkbook.Save
    Application.Quit
    
End Sub
Sub MST()

Sheets("TTDN").Select
If Range("J1") <> 0 Then
Range("C1").Select
MsgBox ("MA SO THUE co the SAI. Vui long kiem tra lai!")
End If
Range("C1").Select

End Sub

Sub DIA_CHI()

Sheets("TTDN").Select
Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(or(MID(R3C3,IF(ISERROR(FIND(""Qu"",R3C3))=TRUE,1,FIND(""Qu"",R3C3)),2)=""Qu"",MID(R3C3,IF(ISERROR(FIND(""Q."",R3C3))=TRUE,1,FIND(""Q."",R3C3)),2)=""Q."",MID(R3C3,IF(ISERROR(FIND(""Hu"",R3C3))=TRUE,1,FIND(""Hu"",R3C3)),2)=""Hu"",MID(R3C3,IF(ISERROR(FIND(""H."",R3C3))=TRUE,1,FIND(""H."",R3C3)),2)=""H.""),0,1)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

If Selection.Value <> 0 Then
Range("C1").Select
MsgBox ("DIA CHI Cong ty co the CHUA GO QUAN-HUYEN. Vui long kiem tra lai!")
End If
Range("A65536").ClearContents
Range("C1").Select

End Sub

Sub KIEMTRATHUE()

    Sheets("NK").Select
           
If Range("thang") = 1 Then
    Range("B2").Activate
    R = MsgBox("Thang 1: Co can kiem tra lai but toan 3338 phai nop da duoc DINH KHOAN chua khong?", vbYesNo, "LUU Y")
    If R = vbYes Then
    Sheets("Khac").Select
    Range("B2").Activate
    Exit Sub
    End If
End If
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "=+IF(OR(thang=3,thang=6,thang=9,thang=12),1,0)"
    Range("K1").Copy
    Range("K1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("K1") = 1 Then
    R = MsgBox("CUOI QUY: Co can kiem tra lai but toan 3334-335 QUY phai nop(neu co) da duoc DINH KHOAN chua khong?", vbYesNo, "LUU Y")
    If R = vbYes Then
    Sheets("Khac").Select
    Range("B2").Activate
    Exit Sub
    End If
End If
    
End Sub

