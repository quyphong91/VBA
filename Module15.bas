Attribute VB_Name = "Module15"
Sub INSO()

    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then

'Kiem tra truoc khi in
Sheets("TTDN").Select
  Range("H11").Activate
R = MsgBox("CHAC CHAN la muon IN TOAN BO khong?", vbYesNo, "LUU Y")
If R = vbNo Then
Exit Sub
End If

DANHSACHINSO
 
  Exit Sub
 Else
  Sheets("TTDN").Select
  Range("H11").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub

Sub DANHSACHINSO()
'Neu kiem tra ky roi thi in
' BANG CDSPS
If Range("TTDN!F2") = 1 Then
   KC_CDSPS
   Sheets("CDSPS").Select
     ActiveSheet.PrintOut
End If
' NHAT KY CHUNG
 If Range("TTDN!F3") = 1 Then
    KC_CDSPS
   Sheets("NKC").Select
     ActiveSheet.PrintOut
End If

Dim i
' SO CAI
If Range("TTDN!F4") = 1 Then
With Range("HTTK!A1")
    For i = 446 To 548
    If .Offset(i, 2) = 1 Then
        Range("SC!C15").Value = .Offset(i, 0)
          SOCAI
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If

'SO QUY 111
If Range("TTDN!F5") = 1 Then
With Range("HTTK!A1")
    For i = 552 To 560
    If .Offset(i, 2) = 1 Then
        Range("SQ111!F2").Value = .Offset(i, 0)
          SOQUY111
        ActiveSheet.PrintOut
End If
 Next i
    End With
End If
'SO QUY 112
If Range("TTDN!F6") = 1 Then
With Range("HTTK!A1")
    For i = 564 To 586
    If .Offset(i, 2) = 1 Then
        Range("SQ112!D2").Value = .Offset(i, 0)
          SOQUY112
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If
' CHI TIET
If Range("TTDN!F7") = 1 Then
With Range("HTTK!A1")
    For i = 590 To 892
    If .Offset(i, 2) = 1 Then
        Range("SCT_TK!C11").Value = .Offset(i, 0)
          KCSCT
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If
' CHI PHI
If Range("TTDN!F8") = 1 Then
With Range("HTTK!A1")
    For i = 896 To 926
    If .Offset(i, 2) = 1 Then
        Range("CP!C6").Value = .Offset(i, 0)
          CPSXKD
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If
'CONG NO
' PHAI THU
If Range("TTDN!I2") = 1 Then
    KC_CDSPS131
   Sheets("131TH").Select
     ActiveSheet.PrintOut
End If
If Range("TTDN!I3") = 1 Then
Sheets("SCT_CN").Select
Range("L9").Select
ActiveCell.Value = 131
With Range("131TH!A1")
    For i = 11 To 199
    If .Offset(i, 8) = 1 Then
        Range("SCTcn_MaKH").Value = .Offset(i, 0)
          SCTCN
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If
' PHAI TRA
If Range("TTDN!I4") = 1 Then
    KC_CDSPS331
   Sheets("331TH").Select
     ActiveSheet.PrintOut
End If
If Range("TTDN!I5") = 1 Then
Sheets("SCT_CN").Select
Range("L9").Select
ActiveCell.Value = 331
With Range("331TH!A1")
    For i = 11 To 199
    If .Offset(i, 8) = 1 Then
        Range("SCTcn_MaKH").Value = .Offset(i, 0)
          SCTCN
        ActiveSheet.PrintOut
    End If
    Next i
    End With
End If

' Khau hao-phan bo-NXT- Nhat ky - Bang luong
If Range("TTDN!F9") = 1 Then
   Sheets("KH").Select
     ActiveSheet.PrintOut
End If
If Range("TTDN!F10") = 1 Then
   Sheets("PB242").Select
     ActiveSheet.PrintOut
End If
If Range("TTDN!F11") = 1 Then
   Sheets("NXT152").Select
    ActiveSheet.PrintOut
End If
If Range("TTDN!F12") = 1 Then
   Sheets("NXT155").Select
    ActiveSheet.PrintOut
End If
If Range("TTDN!F13") = 1 Then
   Sheets("NXT156").Select
    ActiveSheet.PrintOut
End If
If Range("TTDN!F14") = 1 Then
   Sheets("NKban").Select
    ActiveSheet.PrintOut
End If
If Range("TTDN!F15") = 1 Then
   Sheets("NKmua").Select
     ActiveSheet.PrintOut
End If
If Range("TTDN!F16") = 1 Then
   Sheets("BL").Select
     ActiveSheet.PrintOut
End If

If Range("TTDN!F17") = 1 Then
   Sheets("C.Cong").Select
     ActiveSheet.PrintOut
End If

End Sub


Sub KiemTraInSo()
'

HOAN_THIEN

'KH:Sort va loc dong trong
    
    Sheets("KH").Select
    Range("BC11").Select
    ActiveCell.FormulaR1C1 = "Oxoa"
    Range("BC12").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("BC13").Select
    ActiveCell.FormulaR1C1 = "=IF(RC4>0,1,0)"
    Range("BC13").Select
    Selection.Copy
    Application.Goto Reference:="KH_Vfilter"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("KH_cellfilter1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("KH_cellfilter2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("KH_cellfilter3").Select
    ActiveCell.FormulaR1C1 = "1"
     Range("KH_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("BC11:BC12"), Unique:=False
    Application.Goto Reference:="KH_Vfilter1"
    Selection.ClearContents
    Application.CutCopyMode = False
    Range("H8").Select

'PB242:Sort va loc dong trong
    Sheets("PB242").Select
    Range("BC7").Select
    ActiveCell.FormulaR1C1 = "Oxoa"
    Range("BC8").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("BC9").Select
    ActiveCell.FormulaR1C1 = "=IF(RC3>0,1,0)"
    Range("BC9").Select
    Selection.Copy
    Application.Goto Reference:="PB242_Vfilter"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("PB242_Cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("PB242_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("BC7:BC8"), Unique:=False
    Application.Goto Reference:="PB242_Vfilter1"
    Selection.ClearContents
    Application.CutCopyMode = False
    Range("C7").Select
     
' Chon du lieu can in
    Sheets("TTDN").Select
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("F2:F8").Select
    Selection.FillDown
    Range("F9").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(KH_Vsum)<>0,1,"""")"
    Range("F10").Select
    ActiveCell.FormulaR1C1 = "=+IF(PB242_Vsum<>0,1,"""")"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(NXT152_Vsum)<>0,1,"""")"
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(NXT155_Vsum)<>0,1,"""")"
    Range("F13").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(NXT156_Vsum)<>0,1,"""")"
    Range("F14").Select
    ActiveCell.FormulaR1C1 = "=+IF(NKban_Vsum<>0,1,"""")"
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "=+IF(NKmua_Vsum<>0,1,"""")"
'BANG LUONG
    Range("F16").Select
    ActiveCell.FormulaR1C1 = "=+IF(BL_Vsum<>0,""x"","""")"
If Range("TTDN!F16") = "x" Then
     Sheets("BL").Select
     ActiveSheet.PrintPreview
End If
'BANG CHAM CONG
    Sheets("TTDN").Select
    Range("F17").Select
    ActiveCell.FormulaR1C1 = "=+IF(CC_Vsum<>0,""x"","""")"
If Range("TTDN!F17") = "x" Then
     Sheets("C.Cong").Select
     ActiveSheet.PrintPreview
End If
    Sheets("TTDN").Select
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(TH131_Vsum)<>0,1,"""")"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(TH131_Vsum)<>0,1,"""")"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(TH331_Vsum)<>0,1,"""")"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "=+IF(SUM(TH331_Vsum)<>0,1,"""")"
    
    Range("F1:F18").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I1:I6").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("I10").Select
        
End Sub


Sub HOAN_THIEN()

'NXT152:Sort va loc dong trong
    
    Application.Goto Reference:="NXT_V152"
    Selection.Sort Key1:=Range("C12"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-8]<>"""",RC[-6]<>"""",RC[-4]<>""""),1,0)"
    Selection.Copy
    Range("NXT152_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NXT152_cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NXT152_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N10:N11"), Unique:=False
    Range("B12").Select
    Selection.EntireColumn.Hidden = True
    Range("E12").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("H8").Select

'NXT155:Sort va loc dong trong
    Application.Goto Reference:="NXT_V155"
    Selection.Sort Key1:=Range("C12"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-8]<>"""",RC[-6]<>"""",RC[-4]<>""""),1,0)"
    Selection.Copy
    Range("NXT155_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NXT155_cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NXT155_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N10:N11"), Unique:=False
    Range("B12").Select
    Selection.EntireColumn.Hidden = True
    Range("E12").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("H8").Select

    LocNXT156
    LocNKban
    LocNKmua

' BANG LUONG:
    Sheets("BL").Select
    Range("D12:F12").Select
    Selection.EntireColumn.Hidden = True
    Range("I12:N12").Select
    Selection.EntireColumn.Hidden = True
    Range("Q12:R12").Select
    Selection.EntireColumn.Hidden = True
    Range("X12:Z12").Select
    Selection.EntireColumn.Hidden = True
    Range("Ac12:AH12").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("H8").Select

' CONG - NO
    Sheets("131TH").Select
    Range("A11:L11").AutoFilter
    Range("A11:L11").AutoFilter
    Range("A12:B200").Select
    Selection.Copy
    Sheets("TTKH").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    KC_CDSPS131
    Range("J11").Select
    
    Sheets("331TH").Select
    Range("A11:L11").AutoFilter
    Range("A11:L11").AutoFilter
    Range("A12:B200").Select
    Selection.Copy
    Sheets("TTKH").Select
    Range("G4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E4").Select
    KC_CDSPS331
    Sheets("TTKH").Select
    Range("E4").Select
    Sheets("CDSPS").Select
    Range("J10").Select
End Sub

Sub LocNXT156()

'NXT156:Sort va loc dong trong
    Application.Goto Reference:="NXT_V156"
    Selection.Sort Key1:=Range("C12"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-8]<>"""",RC[-6]<>"""",RC[-4]<>""""),1,0)"
    Selection.Copy
    Range("NXT156_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NXT156_cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NXT156_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N10:N11"), Unique:=False
    Range("B12").Select
    Selection.EntireColumn.Hidden = True
    Range("E12").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("H8").Select
    
End Sub

Sub LocNKban()

'NKban:Sort va loc dong trong
    Application.Goto Reference:="NKban_data"
    Selection.Sort Key1:=Range("C11"), Order1:=xlAscending, Key2:=Range("B11" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC5<>"""",1,0)"
    Selection.Copy
    Range("NKban_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NKban_Cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NKban_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("M9:M10"), Unique:=False
    Selection.EntireColumn.Hidden = True
    Range("D11").Select
    Selection.EntireColumn.Hidden = True
    Range("E8").Select
    
End Sub

Sub LocNKmua()

Application.Goto Reference:="NKmua_data"
    Selection.Sort Key1:=Range("C11"), Order1:=xlAscending, Key2:=Range("B11" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC5<>"""",1,0)"
    Selection.Copy
    Range("NKmua_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NKmua_Cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NKmua_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("K9:K10"), Unique:=False
    Selection.EntireColumn.Hidden = True
    Range("D11").Select
    Selection.EntireColumn.Hidden = True
    Range("E8").Select
    
End Sub


