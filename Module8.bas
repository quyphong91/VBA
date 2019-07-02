Attribute VB_Name = "Module8"
Sub KC_CDSPS131()
'
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then
    Sheets("SCT_CN").Select
    Range("A17:J17").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="SCTcn_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    Selection.AutoFilter
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "131"
    Range("NKC_SCTcnfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N1:N2"), Unique:=False
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") <> 0 Then
    Application.Goto Reference:="NKC_SCTcndata1"
    Selection.Copy
    Sheets("SCT_CN").Select
    Range("A18").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.Goto Reference:="NKC_cotTT"
    Selection.Copy
    Range("H18").Select
    Sheets("SCT_CN").Select
    Range("H18").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    End If
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    NKC_daucot
    Range("M2:N6").Select
    Selection.ClearContents
    Range("E10").Select
    
    SCTCN_phan2
    
    Sheets("131TH").Select
    Range("A11:i11").Select
    Selection.AutoFilter
        
' Lay du lieu
    Sheets("131TH").Select
    Range("E20").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SCTcn_cotmaKH,RC1,SCTcn_cotpsno)"
    Range("F20").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SCTcn_cotmaKH,RC1,SCTcn_cotpsco)"
    Range("G20").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC3+RC5-RC4-RC6,0)"
    Range("H20").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC4+RC6-RC3-RC5,0)"
    Range("I20").Select
    ActiveCell.FormulaR1C1 = "=IF((RC[-3]+RC[-4])<>0,1,0)"
    Range("j20").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-7]<>0,RC[-6]<>0,RC[-5]<>0,,RC[-4]<>0),1,0)"
    Range("E20:J20").Select
    Selection.Copy
    Range("CD_131").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("CD_131sps").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("tgddn_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg1_131)"
    Range("tgddc_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg1.2_131)"
    Range("tgpsn_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg2_131)"
    Range("tgpsc_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg3_131)"
    Range("tgdcn_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg4.2_131)"
    Range("tgdcc_131").Select
    ActiveCell.FormulaR1C1 = "=SUM(vg4_131)"
    Range("A11:J11").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=10, Criteria1:="1"
    'Sheets("CDSPS").Select
    'Range("CDSPS_lock1").Select
    'ActiveCell.FormulaR1C1 = _
        "=+IF(or(MID(CELL(""filename""),IF(ISERROR(FIND(""1TS-KH"",CELL(""filename"")))=TRUE,1,FIND(""1TS-KH"",CELL(""filename""))),6)=""1TS-KH"",MID(CELL(""filename""),IF(ISERROR(FIND(""TS-K"",CELL(""filename"")))=TRUE,1,FIND(""TS-K"",CELL(""filename""))),4)=""TS-K""),0,nam)"
    'Range("CDSPS_lock1").Copy
    'Range("CDSPS_lock1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Application.CutCopyMode = False
    Sheets("131TH").Select
    Range("K9:K9").Activate
    Selection.EntireColumn.Hidden = True
    Range("K11").Select
    
  Exit Sub
 Else
  Sheets("131TH").Select
  Range("E8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub

