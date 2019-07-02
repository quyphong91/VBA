Attribute VB_Name = "Module7"
Sub SCTCN()
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
    
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=SCTcn_maKH"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=+SCTcn_loaiCN"
    Range("NKC_SCTcnfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("M1:N2"), Unique:=False
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

  Exit Sub
 Else
  Sheets("SCT_CN").Select
  Range("E16").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub

Sub SCTCN_phan2()

 ' LAM SO DU DKY - PHAT SINH - CUOI KY
    Sheets("SCT_CN").Select
    Range("SCTcn_ddno").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SCTcn_loaiCN=131,SUMIF(MaKH_131,SCTcn_maKH,vg1_131),SUMIF(MaKH_331,SCTcn_maKH,vg1_331))"
    Range("SCTcn_ddco").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SCTcn_loaiCN=131,SUMIF(MaKH_131,SCTcn_maKH,vg1.2_131),SUMIF(MaKH_331,SCTcn_maKH,vg1.2_331))"
    Range("J18").Select
    ActiveCell.FormulaR1C1 = "=+MAX(R[-1]C-R[-1]C[1]+RC[-2]-RC[-1],0)"
    Range("K18").Select
    ActiveCell.FormulaR1C1 = "=+MAX(R[-1]C-R[-1]C[-1]+RC[-2]-RC[-3],0)"
    Range("J18:K18").Select
    Selection.Copy
    Range("SCTcn_Vton").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("SCTcn_psno").Select
    ActiveCell.FormulaR1C1 = "=SUM(SCTcn_cotpsno)"
    Range("SCTcn_psco").Select
    ActiveCell.FormulaR1C1 = "=SUM(SCTcn_cotpsco)"
    Range("SCTcn_dcno").Select
    ActiveCell.FormulaR1C1 = "=R[-3]C"
    Range("SCTcn_dcco").Select
    ActiveCell.FormulaR1C1 = "=R[-3]C"
    
    ' FILTER-DANH SO TRANG
    Range("L18").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-4]+RC[-3])<>0,1,0)"
    Range("M18").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-5]+RC[-4])<>0,R[-1]C+1,R[-1]C)"
    Range("SCTcn_VfilterSTT").Select
    Selection.FillDown
    Range("SCTcn_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(SCTcn_cotSTT)+6,SCTcn_Vtrang,2,1)"
    Range("SCTcn_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(SCT_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(SCT_sotrang2,""00"")"
    Range("SCTcn_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
    Application.Goto Reference:="SCTcn_cotfilter"
    Range("SCTcn_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("L16:L17"), Unique:=False
    Rows("1:3").Select
    Selection.EntireRow.Hidden = True
    Range("D9").Activate
    Selection.EntireColumn.Hidden = True
    Range("L9:M9").Activate
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("E14").Select


End Sub



