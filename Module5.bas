Attribute VB_Name = "Module5"
Sub KCSCT()
'
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then

    Sheets("SCT_tk").Select
    Range("A17:J17").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="SCT_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    Selection.AutoFilter
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=SCT_tk"
    Range("NKC_cotTK").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N1:N2"), Unique:=False
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") <> 0 Then
    Range("NKC_SQ112data").Select
    Selection.Copy
    Sheets("SCT_tk").Select
    Range("A18").Select
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
   
 ' LAM SO DU DKY - PHAT SINH-CUOI KY
    Sheets("SCT_tk").Select
    Range("SCT_ddno").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg1)"
    Range("SCT_ddco").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg2)"
    Range("I18").Select
    ActiveCell.FormulaR1C1 = "=+MAX(R[-1]C-R[-1]C[1]+RC[-2]-RC[-1],0)"
    Range("J18").Select
    ActiveCell.FormulaR1C1 = "=+MAX(R[-1]C-R[-1]C[-1]+RC[-2]-RC[-3],0)"
    Range("I18:J18").Select
    Selection.Copy
    Range("SCT_Vton").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("SCT_PSno").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg3)"
    Range("SCT_PSco").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg4)"
    Range("SCT_dcno").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg5)"
    Range("SCT_dcco").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SCT_tk,vtg6)"
    
    ' FILTER-DANH SO TRANG
    Range("K18").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-4]+RC[-3])<>0,1,0)"
    Range("L18").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-5]+RC[-4])<>0,R[-1]C+1,R[-1]C)"
    Range("SCT_VfilterSTT").Select
    Selection.FillDown
    Range("SCT_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(SCT_cotSTT)+6,SCT_Vtrang,2,1)"
    Range("SCT_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(SCT_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(SCT_sotrang2,""00"")"
    Range("SCT_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.Goto Reference:="SCT_cotfilter"
    Range("SCT_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("K16:K17"), Unique:=False
    
    Rows("1:3").Select
    Selection.EntireRow.Hidden = True
    Range("D9").Activate
    Selection.EntireColumn.Hidden = True
    Range("K9:L9").Activate
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("E14").Select

  Exit Sub
 Else
  Sheets("SCT_tk").Select
  Range("E14").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub



