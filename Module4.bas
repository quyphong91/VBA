Attribute VB_Name = "Module4"
Sub SOQUY112()
'
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then
    Sheets("SQ112").Select
    Range("G10").Select
    Range("A15:J15").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="SQ112_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=SQ112_tk"
    Range("NKC_cotTK").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N1:N2"), Unique:=False
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") <> 0 Then
    Range("NKC_SQ112data").Select
    Selection.Copy
    Sheets("SQ112").Select
    Range("A16").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
' SORT LAI
    Range("J16").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]<>0,1,2)"
    Range("SQ112_VfilterSTT").Select
    Selection.FillDown
    Application.Goto Reference:="SQ112_sort"
    Selection.Sort Key1:=Range("C16"), Order1:=xlAscending, Key2:=Range("J16" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
' Lam Dky - Phat sinh - Ton
    Sheets("SQ112").Select
    Range("SQ112_dk").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SQ112_tk,vtg1)"
    Range("I16").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+RC[-2]-RC[-1]"
    Selection.Copy
    Range("SQ112_Vton").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("SQ112_PSthu").Select
    ActiveCell.FormulaR1C1 = "=SUM(SQ112_Vthu)"
    Range("SQ112_PSchi").Select
    ActiveCell.FormulaR1C1 = "=SUM(SQ112_Vchi)"
    Range("SQ112_ck").Select
    ActiveCell.FormulaR1C1 = "=IF((SQ112_PSthu+SQ112_PSchi)<>0,R[-3]C,SQ112_dk)"
    End If
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    NKC_daucot
    Range("M2:N6").Select
    Selection.ClearContents
    Range("E10").Select
    
'   Kiem tra am-DANH SO TRANG-Filter
    Sheets("SQ112").Select
    Range("J16").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-3]+RC[-2])>0,1,0)+IF(RC[-1]<0,2,0)"
    Range("K16").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-4]+RC[-3])>0,R[-1]C+1,R[-1]C)"
    Range("SQ112_VfilterSTT").Select
    Selection.FillDown
    Range("SQ112_am").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(SQ112_cotfilter,2)"
    Range("SQ112_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(SQ112_cotSTT)+6,SQ112_Vtrang,2,1)"
    Range("SQ112_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(SQ112_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(SQ112_sotrang2,""00"")"
    Range("SQ112_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.Goto Reference:="SQ112_cotfilter"
    Range("SQ112_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("J10:J12"), Unique:=False
     
  If Range("SQ112_am") <> 0 Then
            R = MsgBox("AM QUY - AM QUY - AM QUY. Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
       If R = vbYes Then
         Exit Sub
      End If
      End If
    
    Rows("1:3").Select
    Selection.EntireRow.Hidden = True
    Range("D16").Select
    Selection.EntireColumn.Hidden = True
    Range("J16:K16").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("E12").Select
  Exit Sub
 Else
 Sheets("SQ112").Select
  Range("E12").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub



