Attribute VB_Name = "Module6"
Sub CPSXKD()
'
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("E9").Activate
If Range("A65536") = 1 Then
    Sheets("CP").Select
    Range("A11:O11").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="CP_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    Selection.AutoFilter
    
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=+TEXT(CP_tk,""@"")"
    Range("ghi_SC").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF((RC11+RC12)<>0,(IF(OR(TEXT(CP_tk,""@"")=TEXT(627,""@""),LEN(RC9)<7),TEXT(LEFT(RC9,3),""@""),TEXT(LEFT(RC9,4),""@""))),0)"
    Selection.Copy
    Application.Goto Reference:="cot_v"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.Goto Reference:="NKC_cotghiSC"
    Range("NKC_cotghiSC").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("G2:G3"), Unique:=False
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") <> 0 Then
    Range("NKC_CPdata").Select
    Selection.Copy
    Sheets("CP").Select
    Range("A12").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    End If
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    NKC_daucot
    Range("M2:N6").Select
    Selection.ClearContents
    Application.Goto Reference:="cot_v"
    Selection.ClearContents
    Range("E10").Select
    
    Sheets("CP").Select
    Range("J20").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Selection.Copy
    Range("CP_cotTT").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Range("CP_tk1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,2,0)"
    Range("CP_tk2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,3,0)"
    Range("CP_tk3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,4,0)"
    Range("CP_tk4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,5,0)"
    Range("CP_tk5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,6,0)"
    Range("CP_tk6").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,7,0)"
    Range("CP_tk7").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,8,0)"
    Range("CP_tk8").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,9,0)"
    Range("CP_tk0").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(value(CP_tk),CP_tkdata,10,0)"
    Range("CP_Vtk").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("K20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk1,""@"")=text(RC[-5],""@""),RC[-1],0)"
    Range("L20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk2,""@"")=text(RC[-6],""@""),RC[-2],0)"
    Range("M20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk3,""@"")=text(RC[-7],""@""),RC[-3],0)"
    Range("N20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk4,""@"")=text(RC[-8],""@""),RC[-4],0)"
    Range("O20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk5,""@"")=text(RC[-9],""@""),RC[-5],0)"
    Range("P20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk6,""@"")=text(RC[-10],""@""),RC[-6],0)"
    Range("Q20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk7,""@"")=text(RC[-11],""@""),RC[-7],0)"
    Range("R20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk8,""@"")=text(RC[-12],""@""),RC[-8],0)"
    Range("S20").Select
    ActiveCell.FormulaR1C1 = "=IF(text(CP_tk0,""@"")=text(RC[-13],""@""),RC[-9],0)"
    Range("K20:S20").Select
    Selection.Copy
    Range("CP_Vps").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Range("CP_tps").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotTT)"
    Range("CP_tps1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps1)"
    Range("CP_tps2").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps2)"
    Range("CP_tps3").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps3)"
    Range("CP_tps4").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps4)"
    Range("CP_tps5").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps5)"
    Range("CP_tps6").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps6)"
    Range("CP_tps7").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps7)"
    Range("CP_tps8").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps8)"
    Range("CP_tps0").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,CP_cotps0)"
    Range("CP_Vtps").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    'Filter-DANH SO TRANG
    Sheets("CP").Select
    Range("T12").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC10<>0,1,0)"
    Range("U12").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC10<>0,R[-1]C+1,R[-1]C)"
    Range("T12:U12").Select
    Selection.Copy
    Application.Goto Reference:="CP_VfilterSTT"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("CP_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(CP_cotSTT)+5,CP_Vtrang,2,1)"
    Range("CP_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(CP_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(CP_sotrang2,""00"")"
    Range("CP_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.Goto Reference:="CP_cotfilter"
    Range("CP_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("T10:T11"), Unique:=False
    
    Range("D9").Activate
    Selection.EntireColumn.Hidden = True
     Range("F9").Activate
    Selection.EntireColumn.Hidden = True
    Range("H9:I9").Activate
    Selection.EntireColumn.Hidden = True
    Range("T9:U9").Activate
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("O8").Activate
  Exit Sub
 Else
  Sheets("CP").Select
  Range("E8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If
 
End Sub




