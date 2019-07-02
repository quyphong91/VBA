Attribute VB_Name = "Module3"
Sub SOQUY111()
'
    Sheets("SQ111").Select
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then

    SQ_XuLy
    
    If Range("SQ_am") > 0 Then
         R = MsgBox("AM QUY - AM QUY - AM QUY.")
      End If
    
    
  Exit Sub
 Else
 Sheets("SQ111").Select
  Range("G12").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 End If

End Sub

Sub SQ_XuLy()

    Sheets("SQ111").Select
    Range("G10").Select
    Range("A11:N11").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="SQ_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    Selection.AutoFilter
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=SQ_tk"
    Range("NKC_cotTK").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N1:N2"), Unique:=False
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") <> 0 Then
    Range("NKC_SQdata1").Select
    Selection.Copy
    Sheets("SQ111").Select
    Range("A16").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("NKC").Select
    Range("NKC_SQdata2").Select
    Selection.Copy
    Sheets("SQ111").Select
    Range("F16").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
' SORT LAI
    Range("L16").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]<>0,1,2)"
    Range("SQ_VfilterSTT").Select
    Selection.FillDown
    Application.Goto Reference:="SQ_sort"
    Selection.Sort Key1:=Range("C16"), Order1:=xlAscending, Key2:=Range("L16" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

' Lam phieu thu chi
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "=IF(RC9=0,"""",RC2)"
    Range("E19").Select
    ActiveCell.FormulaR1C1 = "=IF(RC10=0,"""",RC2)"
    Range("D19:E19").Select
    Selection.Copy
    Range("SQ_Vthuchi").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

  ' Lam Dky - Phat sinh - Ton
    Sheets("SQ111").Select
    Range("SQ_dk").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SQ_tk,vtg1)"
    Range("K16").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+RC[-2]-RC[-1]"
    Selection.Copy
    Range("SQ_Vton").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("SQ_PSthu").Select
    ActiveCell.FormulaR1C1 = "=SUM(SQ_Vthu)"
    Range("SQ_PSchi").Select
    ActiveCell.FormulaR1C1 = "=SUM(SQ_Vchi)"
    Range("SQ_ck").Select
    ActiveCell.FormulaR1C1 = "=IF((SQ_PSthu+SQ_PSchi)<>0,R[-3]C,SQ_dk)"
    End If
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    NKC_daucot
    Range("M2:N6").Select
    Selection.ClearContents
    Range("E10").Select
    
'   Filter-Kiem tra am-DANH SO TRANG
    Sheets("SQ111").Select
    Range("L16").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-3]+RC[-2])>0,1,0)+IF(RC[-1]<0,2,0)"
    Range("M16").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-4]+RC[-3])>0,R[-1]C+1,R[-1]C)"
    Range("SQ_VfilterSTT").Select
    Selection.FillDown
    Range("SQ_am").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(SQ_cotfilter,3)"
    Range("SQ_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(SQ_cotSTT)+6,SQ_Vtrang,2,1)"
    Range("SQ_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(SQ_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(SQ_sotrang2,""00"")"
    Range("SQ_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.Goto Reference:="SQ_cotfilter"
    Range("SQ_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("L10:L12"), Unique:=False
    
        
    Rows("1:3").Select
    Selection.EntireRow.Hidden = True
    Range("B9:B9").Activate
    Selection.EntireColumn.Hidden = True
    Range("F9:F9").Activate
    Selection.EntireColumn.Hidden = True
    Range("L9:M9").Activate
    Selection.EntireColumn.Hidden = True
    
    Application.CutCopyMode = False
    Range("G12").Select
End Sub

