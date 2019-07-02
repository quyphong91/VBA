Attribute VB_Name = "Module2"
Sub SOCAI()
'

' Bo Bao Ve
    'S203.Calculate
    'ActiveSheet.Protect ("trithuc"), DrawingObjects:=False, Contents:=True, Scenarios:= _
            True, AllowFiltering:=True
    'ActiveSheet.Unprotect ("trithuc")
    'Cells.Locked = True

    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E1").Select
If Range("A65536") = 1 Then
    ' LOC DU LIEU COPY
    Sheets("SC").Select
    Range("A21:J21").Activate
    Selection.EntireColumn.Hidden = False
    Selection.AutoFilter
    Application.Goto Reference:="SC_nd"
    Selection.ClearContents
    Sheets("NKC").Select
    Range("A12:L12").Activate
    Selection.EntireColumn.Hidden = False
    Range("D_locnk").Activate
    Selection.AutoFilter
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=SC_tk"
SOCAI_ChonTK
    
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC14<>"""",RC14,""x"")"
    Range("O2").Copy
    Range("O2:O60").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("NKC_cotTK").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N1:N60"), Unique:=False
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,NKC_cotTT)"
    If Range("G1") > 0 Then
    Range("NKC_SCdata").Select
    Selection.Copy
    Sheets("SC").Select
    Range("A21").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    End If
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    NKC_daucot
    Range("M2:O60").Select
    Selection.ClearContents
    Range("E10").Select
    
    Sheets("SC").Select
    Range("SC_ddno").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SC_tk,vtg1)"
    Range("SC_ddco").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(cd_shtk,SC_tk,vtg2)"
    Range("SC_psno").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(cd_shtk,SC_tk,vtg3)"
    Range("SC_psco").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(cd_shtk,SC_tk,vtg4)"
    Range("SC_dcno").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(cd_shtk,SC_tk,vtg5)"
    Range("SC_dcco").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(cd_shtk,SC_tk,vtg6)"
    Range("SC_Vtong").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
      
    'Filter-DANH SO TRANG
    Sheets("SC").Select
    Range("K21").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-2]+RC[-1])>0,1,0)"
    Range("L21").Select
    ActiveCell.FormulaR1C1 = "=+IF((RC[-3]+RC[-2])>0,R[-1]C+1,R[-1]C)"
    Range("SC_VfilterSTT").Select
    Selection.FillDown
    Range("SC_sotrang2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(MAX(SC_cotSTT)+4,SC_Vtrang,2,1)"
    Range("SC_sotrang1").Select
    ActiveCell.FormulaR1C1 = _
        "=+LEFT(NKC_celltongtrang,10)&TEXT(SC_sotrang2,""00"")&MID(NKC_celltongtrang,13,26)&TEXT(SC_sotrang2,""00"")"
    Range("SC_sotrang1").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.Goto Reference:="SC_cotfilter"
    Range("SC_cotfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("K19:K20"), Unique:=False
      
    Rows("1:7").Select
    Selection.EntireRow.Hidden = True
    Range("D9").Activate
    Selection.EntireColumn.Hidden = True
    Range("K9:L9").Activate
    Selection.EntireColumn.Hidden = True
    
    Application.CutCopyMode = False
    Range("E14").Select
  Exit Sub
 Else
  Sheets("SC").Select
  Range("E8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If
' Bao ve
    'With Range("a1:M1163").Select
     '   .Locked = False
    'End With
    'ActiveSheet.Protect ("trithuc")
    'ActiveSheet.Protect ("trithuc"), DrawingObjects:=True, Contents:=True, Scenarios:=True

End Sub

Sub SOCAI_ChonTK()

If Range("SC_tk") = 111 Then
    Range("TK_V111").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If

If Range("SC_tk") = 112 Then
    Range("TK_V112").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 121 Then
    Range("TK_V121").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 128 Then
    Range("TK_V128").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 133 Then
    Range("TK_V133").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 136 Then
    Range("TK_V136").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 138 Then
    Range("TK_V138").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 141 Then
    Range("TK_V141").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 144 Then
    Range("TK_V144").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 152 Then
    Range("TK_V152").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 154 Then
    Range("TK_V154").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 155 Then
    Range("TK_V155").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 156 Then
    Range("TK_V156").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 211 Then
    Range("TK_V211").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 212 Then
    Range("TK_V212").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 213 Then
    Range("TK_V213").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 214 Then
    Range("TK_V214").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 228 Then
    Range("TK_V228").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 229 Then
    Range("TK_V229").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 241 Then
    Range("TK_V241").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 242 Then
    Range("TK_V242").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 244 Then
    Range("TK_V244").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 333 Then
    Range("TK_V333").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 334 Then
    Range("TK_V334").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 336 Then
    Range("TK_V336").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 338 Then
    Range("TK_V338").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 341 Then
    Range("TK_V341").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 411 Then
    Range("TK_V411").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 413 Then
    Range("TK_V413").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 421 Then
    Range("TK_V421").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 511 Then
    Range("TK_V511").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 515 Then
    Range("TK_V515").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 521 Then
    Range("TK_V521").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If

If Range("SC_tk") = 611 Then
    Range("TK_V611").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 621 Then
    Range("TK_V621").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 622 Then
    Range("TK_V622").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 623 Then
    Range("TK_V623").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 627 Then
    Range("TK_V627").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 631 Then
    Range("TK_V631").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 632 Then
    Range("TK_V632").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 635 Then
    Range("TK_V635").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 641 Then
    Range("TK_V641").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 642 Then
    Range("TK_V642").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If
If Range("SC_tk") = 821 Then
    Range("TK_V821").Copy
    Range("N3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End If


End Sub

