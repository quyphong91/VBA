Attribute VB_Name = "Module14"
Sub DL_THUCHI()
'
' Macro1 Macro
' Macro recorded 26/07/2007 by PC01
    
    
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

If Range("A65536") = 1 Then
    
    Range("Q3:Z2000").Select
    Selection.ClearContents
    ' Lam334
    If Range("NKC_PLno") <> 0 Then
    Sheets("THU_CHI").Select
    Range("Q2001").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(thang,Date,3,0)"
    Range("R2001").Select
    ActiveCell.FormulaR1C1 = "=""BL""&TEXT(thang,""00"")"
    Range("S2001").Select
    ActiveCell.FormulaR1C1 = "=+""PC""&TEXT(thang,""00"")&""-""&PCL"
    Range("T2001").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(thang,Date,3,0)"
    Range("U2001").Select
    ActiveCell.FormulaR1C1 = "=TTDN!R1C3"
    Sheets("NKC").Select
    Range("NKC_PLdiengiai").Select
    Selection.Copy
    Sheets("THU_CHI").Select
    Range("W2001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("NKC").Select
    Range("I:K").Select
    Selection.EntireColumn.Hidden = False
    Range("NKC_dong334").Select
    Selection.Copy
    Range("D9").Activate
    Selection.EntireColumn.Hidden = True
    Range("J9").Activate
    Selection.EntireColumn.Hidden = True
    Range("E9").Activate
    Sheets("THU_CHI").Select
    Range("X2001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   Range("Q2001:Z2001").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   End If
' COPY THU - CHI:
   Range("V2001").Select
   ActiveCell.FormulaR1C1 = "=COUNTA(NK1!R3C2:R2000C2)"
   If Range("V2001") <> 0 Then
    XuLyTHU_CHI
    End If
    Range("Q3:Z2001").Select
    Selection.Sort Key1:=Range("S3"), Order1:=xlAscending, Key2:=Range("T3") _
        , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:= _
        xlSortNormal
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=+COUNTA(R3C19:R2000C19)+2"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "=VNDuni(R12C4)"
    Range("D23").Select
    ActiveCell.FormulaR1C1 = "=VNDuni(R12C4)"
    Sheets("NK1").Select
    Selection.AutoFilter
    Range("O1:S2").Select
    Selection.FillDown
    Selection.ClearContents
    Range("M3:Y2000").Select
    Selection.ClearContents
    Range("A3").Select
    Sheets("THU_CHI").Select
    Range("K6").Select
Else
  Sheets("THU_CHI").Select
  Range("K8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 
 End If
End Sub


Sub XuLyTHU_CHI()
'
    Sheets("NK1").Select
    Range("T2").Select
    Range("C2:C2000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "N1:N2"), CopyToRange:=Range("T2"), Unique:=True
    
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=+RC1"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=+RC2"
    
    Range("O3").Select
     ActiveCell.FormulaR1C1 = _
        "=+IF(OR(RC10=1111,RC11=1111),(IF(AND(RC2=R[1]C2,RC4=R[1]C4,RC5=R[1]C5),RC7&""/""&R[1]C7&""É"",RC7)),RC7)"
    Range("P3").Select
     ActiveCell.FormulaR1C1 = _
        "=+IF(OR(RC11=1111),(IF(AND(RC2=R[1]C2,RC4=R[1]C4,RC5=R[1]C5,LEFT(RC3,2)=""PC""),RC10&""/""&R[1]C10&""É"",RC10)),RC10)"
    Range("Q3").Select
     ActiveCell.FormulaR1C1 = _
        "=+IF(OR(RC10=1111),(IF(AND(RC2=R[1]C2,RC4=R[1]C4,RC5=R[1]C5,LEFT(RC3,2)=""PT""),LEFT(RC11,4)&""/""&LEFT(R[1]C11,4)&""/133É"",RC11)),RC11)"
    
    Range("R3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,11,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,11,0))"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,12,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,12,0))"
    Range("M3:S2000").Select
    Selection.FillDown
    Range("U3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,2,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,2,0))"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,3,0))=TRUE,"""",(IF(VLOOKUP(RC20,R3C3:R2000C17,3,0)="""",congty,VLOOKUP(RC20,R3C3:R2000C17,3,0))))"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,4,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,4,0))"
    Range("X3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,13,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,13,0))"
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,14,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,14,0))"
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC20,R3C3:R2000C17,15,0))=TRUE,"""",VLOOKUP(RC20,R3C3:R2000C17,15,0))"
    Range("AA3").Select
    ActiveCell.FormulaR1C1 = "=+SUMIF(R2C3:R2000C3,RC20,R2C12:R2000C12)"
    Range("U3:AA2000").Select
    Selection.FillDown
    Range("M3:AA2000").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "DATE"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = ">0"
    Range("R1:R2000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
       Range("R1:R2"), Unique:=False
    Range("R3:AA2000").Select
    Selection.Copy
    Sheets("THU_CHI").Select
    Range("Q3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

