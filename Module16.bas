Attribute VB_Name = "Module16"
Sub cthh()
   
   
   If Range("nam") = 2018 Then
   
   Sheets("SCT156").Select
   Range("SCT156!A14:S14").AutoFilter
   Range("SCT156!10:3000").EntireRow.Hidden = False
   Range("SCT156!A:i").EntireColumn.Hidden = False
   Range("I2").Select
   Selection.ClearContents
   Range("Q11").Select
   Selection.ClearContents
   Range("SCT156!SCT156_data").ClearContents
   
   If Range("SCT_maHH") <> "" Then
   Sheets("N").Select
   Range("N!A11:P11").AutoFilter
   Range("N!A:i").EntireColumn.Hidden = False
   Sheets("X").Select
   Range("X!A11:P11").AutoFilter
   Range("X!A:i").EntireColumn.Hidden = False
   Sheets("NXT").Select
   Range("NXT!A11:T11").AutoFilter
   Range("NXT!A11:T11").AutoFilter
   Range("NXT!A:i").EntireColumn.Hidden = False
   
'chep
    
    Sheets("N").Select
    Range("O10").Select
    ActiveCell.FormulaR1C1 = "MaHH"
    Range("O11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("O12").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC4=SCT_maHH,1,0)"
    Range("N_VfilterMH1").FillDown
    Range("D4").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(COUNTIF(N_VfilterMH1,1)>0,1,0)"
    If Range("D4") = 1 Then
    Range("N_VfilterMH2").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("O10:O11"), Unique:=False
    Range("N!N_data").Copy
    Range("SCT156_cellN1").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    End If
    Sheets("X").Select
    Range("O10").Select
    ActiveCell.FormulaR1C1 = "MaHH"
    Range("O11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("O12").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC4=SCT_maHH,1,0)"
    Range("X_VfilterMH1").FillDown
    Range("D4").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(COUNTIF(X_VfilterMH1,1)>0,1,0)"
    If Range("D4") = 1 Then
    Range("X_VfilterMH2").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("O10:O11"), Unique:=False
    Range("X!X_data").Copy
    Range("SCT156_cellX1").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    End If
    
'PHUC HOI HIEN TRANG N-X-NXT
    Sheets("N").Select
    Range("N!A11:P11").AutoFilter
    Range("N!D:D").EntireColumn.Hidden = True
    Range("E8").Select
    Sheets("X").Select
    Range("X!A11:P11").AutoFilter
    Range("X!D:D").EntireColumn.Hidden = True
    Range("E8").Select
    Sheets("NXT").Select
    Range("NXT!A11:T11").AutoFilter
    Range("NXT!E:E").EntireColumn.Hidden = True
    
'lam phan nhap-xuat-ton SCT156
    Sheets("SCT156").Select
    Range("SCT156!SCT156_Vnhap").Copy
    Range("SCT156_cellN2").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("SCT156!SCT156_Vxuat").Copy
    Range("SCT156_cellX2").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.Goto Reference:="SCT156_cellT1"
    ActiveCell.FormulaR1C1 = "=+R[-1]C+RC[-4]-RC[-2]"
    Selection.Copy
    Application.Goto Reference:="SCT156_VtonHH"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
' Lam dien giai
    Range("SCT156_cellDG").Select
    ActiveCell.FormulaR1C1 = "=IF(RC4<>"""",VLOOKUP(RC4,NXT_Vmh,2,0),"""")"
    Selection.Copy
    Range("SCT156_Vdg").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'sort
    Range("SCT156_data").Sort Key1:=Range("A16"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    
    End If
'loc record
    Range("S15").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("R15").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("R16").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(RC[-8]:RC[-5])<>0,1,0)"
    Range("S16").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-5]<0,RC[-4]<0),1,0)"
    Range("SCT156_Vfilter").FillDown
    Range("SCT156_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
    Range("S14:S15"), Unique:=False
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(R16C19:R3015C19)=0,"""",""AM HANG-AM HANG-AM HANG"")"
'Kiem tra am
    Sheets("NXT").Select
    Range("Q12").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC2=SCT_maHH,SCT156_cellAH,RC16)"
    Range("NXT_Vamhang").Select
    Selection.FillDown
    Selection.Copy
    Range("P12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("NXT_Vamhang").ClearContents
    Range("K8").Select
' an cot
    Sheets("SCT156").Select
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:S").Select
    Selection.EntireColumn.Hidden = True
    
    Sheets("SCT156").Select
    Application.CutCopyMode = False
    Range("A10").Select
    Exit Sub
  Else
    Sheets("SCT156").Select
    Range("A10").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If
  
End Sub

Sub TimDL1()
'
    Range("I2").Select
    Selection.Copy
    Range("SCT_MaHH").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I2").Select
    Selection.ClearContents
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(SCT_MaHH,NXT_data,2,0))=TRUE,""."",VLOOKUP(SCT_MaHH,NXT_data,2,0))"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(SCT_MaHH,NXT_data,3,0))=TRUE,"""",VLOOKUP(SCT_MaHH,NXT_data,3,0))"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(SCT_MaHH,NXT_data,5,0))=TRUE,0,VLOOKUP(SCT_MaHH,NXT_data,5,0))"
    Range("L3").Select
End Sub

Sub LayDL_N()

    Range("D4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("D4").Copy
    Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("D4") = 1 Then
    
    R = MsgBox("Ban CO CHAC la MUON THUC HIEN LENH NAY ko?", vbYesNo, "NGUY HIEM")
    If R = vbNo Then
    Range("E8").Activate
    Exit Sub
    End If
    
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "=CELL(""row"",N_cellsum)-CELL(""row"",NKmua_Vsum)-2"
    Selection.Copy
    Range("G8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("G8") < 0 Then
 
 MsgBox ("Sheet ""N"" KHONG DU DONG")
    
Else
    Application.Goto Reference:="N_dataN"
    Selection.ClearContents
    Range("J12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC2,NK1!R2C2:R20000C11,4,0))=TRUE,""-"",VLOOKUP(RC2,NK1!R2C2:R20000C11,4,0))"
    Range("N_data5").FillDown
    Application.Goto Reference:="NKmua_data"
    Selection.Sort Key1:=Range("C11"), Order1:=xlAscending, Key2:=Range("B11" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    Selection.EntireColumn.Hidden = False
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC4<>"""",1,0)"
    Selection.Copy
    Range("NKmua_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NKmua_Cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NKmua_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("K9:K10"), Unique:=False
    Application.Goto Reference:="NKmua_dataSCT"
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("N").Select
    Range("B12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LocNKmua
End If
End If
    Sheets("N").Select
    Range("E8").Select
End Sub


Sub LayDL_X()

    Range("D4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("D4").Copy
    Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("D4") = 1 Then
    
    R = MsgBox("Ban CO CHAC la MUON THUC HIEN LENH NAY ko?", vbYesNo, "NGUY HIEM")
    If R = vbNo Then
    Range("E8").Activate
    Exit Sub
    End If
    
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "=+CELL(""row"",X_cellsum)-CELL(""row"",NKban_Vsum)-2"
    Selection.Copy
    Range("G8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("G8") < 0 Then
 
 MsgBox ("Sheet ""X"" KHONG DU DONG")
    
Else
    Application.Goto Reference:="X_dataX"
    Selection.ClearContents
    Range("J12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISNA(VLOOKUP(RC2,NK1!R2C2:R20000C11,4,0))=TRUE,""-"",VLOOKUP(RC2,NK1!R2C2:R20000C11,4,0))"
    Range("X_data5").FillDown
    
    Application.Goto Reference:="NKban_data"
    Selection.Sort Key1:=Range("C11"), Order1:=xlAscending, Key2:=Range("B11" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    Selection.EntireColumn.Hidden = False
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC4<>"""",1,0)"
    Selection.Copy
    Range("NKban_Vfilter").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("NKban_Cellfilter").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("NKban_Vfilter1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("M9:M10"), Unique:=False
    Application.Goto Reference:="NKban_dataSCT"
    Selection.Copy
    Sheets("X").Select
    Range("B12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LocNKban
End If
End If
    Sheets("X").Select
    Range("E8").Select
End Sub

Sub LayDL_NXT()

    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("B4").Copy
    Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("B4") = 1 Then
 Range("J8").Select
    ActiveCell.FormulaR1C1 = "=+CELL(""row"",NXT_cellsum)-CELL(""row"",NXT156_Vsum)-2"
    Selection.Copy
    Range("J8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
If Range("J8") < 0 Then
 
 MsgBox ("Sheet ""NXT"" KHONG DU DONG")
    
Else

    Sheets("NXT").Select
    Range("NXT!A11:P11").AutoFilter
    Range("NXT!10:3000").EntireRow.Hidden = False
    Application.Goto Reference:="NXT_data"
    Selection.ClearContents
    Application.Goto Reference:="NXT_Vamhang2"
    Selection.ClearContents
    Sheets("NXT156").Select
    Range("NXT156!A11:P11").AutoFilter
    Application.Goto Reference:="NXT_V156"
    Selection.Sort Key1:=Range("C11"), Order1:=xlAscending, Key2:=Range("B11" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    Selection.EntireColumn.Hidden = False
    Selection.Copy
    Sheets("NXT").Select
    Range("B12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LocNXT156
End If
    
End If
    Sheets("NXT").Select
    Range("K8").Select
End Sub

Sub Amhang()

'Kiem tra truoc khi in
    Sheets("SCT156").Select
   If Range("nam") = 2018 Then

R = MsgBox("Kiem tra am hang a?", vbYesNo, "Coi chung")
If R = vbNo Then
Exit Sub
End If
'Neu kiem tra ky roi thi in
    Sheets("NXT").Select
    Range("O12").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(RC8:RC11)>0,1,0)"
    Selection.Copy
    Application.Goto Reference:="NXT_DSinSCT"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("SCT156").Select
    Range("I2").Select
    Selection.ClearContents
    Range("A10").Select
    
Dim i
With Range("NXT!B1")
    For i = 11 To 1498
    If .Offset(i, 13) = 1 Then
        Range("SCT156!SCT_maHH").Value = .Offset(i, 0)
        Sheets("SCT156").Select
        cthh
        Range("SCT156!Q11").Value = .Offset(i, -1)
        'ActiveSheet.PrintPreview
    End If
    Next i
    End With

Exit Sub
  Else
  Sheets("SCT156").Select
  Range("A10").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If

End Sub

Sub ISCT()
   Sheets("SCT156").Select
   If Range("nam") = 2018 Then

'Kiem tra truoc khi in
R = MsgBox("Kiem tra ky chua?", vbYesNo, "Coi chung")
If R = vbNo Then
Exit Sub
End If
'Neu kiem tra ky roi thi in
Sheets("NXT").Select
    Range("O12").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(RC8:RC11)>0,1,0)"
    Selection.Copy
    Application.Goto Reference:="NXT_DSinSCT"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("SCT156").Select
    Range("I2").Select
    Selection.ClearContents
    Range("A10").Select
Dim i

With Range("NXT!B1")
    For i = 11 To 1498
    If .Offset(i, 13) = 1 Then
        Range("SCT156!SCT_maHH").Value = .Offset(i, 0)
        Sheets("SCT156").Select
        cthh
        Range("SCT156!Q11").Value = .Offset(i, -1)
        ActiveSheet.PrintOut
    End If
    Next i
    End With

Exit Sub
  Else
Sheets("SCT156").Select
  Range("A10").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If

End Sub



