Attribute VB_Name = "Module10"
Sub Begin()
'
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E9").Activate
If Range("A65536") = 1 Then
    Sheets("CDSPS").Select
    Range("B8").Activate
    R = MsgBox("Ban CO CHAC la MUON THUC HIEN LENH NAY ko?", vbYesNo, "NGUY HIEM")
    If R = vbNo Then
    Sheets("CDSPS").Select
    Range("B8").Activate
    Exit Sub
    End If
    
    MST
    DIA_CHI
    Sheets("TTDN").Select
    Range("I10:I11").Select
    Selection.ClearContents
    Range("I10").Select
    
    Sheets("CDSPS").Select
    Range("D10:I10").AutoFilter
    Range("D10:I10").AutoFilter
    
    Sheets("SCT156").Select
    Range("SCT156!A14:S14").Select
    Selection.AutoFilter
    Range("SCT156!SCT156_data").ClearContents
    Range("A10").Select
    
    Sheets("NXT").Select
    Range("NXT!A11:S11").Select
    Selection.AutoFilter
    Range("NXT!NXT_data").ClearContents
    Range("I8").Select
    
    Sheets("N").Select
    Range("N!A11:N11").Select
    Selection.AutoFilter
    Range("N!N_dataN").ClearContents
    Range("E8").Select
    
    Sheets("X").Select
    Range("X!A11:N11").Select
    Selection.AutoFilter
    Range("X!X_dataX").ClearContents
    Range("E8").Select
    
    Sheets("PN").Select
    Range("PN!A12:N12").Select
    Selection.AutoFilter
    Range("PN!PN_data").ClearContents
    Range("I9").Select
    
    Sheets("PX").Select
    Range("PX!A12:N12").Select
    Selection.AutoFilter
    Range("PN!PX_data").ClearContents
    Range("I9").Select
    
    Sheets("THU_CHI").Select
    Range("K6:K7").Select
    Selection.ClearContents
    Range("Q3:Z500").Select
    Selection.ClearContents
    Range("Q3").Select
        
    Sheets("BR").Select
    Range("A1:S1").Select
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = False
    Range("K1:P1").Select
    Selection.EntireColumn.Hidden = True
    Range("A2:J2000").Select
    Selection.ClearContents
    Range("A2").Select
    
    Sheets("MV").Select
    Range("A1:S1").Select
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = False
    Range("K1:P1").Select
    Selection.EntireColumn.Hidden = True
    Range("A2:J2000").Select
    Selection.ClearContents
    Range("A2").Select
    
    Sheets("NH").Select
    Range("A1:K1").Select
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = False
    Range("I1").Select
    Selection.EntireColumn.Hidden = True
    Range("A2:H2000").Select
    Selection.ClearContents
    Range("A2").Select
    
    Sheets("Khac").Select
    Range("A1:K1").Select
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = False
    Range("I1").Select
    Selection.EntireColumn.Hidden = True
    Range("A2:H2000").Select
    Selection.ClearContents
    Range("A2").Select
    
    Sheets("131TH").Select
    Range("A11:J11").AutoFilter
    Range("A11:J11").AutoFilter
    Range("A12").Select
    
    Sheets("331TH").Select
    Range("A11:J11").AutoFilter
    Range("A11:J11").AutoFilter
    Range("A12").Select
    
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    Selection.AutoFilter
    Application.Goto Reference:="NKC_data1"
    Selection.ClearContents
    Application.Goto Reference:="NKC_data2"
    Selection.ClearContents
    Range("NKC_VAT642").Select
    Selection.ClearContents
    Range("NKC_VAT642no").Select
    ActiveCell.FormulaR1C1 = "=NKC_VAT642"
    Range("NKC_Cell1541").Select
    Selection.ClearContents
    Range("NKC_Cell1542").Select
    Selection.ClearContents
    Range("DC33311n").Select
    Selection.ClearContents
    Range("NKC_CL8211").Select
    ActiveCell.FormulaR1C1 = "=+R[-1]C[-1]"
    Range("A13").Select
    
    Sheets("NXT152").Select
    Range("A11:O11").AutoFilter
    Range("A11:O11").AutoFilter
    Range("A11:M11").Select
    Selection.EntireColumn.Hidden = False
    Range("B9:B2000").Select
    Selection.EntireRow.Hidden = False
    Application.Goto Reference:="NXT_V152"
    Selection.ClearContents
    Range("B12").Select
    
    Sheets("NXT155").Select
    Range("A11:O11").AutoFilter
    Range("A11:O11").AutoFilter
    Range("A11:M11").Select
    Selection.EntireColumn.Hidden = False
    Range("B9:B2000").Select
    Selection.EntireRow.Hidden = False
    Application.Goto Reference:="NXT_V155"
    Selection.ClearContents
    Range("B12").Select
    
    Sheets("NXT156").Select
    Range("A11:O11").AutoFilter
    Range("A11:O11").AutoFilter
    Range("A11:M11").Select
    Selection.EntireColumn.Hidden = False
    Range("B9:B2000").Select
    Selection.EntireRow.Hidden = False
    Application.Goto Reference:="NXT_V156"
    Selection.ClearContents
    Range("B12").Select
'LayDL_NXT
    Sheets("NKban").Select
    Range("A11:O11").AutoFilter
    Range("A11:M11").Select
    Selection.EntireColumn.Hidden = False
    Range("B9:B2000").Select
    Selection.EntireRow.Hidden = False
    Application.Goto Reference:="NK_Vban"
    Selection.ClearContents
    Range("B12").Select
    
    Sheets("NKmua").Select
    Range("A11:O11").AutoFilter
    Range("A11:M11").Select
    Selection.EntireColumn.Hidden = False
    Range("B9:B2000").Select
    Selection.EntireRow.Hidden = False
    Application.Goto Reference:="NK_Vmua"
    Selection.ClearContents
    Range("B12").Select
        
    Sheets("BL").Select
    Range("A11:R11").Select
    Selection.EntireColumn.Hidden = False
    Range("A8").Activate
    
    Sheets("BR").Select
    Range("A2").Activate
    
   Exit Sub
    Else
  Sheets("CDSPS").Select
  Range("E8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 
 End If
 
End Sub



