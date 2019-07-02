Attribute VB_Name = "Module1"
Sub KC_CDSPS()
   
    Range("A65536").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("A65536").Copy
    Range("A65536").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  Range("E8").Activate
If Range("A65536") = 1 Then
' KIEM TRA BANG LUONG
    If Range("BL_doichieu") <> 0 Then
    Sheets("BL").Select
    Range("AB12").Activate
            R = MsgBox("BANG LUONG bi LECH. Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
       If R = vbYes Then
         Exit Sub
       End If
      End If
    
SQ_XuLy
      If Range("SQ_am") > 0 Then
            R = MsgBox("AM QUY - AM QUY - AM QUY. Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
       If R = vbYes Then
         Exit Sub
      End If
      End If

    
'KET CHUYEN NHAT KY CHUNG
    Sheets("NKC").Select
    Range("D_locnk").Activate
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = False
    
    Range("DateKC").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(thang,date,3,0)"
    Range("DateKC").Copy
    Range("Datev1").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("Datev2").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
If Range("A10") = 1 Then
    Application.Goto Reference:="NKC_FormulaKC"
    Selection.Copy
    Range("KC133n").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Range("KC8211nc3").Select
    ActiveCell.FormulaR1C1 = "=KC8211nc2"
    Range("KC8211nc4").Select
    ActiveCell.FormulaR1C1 = "=KC8211nc1"
    Range("KC8211c").Select
    ActiveCell.FormulaR1C1 = "=KC8211n"
    Range("NKC_CL8211").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    
    Range("KC421n").Select
    ActiveCell.FormulaR1C1 = _
        "=ABS(sum(Vds)-sum(Vcp)-KC8211n)"
    Range("KC421c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("LL1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(Vds)>(SUM(Vcp)+KC8211n),911,4212)"
    
    Range("LL2").Select
    ActiveCell.FormulaR1C1 = "=IF(LL1=911,4212,911)"
    Range("LL3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[1]"
    Range("LL4").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    
Else
   NKC_KETCHUYEN
End If

    'Range("V_kc").Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        
    Range("tgpsno").Select
    ActiveCell.FormulaR1C1 = "=SUM(ST_NO1)"
    Range("tgpsco").Select
    ActiveCell.FormulaR1C1 = "=SUM(ST_CO1)"
    Range("V_tnkc").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Range("ghi_SC").Select
    ActiveCell.FormulaR1C1 = "=IF((RC[5]+RC[6])<>0,""v"","""")"
    Range("NKC_celltrang").Select
    ActiveCell.FormulaR1C1 = "=IF(RC8<>0,VLOOKUP(RC8,NKC_Vtrang,2,1),0)"
    Range("NKC_cellSTT").Select
    ActiveCell.FormulaR1C1 = "=IF((RC[3]+RC[4])<>0,R[-1]C+1,R[-1]C)"
    Range("F13:H13").Copy
    Application.Goto Reference:="NKC_Vtrangso"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("NKC_celltongtrang").Select
    ActiveCell.FormulaR1C1 = "=NKC_celltongtrang1"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    If Range("tgpsno") <> Range("tgpsco") Then
            R = MsgBox("NHAT KY CHUNG KHONG CAN. Ban muon KIEM TRA LAI ko?", vbYesNo, "CHU Y")
       If R = vbYes Then
         Exit Sub
       End If
      End If
   
' LAM BANG CDSPS
    Sheets("CDSPS").Select
    Range("D10:I10").AutoFilter
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "=+NKC!R10C1"
If Range("A8") = 1 Then
    Application.Goto Reference:="CDSPS_FormulaKC"
    Selection.Copy
    Range("psn111").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
 Else
    CDSPS_XL
 End If
' cong tong cong
    Range("tgddn").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg1)-ddn111-ddn112-ddn1121-ddn1122-ddn121-ddn128-ddn133-ddn138-ddn1388-ddn136-ddn1368-ddn141-ddn152-ddn1521-ddn154-ddn155-ddn1551-ddn156-ddn1561-ddn211-ddn212-ddn213-ddn228-ddn229-ddn241-ddn242-ddn244-ddn333-ddn334-ddn336-ddn3368-ddn338-ddn3388-ddn413-ddn421-ddn611-ddn6111-ddn821"
    Range("tgddc").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg2)-ddc121-ddc128-ddc138-ddc1388-ddc136-ddc1368-ddc141-ddc214-ddc228-ddc229-ddc244-ddc333-ddc334-ddc336-ddc3368-ddc338-ddc3388-ddc341-ddc3411-ddc411-ddc413-ddc421"
    Range("tgpsn").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg3)-psn111-psn112-psn1121-psn1122-psn121-psn128-psn133-psn138-psn1388-psn136-psn1368-psn141-psn152-psn1521-psn154-psn155-psn1551-psn156-psn1561-psn211-psn212-psn213-psn214-psn228-psn229-psn241-psn242-psn244-psn333-psn334-psn336-psn3368-psn338-psn3388-psn341-psn3411-psn411-psn413-psn421-psn511-psn5111-psn5112-psn5113-psn515-psn611-psn6111-psn621-psn622-psn623-psn6231-psn6232-psn6233-psn6234-psn6237-psn6238-psn627-psn6271-psn6273-psn6274-psn6278-psn632-psn635-psn641-psn642-psn821"
    Range("tgpsc").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg4)-psc111-psc112-psc1121-psc1122-psc121-psc128-psc133-psc138-psc1388-psc136-psc1368-psc141-psc152-psc1521-psc154-psc155-psc1551-psc156-psc1561-psc211-psc212-psc213-psc214-psc228-psc229-psc241-psc242-psc244-psc333-psc334-psc336-psc3368-psc338-psc3388-psc341-psc3411-psc411-psc413-psc421-psc511-psc5111-psc5112-psc5113-psc515-psc611-psc6111-psc621-psc622-psc623-psc6231-psc6232-psc6233-psc6234-psc6237-psc6238-psc627-psc6271-psc6273-psc6274-psc6278-psc632-psc635-psc641-psc642-psc821"
    Range("tgdcn").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg5)-dcn111-dcn112-dcn1121-dcn1122-dcn121-dcn128-dcn133-dcn138-dcn1388-dcn136-dcn1368-dcn141-dcn152-dcn1521-dcn154-dcn155-dcn1551-dcn156-dcn1561-dcn211-dcn212-dcn213-dcn228-dcn229-dcn241-dcn242-dcn244-dcn333-dcn334-dcn336-dcn3368-dcn338-dcn3388-dcn413-dcn421-dcn611-dcn6111-dcn821"
    Range("tgdcc").Select
    ActiveCell.FormulaR1C1 = "=SUM(vtg6)-dcc121-dcc128-dcc138-dcc1388-dcc136-dcc1368-dcc141-dcc214-dcc228-dcc229-dcc244-dcc333-dcc334-dcc336-dcc3368-dcc338-dcc3388-dcc341-dcc3411-dcc411-dcc413-dcc421"
    
    Range("CDsps").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Range("D_loccdps").Select
    Selection.AutoFilter
    Range("cell_loccdps").Select
    Selection.AutoFilter Field:=7, Criteria1:="1"
    Range("thu_cdps").Select
    ActiveCell.FormulaR1C1 = "=ROUND(tgddn+tgpsn+tgdcn-tgddc-tgpsc-tgdcc,0)"
    
    If Range("thu_cdps") <> 0 Then
            R = MsgBox("Bang CDPS KHONG CAN. Ban muon KIEM TRA LAI ko?", vbYesNo, "CHU Y")
       If R = vbYes Then
         Exit Sub
       End If
      End If
    Range("thu_cdps1").Select
    ActiveCell.FormulaR1C1 = "=round(SUM(V_dccTS)-dcc131-dcc214*2-dcc229*2,0)"
    If Range("thu_cdps1") <> 0 Then
            R = MsgBox("SAI RUI! TK loai I-II KHONG DUOC COSDCK BEN CO(tru 131;214;229). Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
       If R = vbYes Then
         Exit Sub
       End If
      End If
    Range("thu_cdps2").Select
    ActiveCell.FormulaR1C1 = "=round(SUM(V_dcloai56)-dcn821*2,0)"
    If Range("thu_cdps2") <> 0 Then
            R = MsgBox("SAI RUI! TK loai 6-7-8-9 KHONG DUOC CO SDCK(tru 821). Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
       If R = vbYes Then
         Exit Sub
       End If
      End If
    Range("P8").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(dcn133>0,IF(OR(CDSPS_check133="""",CDSPS_check133<>0),1,0),0)"
    Range("CDSPS_check133").Select
    If Range("P8") = 1 Then
            R = MsgBox("VAT duoc KHAU TRU SAI hoac CHUA DOI CHIEU. Ban muon KIEM TRA LAI ko?", vbYesNo, "NGUY HIEM")
    If R = vbYes Then
         Exit Sub
       End If
      End If
    Sheets("LAI_LO").Select
    If Range("LAILO_DC") <> 0 Then
            MsgBox ("LAI_LO LECH. Kiem tra lai co DINH KHOAN c— SAI KHONG.Neu dung thi xac dinh nguyen nhan va ghi nhan vao sheet""NOTE""")
     End If
    
    KC_CDSPS131
    KC_CDSPS331

    Sheets("NKC").Select
    Range("D_locnk").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("cell_locnk").Select
    Selection.AutoFilter Field:=5, Criteria1:="v"
NKC_daucot
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E8").Select
    
    Sheets("CDSPS").Select
    Range("J10").Activate
    
  Exit Sub
  Else
  Sheets("NKC").Select
  Range("E8").Activate
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
 
 End If

End Sub

Sub NKC_daucot()
    Sheets("NKC").Select
    Range("D9:D9").Activate
    Selection.EntireColumn.Hidden = True
    Range("G9:G9").Activate
    Selection.EntireColumn.Hidden = True
    Range("J9").Activate
    Selection.EntireColumn.Hidden = True
End Sub

Sub NKC_KETCHUYEN()

' BUT TOAN KET CHUYEN
    Sheets("NKC").Select
    
' thue
    Range("KC133n").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((ddn1331+SUMIF(SH_TK,""1331"",ST_NO)-SUMIF(SH_TK,""1331"",ST_CO))>(SUMIF(SH_TK,""33311"",ST_CO)-DC33311n),SUMIF(SH_TK,""33311"",ST_CO)-DC33311n,(ddn1331+SUMIF(SH_TK,""1331"",ST_NO)-DC33311n-SUMIF(SH_TK,""1331"",ST_CO)))"
    Range("KC133c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC1332n").Select
    ActiveCell.FormulaR1C1 = _
        "=MAX(IF(((SUMIF(SH_TK,""33311"",ST_CO)-DC33311n)-KC133n)<(ddn1332+SUMIF(SH_TK,""1332"",ST_NO)-SUMIF(SH_TK,""1332"",ST_CO)),(SUMIF(SH_TK,""33311"",ST_CO)-DC33311n)-KC133n,(ddn1332+SUMIF(SH_TK,""1332"",ST_NO)-SUMIF(SH_TK,""1332"",ST_CO)-DC33311n)),0)"
    Range("KC1332c").Select
    ActiveCell.FormulaR1C1 = "=+KC1332n"
    Range("KC133B1n").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((sdn133B1+SUMIF(SH_TK,""133B1"",ST_NO)-SUMIF(SH_TK,""133B1"",ST_CO))>SUMIF(SH_TK,""3331B1"",ST_CO),SUMIF(SH_TK,""3331B1"",ST_CO),(sdn133B1+SUMIF(SH_TK,""133B1"",ST_NO)-SUMIF(SH_TK,""133B1"",ST_CO)))"
    Range("KC133B1c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC133B2n").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((sdn133B2+SUMIF(SH_TK,""133B2"",ST_NO)-SUMIF(SH_TK,""133B2"",ST_CO))>SUMIF(SH_TK,""3331B2"",ST_CO),SUMIF(SH_TK,""3331B2"",ST_CO),(sdn133B2+SUMIF(SH_TK,""133B2"",ST_NO)-SUMIF(SH_TK,""133B2"",ST_CO)))"
    Range("KC133B2c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
' vung thu nhap
    Range("KC5111_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-01"",ST_CO)-SUMIF(SH_TK,""5111-01"",ST_NO)"
    Range("KC5111_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5111_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-02"",ST_CO)-SUMIF(SH_TK,""5111-02"",ST_NO)"
    Range("KC5111_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5111_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-03"",ST_CO)-SUMIF(SH_TK,""5111-03"",ST_NO)"
    Range("KC5111_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5111_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-04"",ST_CO)-SUMIF(SH_TK,""5111-04"",ST_NO)"
    Range("KC5111_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5111_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-05"",ST_CO)-SUMIF(SH_TK,""5111-05"",ST_NO)"
    Range("KC5111_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5111_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5111-06"",ST_CO)-SUMIF(SH_TK,""5111-06"",ST_NO)"
    Range("KC5111_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-01"",ST_CO)-SUMIF(SH_TK,""5112-01"",ST_NO)"
    Range("KC5112_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-02"",ST_CO)-SUMIF(SH_TK,""5112-02"",ST_NO)"
    Range("KC5112_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-03"",ST_CO)-SUMIF(SH_TK,""5112-03"",ST_NO)"
    Range("KC5112_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-04"",ST_CO)-SUMIF(SH_TK,""5112-04"",ST_NO)"
    Range("KC5112_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-05"",ST_CO)-SUMIF(SH_TK,""5112-05"",ST_NO)"
    Range("KC5112_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5112_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5112-06"",ST_CO)-SUMIF(SH_TK,""5112-06"",ST_NO)"
    Range("KC5112_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-01"",ST_CO)-SUMIF(SH_TK,""5113-01"",ST_NO)"
    Range("KC5113_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-02"",ST_CO)-SUMIF(SH_TK,""5113-02"",ST_NO)"
    Range("KC5113_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-03"",ST_CO)-SUMIF(SH_TK,""5113-03"",ST_NO)"
    Range("KC5113_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-04"",ST_CO)-SUMIF(SH_TK,""5113-04"",ST_NO)"
    Range("KC5113_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-05"",ST_CO)-SUMIF(SH_TK,""5113-05"",ST_NO)"
    Range("KC5113_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5113_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5113-06"",ST_CO)-SUMIF(SH_TK,""5113-06"",ST_NO)"
    Range("KC5113_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5114n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5114"",ST_CO)-SUMIF(SH_TK,""5114"",ST_NO)"
    Range("KC5114c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC5117n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5117"",ST_CO)-SUMIF(SH_TK,""5117"",ST_NO)"
    Range("KC5117c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
     Range("KC5118n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""5118"",ST_CO)-SUMIF(SH_TK,""5118"",ST_NO)"
    Range("KC5118c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-01"",ST_CO)-SUMIF(SH_TK,""515-01"",ST_NO)"
    Range("KC515_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-02"",ST_CO)-SUMIF(SH_TK,""515-02"",ST_NO)"
    Range("KC515_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-03"",ST_CO)-SUMIF(SH_TK,""515-03"",ST_NO)"
    Range("KC515_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-04"",ST_CO)-SUMIF(SH_TK,""515-04"",ST_NO)"
    Range("KC515_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-05"",ST_CO)-SUMIF(SH_TK,""515-05"",ST_NO)"
    Range("KC515_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-06"",ST_CO)-SUMIF(SH_TK,""515-06"",ST_NO)"
    Range("KC515_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_07n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-07"",ST_CO)-SUMIF(SH_TK,""515-07"",ST_NO)"
    Range("KC515_07c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC515_08n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""515-08"",ST_CO)-SUMIF(SH_TK,""515-08"",ST_NO)"
    Range("KC515_08c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC711n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""711"",ST_CO)-SUMIF(SH_TK,""711"",ST_NO)"
    Range("KC711c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
' vung chi phi
    Range("KC632_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-01"",ST_NO)-SUMIF(SH_TK,""632-01"",ST_CO)"
    Range("KC632_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-02"",ST_NO)-SUMIF(SH_TK,""632-02"",ST_CO)"
    Range("KC632_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-03"",ST_NO)-SUMIF(SH_TK,""632-03"",ST_CO)"
    Range("KC632_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-04"",ST_NO)-SUMIF(SH_TK,""632-04"",ST_CO)"
    Range("KC632_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-05"",ST_NO)-SUMIF(SH_TK,""632-05"",ST_CO)"
    Range("KC632_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-06"",ST_NO)-SUMIF(SH_TK,""632-06"",ST_CO)"
    Range("KC632_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_07n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-07"",ST_NO)-SUMIF(SH_TK,""632-07"",ST_CO)"
    Range("KC632_07c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC632_08n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""632-08"",ST_NO)-SUMIF(SH_TK,""632-08"",ST_CO)"
    Range("KC632_08c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6411n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6411"",ST_NO)-SUMIF(SH_TK,""6411"",ST_CO)"
    Range("KC6411c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6412n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6412"",ST_NO)-SUMIF(SH_TK,""6412"",ST_CO)"
    Range("KC6412c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6413n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6413"",ST_NO)-SUMIF(SH_TK,""6413"",ST_CO)"
    Range("KC6413c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6414n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6414"",ST_NO)-SUMIF(SH_TK,""6414"",ST_CO)"
    Range("KC6414c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6415n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6415"",ST_NO)-SUMIF(SH_TK,""6415"",ST_CO)"
    Range("KC6415c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6417n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6417"",ST_NO)-SUMIF(SH_TK,""6417"",ST_CO)"
    Range("KC6417c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6418n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6418"",ST_NO)-SUMIF(SH_TK,""6418"",ST_CO)"
    Range("KC6418c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC64180n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""64180"",ST_NO)-SUMIF(SH_TK,""64180"",ST_CO)"
    Range("KC64180c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6421n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6421"",ST_NO)-SUMIF(SH_TK,""6421"",ST_CO)"
    Range("KC6421c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6422n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6422"",ST_NO)-SUMIF(SH_TK,""6422"",ST_CO)"
    Range("KC6422c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6423n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6423"",ST_NO)-SUMIF(SH_TK,""6423"",ST_CO)"
    Range("KC6423c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6424n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6424"",ST_NO)-SUMIF(SH_TK,""6424"",ST_CO)"
    Range("KC6424c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6425n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6425"",ST_NO)-SUMIF(SH_TK,""6425"",ST_CO)"
    Range("KC6425c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6426n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6426"",ST_NO)-SUMIF(SH_TK,""6426"",ST_CO)"
    Range("KC6426c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6427n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6427"",ST_NO)-SUMIF(SH_TK,""6427"",ST_CO)"
    Range("KC6427c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC6428n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""6428"",ST_NO)-SUMIF(SH_TK,""6428"",ST_CO)"
    Range("KC6428c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC64280n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""64280"",ST_NO)-SUMIF(SH_TK,""64280"",ST_CO)"
    Range("KC64280c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_01n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-01"",ST_NO)-SUMIF(SH_TK,""635-01"",ST_CO)"
    Range("KC635_01c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_02n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-02"",ST_NO)-SUMIF(SH_TK,""635-02"",ST_CO)"
    Range("KC635_02c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_03n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-03"",ST_NO)-SUMIF(SH_TK,""635-03"",ST_CO)"
    Range("KC635_03c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_04n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-04"",ST_NO)-SUMIF(SH_TK,""635-04"",ST_CO)"
    Range("KC635_04c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_05n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-05"",ST_NO)-SUMIF(SH_TK,""635-05"",ST_CO)"
    Range("KC635_05c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_06n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-06"",ST_NO)-SUMIF(SH_TK,""635-06"",ST_CO)"
    Range("KC635_06c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_07n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-07"",ST_NO)-SUMIF(SH_TK,""635-07"",ST_CO)"
    Range("KC635_07c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC635_08n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""635-08"",ST_NO)-SUMIF(SH_TK,""635-08"",ST_CO)"
    Range("KC635_08c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("KC811n").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK,""811"",ST_NO)-SUMIF(SH_TK,""811"",ST_CO)"
    Range("KC811c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
            
    Range("KC8211nc3").Select
    ActiveCell.FormulaR1C1 = "=KC8211nc2"
    Range("KC8211nc4").Select
    ActiveCell.FormulaR1C1 = "=KC8211nc1"
    Range("KC8211c").Select
    ActiveCell.FormulaR1C1 = "=KC8211n"
    Range("NKC_CL8211").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    
    Range("KC421n").Select
    ActiveCell.FormulaR1C1 = _
        "=ABS(sum(Vds)-sum(Vcp)-KC8211n)"
    Range("KC421c").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
    Range("LL1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(Vds)>(SUM(Vcp)+KC8211n),911,4212)"
    
    Range("LL2").Select
    ActiveCell.FormulaR1C1 = "=IF(LL1=911,4212,911)"
    Range("LL3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[1]"
    Range("LL4").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
End Sub
Sub CDSPS_XL()

' So phat sinh trong ky
    Sheets("CDSPS").Select
    Range("E20").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK1,RC1,ST_NO1)"
    Range("F20").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SH_TK1,RC1,ST_CO1)"
' so du cuoi ky
    Range("G20").Select
    ActiveCell.FormulaR1C1 = _
        "=max((RC[-4]+RC[-2])-(RC[-3]+RC[-1]),0)"
    Range("H20").Select
    ActiveCell.FormulaR1C1 = _
        "=max((RC[-2]+RC[-4])-(RC[-5]+RC[-3]),0)"
    Range("E20:H20").Select
' copy ca vung
    Selection.Copy
    Range("CD").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I20").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(RC[-4]:RC[-1])<>0,1,0)"
    Selection.Copy
    Range("V_loccdps").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
' xu ly phan tong phat sinh Tk cap 1-2
    Range("psn133").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn133,1,0,2,1))"
    Range("psc133").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc133,1,0,2,1))"
    Range("psn111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn111,1,0,3,1))"
    Range("psc111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc111,1,0,3,1))"
    Range("psn1121").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1121,1,0,12,1))"
    Range("psc1121").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1121,1,0,12,1))"
    Range("psn1122").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1122,1,0,8,1))"
    Range("psc1122").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1122,1,0,8,1))"
    Range("psn112").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn112,1,0,23,1))-psn1121-psn1122"
    Range("psc112").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc112,1,0,23,1))-psc1121-psc1122"
    Range("psn121").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn121,1,0,3,1))"
    Range("psc121").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc121,1,0,3,1))"
    Range("psn128").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn128,1,0,4,1))"
    Range("psc128").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc128,1,0,4,1))"
    Range("psn1368").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1368,1,0,5,1))"
    Range("psc1368").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1368,1,0,5,1))"
    Range("psn136").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn136,1,0,4,1))"
    Range("psc136").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc136,1,0,4,1))"
    Range("psn1388").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1388,1,0,8,1))"
    Range("psc1388").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1388,1,0,8,1))"
    Range("psn138").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn138,1,0,3,1))"
    Range("psc138").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc138,1,0,3,1))"
    Range("psn141").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn141,1,0,2,1))"
    Range("psc141").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc141,1,0,2,1))"
    Range("psn1521").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1521,1,0,4,1))"
    Range("psc1521").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1521,1,0,4,1))"
    Range("psn152").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn152,1,0,6,1))-psn1521"
    Range("psc152").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc152,1,0,6,1))-psc1521"
    Range("psn154").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn154,1,0,8,1))"
    Range("psc154").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc154,1,0,8,1))"
    Range("psn155").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn155,1,0,6,1))-psn1551"
    Range("psc155").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc155,1,0,6,1))-psc1551"
    Range("psn1551").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1551,1,0,4,1))"
    Range("psc1551").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1551,1,0,4,1))"
    Range("psn1561").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn1561,1,0,6,1))"
    Range("psc1561").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc1561,1,0,6,1))"
    Range("psn156").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn156,1,0,9,1))-psn1561"
    Range("psc156").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc156,1,0,9,1))-psc1561"
    Range("psn211").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn211,1,0,6,1))"
    Range("psc211").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc211,1,0,6,1))"
    Range("psn212").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn211,1,0,2,1))"
    Range("psc212").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc211,1,0,2,1))"
    Range("psn213").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn213,1,0,7,1))"
    Range("psc213").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc213,1,0,7,1))"
    Range("psn214").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn214,1,0,4,1))"
    Range("psc214").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc214,1,0,4,1))"
    Range("psn228").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn228,1,0,2,1))"
    Range("psc228").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc228,1,0,2,1))"
    Range("psn229").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn229,1,0,4,1))"
    Range("psc229").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc229,1,0,4,1))"
    Range("psn241").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn241,1,0,3,1))"
    Range("psc241").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc241,1,0,3,1))"
    Range("psn242").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn242,1,0,4,1))"
    Range("psc242").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc242,1,0,4,1))"
    Range("psn244").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn244,1,0,4,1))"
    Range("psc244").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc244,1,0,4,1))"
    Range("psn333").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn333,1,0,10,1))"
    Range("psc333").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc333,1,0,10,1))"
    'Range("dcn333").Select
    'ActiveCell.FormulaR1C1 = "=sum(offset(dcn333,1,0,10,1))"
    'Range("dcc333").Select
    'ActiveCell.FormulaR1C1 = "=sum(offset(dcc333,1,0,10,1))"
    Range("psn334").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn334,1,0,2,1))"
    Range("psc334").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc334,1,0,2,1))"
    Range("psn3368").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn3368,1,0,5,1))"
    Range("psc3368").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc3368,1,0,5,1))"
    Range("psn336").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn336,1,0,4,1))"
    Range("psc336").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc336,1,0,4,1))"
    Range("psn3388").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn3388,1,0,5,1))"
    Range("psc3388").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc3388,1,0,5,1))"
    Range("psn338").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn338,1,0,13,1))-psn3388"
    Range("psc338").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc338,1,0,13,1))-psc3388"
    Range("psn3411").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn3411,1,0,8,1))"
    Range("psc3411").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc3411,1,0,8,1))"
    Range("psn341").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn341,1,0,10,1))-psn3411"
    Range("psc341").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc341,1,0,10,1))-psc3411"
    Range("psn411").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn411,1,0,4,1))"
    Range("psc411").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc411,1,0,4,1))"
    Range("psn413").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn413,1,0,2,1))"
    Range("psc413").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc413,1,0,2,1))"
    Range("psn421").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn421,1,0,2,1))"
    Range("psc421").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc421,1,0,2,1))"
    Range("psn5111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn5111,1,0,6,1))"
    Range("psc5111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc5111,1,0,6,1))"
    Range("psn5112").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn5112,1,0,6,1))"
    Range("psc5112").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc5112,1,0,6,1))"
    Range("psn5113").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn5113,1,0,6,1))"
    Range("psc5113").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc5113,1,0,6,1))"
    Range("psn511").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn511,1,0,24,1))-psn5111-psn5112-psn5113"
    Range("psc511").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc511,1,0,24,1))-psc5111-psc5112-psc5113"
    Range("psn515").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn515,1,0,8,1))"
    Range("psc515").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc515,1,0,8,1))"
    Range("psn521").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn521,1,0,3,1))"
    Range("psc521").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc521,1,0,3,1))"
    Range("psc611").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc611,1,0,6,1))-psc6111"
    Range("psn611").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn611,1,0,6,1))-psn6111"
    Range("psc6111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6111,1,0,4,1))"
    Range("psn6111").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6111,1,0,4,1))"
    Range("psc621").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc621,1,0,4,1))"
    Range("psn621").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn621,1,0,4,1))"
    Range("psn622").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn622,1,0,4,1))"
    Range("psc622").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc622,1,0,4,1))"
    Range("psn6231").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6231,1,0,8,1))"
    Range("psc6231").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6231,1,0,8,1))"
    Range("psn6232").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6232,1,0,8,1))"
    Range("psc6232").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6232,1,0,8,1))"
    Range("psn6233").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6233,1,0,8,1))"
    Range("psc6233").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6233,1,0,8,1))"
    Range("psn6234").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6234,1,0,8,1))"
    Range("psc6234").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6234,1,0,8,1))"
    Range("psn6237").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6237,1,0,8,1))"
    Range("psc6237").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6237,1,0,8,1))"
    Range("psn6238").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6238,1,0,8,1))"
    Range("psc6238").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6238,1,0,8,1))"
    Range("psn623").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn623,1,0,54,1))-psn6231-psn6232-psn6233-psn6234-psn6237-psn6238"
    Range("psc623").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc623,1,0,54,1))-psc6231-psc6232-psc6233-psc6234-psc6237-psc6238"
    Range("psn6271").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6271,1,0,4,1))"
    Range("psc6271").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6271,1,0,4,1))"
    Range("psn6273").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6273,1,0,4,1))"
    Range("psc6273").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6273,1,0,4,1))"
    Range("psn6274").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6274,1,0,4,1))"
    Range("psc6274").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6274,1,0,4,1))"
    Range("psn6278").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn6278,1,0,4,1))"
    Range("psc6278").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc6278,1,0,4,1))"
    Range("psn627").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn627,1,0,23,1))-psn6271-psn6273-psn6274-psn6278"
    Range("psc627").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc627,1,0,23,1))-psc6271-psc6273-psc6274-psc6278"
    Range("psn631").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn631,1,0,4,1))"
    Range("psc631").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc631,1,0,4,1))"
    Range("psn621").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn621,1,0,4,1))"
    Range("psn632").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn632,1,0,8,1))"
    Range("psc632").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc632,1,0,8,1))"
    Range("psn635").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn635,1,0,8,1))"
    Range("psc635").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc635,1,0,8,1))"
    Range("psn641").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn641,1,0,8,1))"
    Range("psc641").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc641,1,0,8,1))"
    Range("psn642").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn642,1,0,9,1))"
    Range("psc642").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc642,1,0,9,1))"
    Range("psn821").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psn821,1,0,2,1))"
    Range("psc821").Select
    ActiveCell.FormulaR1C1 = "=sum(offset(psc821,1,0,2,1))"

End Sub

