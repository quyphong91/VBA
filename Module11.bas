Attribute VB_Name = "Module11"
Sub lamnk()
 
    Range("K1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(and(MID(CELL(""filename""),IF(ISERROR(FIND(""-2018"",CELL(""filename"")))=TRUE,1,FIND(""-2018"",CELL(""filename""))+1),4)=""2018"",(YEAR(NKC!R1C251)+YEAR(NKC!R2C251)+YEAR(NKC!R3C251)+YEAR(NKC!R4C251)+YEAR(NKC!R5C251)+YEAR(NKC!R6C251)+YEAR(NKC!R7C251)+YEAR(NKC!R8C251)+YEAR(NKC!R9C251)+YEAR(NKC!R10C251)+YEAR(NKC!R11C251)+YEAR(NKC!R12C251))=24204),1,0)"
    Range("K1").Copy
    Range("K1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

If Range("K1") = 1 Then
MST
DIA_CHI
nk1
Sheets("TTDN").Select
If Range("J2") = 0 Then
Sheets("nk").Select
Range("A3:J3").Select
Selection.AutoFilter

'xoa noi dung cu
    Range("a3:I10046").ClearContents
'Chep noi dung tu Nk1 sang Nk de lam nhat ky chung
    Range("NK1!C3:D5000").Copy Destination:=Range("NK!B3")
    Range("NK1!F3:F5000").Copy Destination:=Range("NK!D3")
    Range("NK1!L3:L5000").Copy Destination:=Range("NK!E3")
    Range("NK1!J3:K5000").Copy Destination:=Range("NK!F3")
    Range("NK1!H3:H5000").Copy Destination:=Range("NK!H3")
' Ngay thang ghi so
    Range("A3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC3="""",""End"",IF(MONTH(RC3)<>MONTH(R2C9),R2C9,RC3))"
    Range("A3:A5000").FillDown
    Range("A3:D5000").Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    'Range("A3:C1000").Copy
    Range("A5001").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    'Range("A3:D1000").Copy
    Range("A10001").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    'Range("A3:C1000").Copy
    Range("A15000").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
' noi dung
    Range("E3:E5000").Copy Destination:=Range("E5001")
' tk doi ung
    Range("F3:F5000").Copy Destination:=Range("G5001")
    Range("G3:G5000").Copy Destination:=Range("F5001")
' so tien co
    Range("H3:H5000").Copy Destination:=Range("I5001")
' NOI DUNG THUE
    Range("NK1!M3:M5000").Copy
    Range("NK!E10001").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("NK!E15000").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
 ' TK DOI UNG THUE
    Range("NK1!N3:N5000").Copy
    Range("NK!F10001").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("NK!G15000").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("NK1!O3:O5000").Copy
    Range("NK!G10001").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("NK!F15000").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
 ' SO TIEN THUE
    Range("NK1!I3:I5000").Copy
    Range("NK!H10001").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("NK!I15000").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ' Hoan thien NK1
    Sheets("NK1").Select
    Range("L3").Select
    ActiveCell.FormulaR1C1 = _
        "=if(RC9<>"""",(if(or(LEFT(RC3,2)=""PC"",LEFT(RC10,2)=""15""),RC10&""/133"",RC10)),RC10)"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC9<>"""",LEFT(RC11,2)=""51""),RC11&""/3331"",RC11)"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = _
        "=RC8+RC9"
    Range("L3:N20000").FillDown
    Range("L3:N20000").Copy
    Range("NK1!J3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ' XOA DONG DU LIEU NK1
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "DATA"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "Erase"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC7<>"""",RC11<>""""),""Keep"",""Erase"")"
    Range("P3:P20001").FillDown
    Range("P1:P20001").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("P1:P2"), Unique:=False
    Range("A3:T20000").Select
    Selection.ClearContents
    Range("A2:L2").AutoFilter
    Range("A2:L2").AutoFilter
    Range("M3:T20000").Select
    Selection.ClearContents
    Range("B1").Select
 ' DANH SO THU TU
    Sheets("NK").Select
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("J3:J4").Select
    Selection.AutoFill Destination:=Range("J3:J5000"), Type:=xlFillDefault
    Range("J3:J5000").Copy Destination:=Range("J5001")
    Range("J10001").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("J10002").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("J10001:J10002").Select
    Selection.AutoFill Destination:=Range("J10001:J14999"), Type:=xlFillDefault
    Range("J10001:J14999").Copy Destination:=Range("J15000")
  ' TRON NHAT KY
    Range("A3:J20001").Select
    Application.CutCopyMode = False
    Selection.Sort Key1:=Range("C3"), Order1:=xlAscending, Key2:=Range("J3") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
 ' XOA DONG KHONG DU LIEU NK
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC5<>"""",SUM(RC8:RC9)<>0),""MyNK"",""TrongRong"")"
    Range("J3:J20001").FillDown
    Range("A2:J2").Select
    Selection.AutoFilter
    Range("J2").Select
    Selection.AutoFilter Field:=10, Criteria1:="TrongRong"
    Range("J3:J20001").Select
    Selection.EntireRow.Delete
    Selection.AutoFilter Field:=10
    Range("NK!J2:P20000").ClearContents
    Range("J2").Select
End If

Exit Sub
  Else
  Sheets("NK").Select
  Range("I2").Select
  MsgBox " SORRY. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If
   
End Sub


Sub PB_KH()
    Sheets("PB_KH").Select
    Range("A1:I1").Select
    Selection.AutoFilter
    Range("A2:I1000").Select
    Selection.ClearContents
    'copyright
    Sheets("PB242").Select
    Range("A8:N8").AutoFilter
    Range("IN8").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC11<>0,(IF(RC4<>0,RC4,"""")),"""")"
    Range("IO8").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC11<>0,VLOOKUP(thang,Date,3,0),"""")"
    Range("IR8").Select
    ActiveCell.FormulaR1C1 = "=+""CP PB242-""&RC3"
    Range("IS8").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC11<>0,RC10,0)"
    Range("IT8").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC11<>0,RC13,"""")"
    Range("IU8").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC11>0,RC14,"""")"
    Range("IV8").Select
    ActiveCell.FormulaR1C1 = "3"
    Application.Goto Reference:="PB242_Khacdata1"
    Selection.Copy
    Application.Goto Reference:="PB242_Khacdata"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Selection.Copy
    Sheets("PB_KH").Select
    Range("A500").Select
    ActiveSheet.Paste
        
    Sheets("KH").Select
    Range("A12:W12").AutoFilter
    Range("IN12").Select
    ActiveCell.FormulaR1C1 = "=+IF(AND(MID(RC4,3,4)<>""TSC_"",RC20<>0),(IF(RC7<>0,RC7,"""")),"""")"
    Range("IO12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(MID(RC4,3,4)<>""TSC_"",RC20<>0),VLOOKUP(thang,Date,3,0),"""")"
    Range("IR12").Select
    ActiveCell.FormulaR1C1 = "=+""CP KH-""&RC4"
    Range("IS12").Select
    ActiveCell.FormulaR1C1 = "=+IF(AND(MID(RC4,3,4)<>""TSC_"",RC20<>0),RC19,0)"
    Range("IT12").Select
    ActiveCell.FormulaR1C1 = "=+IF(AND(MID(RC4,3,4)<>""TSC_"",RC20<>0),RC21,"""")"
    Range("IU12").Select
    ActiveCell.FormulaR1C1 = "=+IF(AND(MID(RC4,3,4)<>""TSC_"",RC20<>0),RC22,"""")"
    Range("IV12").Select
    ActiveCell.FormulaR1C1 = "4"
    Application.Goto Reference:="KH_Khacdata1"
    Selection.Copy
    Application.Goto Reference:="KH_Khacdata"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Selection.Copy
    Sheets("PB_KH").Select
    Range("A901").Select
    ActiveSheet.Paste
    
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=+IF(AND(RC6<>"""",RC6<>0),1,""x"")"
    Range("I2:I1000").Select
    Selection.FillDown
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "x"
    Range("I1:I1000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("I1:I2"), Unique:=False
    Range("A2:I1000").Select
    Selection.ClearContents
    Range("A1:I1").Select
    Selection.AutoFilter
    Range("A2:I1000").Select
    Selection.Sort Key1:=Range("B26"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC1<>"""",RC5&"" (""&RC1&"" )"",RC5)"
    Range("I2:I1000").Select
    Selection.FillDown
    Range("J1").Select
    Range("E6").Select
    Sheets("PB242").Select
    Application.Goto Reference:="PB242_Khacdata"
    Selection.ClearContents
    Range("E6").Select
    Sheets("KH").Select
    Application.Goto Reference:="KH_Khacdata"
    Selection.ClearContents
    Range("E8").Select
    
End Sub

Sub nk1()
   Sheets("TTDN").Select
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=+IF((YEAR(NOW())-khoaso)>0,IF(MONTH(NOW())>4,1,0),0)"
If Range("J2") = 0 Then
   Sheets("NK").Select
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=+VLOOKUP(thang,Date,2,0)"
   Sheets("NK1").Select
   Range("A2:K2").AutoFilter

Range("NK1!a3:m4002").ClearContents
        'Hoan chinh phan ban ra
        Sheets("br").Select
        Range("A1:T1").Select
        Selection.AutoFilter
        Range("K2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC1<>"""",RC5&"" ""&R1C11&"" ""&RC1,RC5)"
        Range("L2").Select
        ActiveCell.FormulaR1C1 = _
        "=IF(RC9<>"""",(IF(LEFT(RC5,3)=""TSC"",""TSC_: ""&R1C12&RC1,R1C&RC1)),"""")"
        Range("M2").Select
        ActiveCell.FormulaR1C1 = "=RC9"
        Range("N2").Select
        ActiveCell.FormulaR1C1 = _
        "=+IF(RC9<>"""",(IF(LEFT(RIGHT(RC20,2),1)=""B"",RC20,33311)),"""")"
        Range("O2").Select
        ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC2=R[-1]C2,RC4=R[-1]C4),R[-1]C15,R[-1]C15+1)"
        Range("P2").Select
        ActiveCell.FormulaR1C1 = "=+SUMIF(R2C15:R1001C15,RC15,R2C17:R1001C17)"
        Range("Q2").Select
        ActiveCell.FormulaR1C1 = "=RC6+RC8"
        Range("R2").Select
        ActiveCell.FormulaR1C1 = "=+IF(RC17>=20000000,1,0)"
        Range("S2").Select
        ActiveCell.FormulaR1C1 = _
        "=+IF(OR(RC9=131,RC10=131),IF(ISNA(VLOOKUP(RC4,TTKH_131TH,3,0))=TRUE,1,0),0)+IF(OR(RC9=331,RC10=331),IF(ISNA(VLOOKUP(RC4,TTKH_331TH,3,0))=TRUE,1,0),0)"
        Range("K2:S1001").FillDown
'Chep ban ra
Range("BR!A2:A1001").Copy
Range("NK1!B3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("BR!B2:F1001").Copy
Range("NK1!D3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("BR!H2:N1001").Copy
Range("NK1!I3").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        
'Hoan chinh phan mua vao truoc khi chep
        Sheets("MV").Select
        Range("A1:T1").Select
        Selection.AutoFilter
        Range("K2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC1<>"""",RC5&"" ""&R1C11&"" ""&RC1,RC5)"
        Range("L2").Select
        ActiveCell.FormulaR1C1 = _
        "=IF(RC9<>"""",(IF(OR(VALUE(LEFT(RC9,2))=21,VALUE(LEFT(RC9,3))=241),""TSC_: ""&R1C12&RC1,R1C12&RC1)),"""")"
        
        Range("N2").Select
        ActiveCell.FormulaR1C1 = "=RC10"
        Range("M2").Select
        ActiveCell.FormulaR1C1 = _
        "=IF(RC9<>"""",(IF(LEFT(RIGHT(RC20,2),1)=""B"",RC20,(IF(OR(VALUE(LEFT(RC9,2))=21,VALUE(LEFT(RC9,3))=241),1332,1331)))),"""")"
        Range("O2").Select
        ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC2=R[-1]C2,RC4=R[-1]C4),R[-1]C15,R[-1]C15+1)"
        Range("P2").Select
        ActiveCell.FormulaR1C1 = "=+SUMIF(R2C15:R1001C15,RC15,R2C17:R1001C17)"
        Range("Q2").Select
        ActiveCell.FormulaR1C1 = "=RC6+RC8"
        Range("R2").Select
        ActiveCell.FormulaR1C1 = "=+IF(RC16>=20000000,1,0)"
        Range("S2").Select
        ActiveCell.FormulaR1C1 = _
        "=+IF(OR(RC9=131,RC10=131),IF(ISNA(VLOOKUP(RC4,TTKH_131TH,3,0))=TRUE,1,0),0)+IF(OR(RC9=331,RC10=331),IF(ISNA(VLOOKUP(RC4,TTKH_331TH,3,0))=TRUE,1,0),0)"
        Range("K2:S1001").FillDown
        
'Chep mua vao
Range("MV!A2:A1001").Copy
Range("NK1!B1002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("MV!B2:F1001").Copy
Range("NK1!D1002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("MV!H2:N1001").Copy
Range("NK1!I1002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        
'Hoan chinh phan ngan hang truoc khi chep
        Sheets("NH").Select
        Range("A1:T1").Select
        Selection.AutoFilter
        Range("I2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC1<>"""",RC5&"" ""&R1C9&"" ""&RC1,RC5)"
        Sheets("NH").Select
        Range("J2").Select
        ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC7=131,RC8=131),IF(ISNA(VLOOKUP(RC4,TTKH_131TH,3,0))=TRUE,1,0),0)+IF(OR(RC7=331,RC8=331),IF(ISNA(VLOOKUP(RC4,TTKH_331TH,3,0))=TRUE,1,0),0)"
        Range("I2:J1001").FillDown
'Chep ngan hang

Range("NH!A2:A1001").Copy
Range("NK1!B2002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("NH!B2:F1001").Copy
Range("NK1!D2002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("NH!G2:I1001").Copy
Range("NK1!J2002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
'Hoan thanh phan khac truoc khi chep
        Sheets("khac").Select
        Range("A1:T1").Select
        Selection.AutoFilter
        Range("I2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC1<>"""",RC5&"" ""&R1C9&"" ""&RC1,RC5)"
        Sheets("Khac").Select
        Range("J2").Select
        ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC7=131,RC8=131),IF(ISNA(VLOOKUP(RC4,TTKH_131TH,3,0))=TRUE,1,0),0)+IF(OR(RC7=331,RC8=331),IF(ISNA(VLOOKUP(RC4,TTKH_331TH,3,0))=TRUE,1,0),0)"
        Range("I2:J1001").FillDown
'Chep cac phat sinh khac
Range("KHAC!A2:A1001").Copy
Range("NK1!B3002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("KHAC!B2:F1001").Copy
Range("NK1!D3002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("KHAC!G2:I1001").Copy
Range("NK1!J3002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
' PHAN BO - KHAU HAO
PB_KH
Range("PB_KH!A2:A1000").Copy
Range("NK1!B4002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("PB_KH!B2:F1000").Copy
Range("NK1!D4002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("PB_KH!G2:I1000").Copy
Range("NK1!J4002").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False

'Sort lai VË loai bo nhung dong trong:
Sheets("NK1").Select
Range("A3:P5001").Sort Key1:=Range("D3"), Order1:=xlAscending, Key2:=Range("A3") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
Range("M1").Select
    ActiveCell.FormulaR1C1 = "=NK!R2C9"
Range("A3").Select
ActiveCell.FormulaR1C1 = _
        "=IF(RC4="""","""",IF(MONTH(RC4)<>MONTH(R1C13),R1C13,RC4))"
    Range("A3:A5000").FillDown
    Range("A3:A5000").Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Range("Q1").Select
ActiveCell.FormulaR1C1 = "DATA"
Range("Q2").Select
ActiveCell.FormulaR1C1 = "=IF(AND(RC10<>"""",RC11<>""""),""Keep"",""Erase"")"
Range("Q2:Q5000").FillDown
Range("Q2").Select
ActiveCell.FormulaR1C1 = "Erase"
Range("Q1:Q5000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Q1:Q2"), Unique:=False
Range("A3:Q5000").Select
Selection.ClearContents
Range("A2:K2").AutoFilter
Range("A2:K2").AutoFilter
Range("Q1:Q5000").ClearContents
Range("a3").Select
'Danh phieu thu chi
'DanhPTC
DanhTHU_CHI
'Sort NK1 lai de uu tien phieu thu len truoc:
Sheets("NK1").Select
Range("P3").Select
ActiveCell.FormulaR1C1 = "=IF(AND(RC3="""",LEFT(RC7,2)<>""CP""),1,IF(LEFT(RC3,2)=""PT"",2,(IF(LEFT(RC3,2)=""PC"",3,4))))"
Range("P3:P5000").FillDown
Range("P3:P5000").Copy
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False

Range("A3:P5001").Sort Key1:=Range("D3"), Order1:=xlAscending, Key2:=Range("P3") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
Exit Sub
  Else
  Sheets("NK").Select
  Range("K2").Select

End If

End Sub
Sub DanhPTC()
Dim J, k, thu, chi As Integer
k = 1: thu = 1:  chi = 1
With Range("D1")
        Do While IsEmpty(.Offset(k, 0).Value) = False
        If .Offset(k, 6) = "1111" Then
           .Offset(k, -1).Value = "PT" & Format(Month(.Offset(0, 9)), "00") & "-" & Format(thu, "000")
            thu = thu + 1
        ElseIf .Offset(k, 7) = "1111" Then
            .Offset(k, -1).Value = "PC" & Format(Month(.Offset(0, 9)), "00") & "-" & Format(chi, "000")
            chi = chi + 1
        End If
    k = k + 1
    Loop
End With
End Sub


Sub DanhTHU_CHI()

    Range("AA3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC8>0,RC10=1111),(IF(AND(RC2=R[-1]C2,RC4=R[-1]C4,RC5=R[-1]C5,RC10=R[-1]C10),R[-1]C27,R[-1]C27+1)),R[-1]C27)"
    Range("AB3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC8>0,RC11=1111),(IF(AND(RC2=R[-1]C2,RC4=R[-1]C4,RC5=R[-1]C5,RC11=R[-1]C11),R[-1]C28,R[-1]C28+1)),R[-1]C28)"
    Range("AC3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC8>0,RC10=1111),""PT""&TEXT(thang,""00"")&""-""&TEXT(RC27,""000""),"""")"
    Range("AD3").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC8>0,RC11=1111),""PC""&TEXT(thang,""00"")&""-""&TEXT(RC28,""000""),"""")"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=+IF(LEFT(RC29,2)=""PT"",RC29,RC30)"
    Range("AA3:AE1000").Select
    Selection.FillDown
    Range("AE3:AE1000").Select
    Selection.Copy
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA3:AE1000").Select
    Selection.ClearContents

End Sub

