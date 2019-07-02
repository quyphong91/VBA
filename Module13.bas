Attribute VB_Name = "Module13"
Function VNDUni(baonhieu)
' Tien Viet tieng Viet Font Unicode

Dim KetQua, SoTien, Nhom, Chu, Dich, S1, S2, S3 As String
Dim i, J, ViTri As Byte, S As Double
Dim Hang, Doc, Dem
If baonhieu = 0 Then
KetQua = "Kh" & ChrW$(244) & "ng " & ChrW$(273) & ChrW$(7891) & "ng"
Else
If Abs(baonhieu) >= 1E+15 Then
KetQua = "S" & ChrW$(7889) & " qu" & ChrW$(225) & " l" & ChrW$(7899) & "n - H" & ChrW$(224) & "m " & ChrW$(273) & ChrW$(7893) & "i s" & ChrW$(7889) & " ra ch" & ChrW$(7919) & " Vi" & ChrW$(7879) & "t Nam; font ch" & ChrW$(7919) & " .Vntime - Copyright by MaiKa of AQN (0953-357-988)"
Else
If baonhieu < 0 Then
KetQua = ChrW$(194) & "m" & Space(1)
Else
KetQua = Space(0)
End If
SoTien = Format(Abs(baonhieu), "##############0.00")
SoTien = Right(Space(15) & SoTien, 18)
Hang = Array("None", "tr" & ChrW$(259) & "m", "m" & ChrW$(432) & ChrW$(417) & "i", "g" & ChrW$(236) & " " & ChrW$(273) & "_")
Doc = Array("None", "ng" & ChrW$(224) & "n t" & ChrW$(253), "t" & ChrW$(253), "tri" & ChrW$(7879) & "u", "ng" & ChrW$(224) & "n", ChrW$(273) & ChrW$(7891) & "ng", "")
Dem = Array("None", "m" & ChrW$(7897) & "t", "hai", "ba", "b" & ChrW$(7889) & "n", "n" & ChrW$(259) & "m", "s" & ChrW$(225) & "u", "b" & ChrW$(7849) & "y", "t" & ChrW$(225) & "m", "ch" & ChrW$(237) & "n")
For i = 1 To 6
Nhom = Mid(SoTien, i * 3 - 2, 3)
If Nhom <> Space(3) Then
Select Case Nhom
Case "000"
If i = 5 Then
Chu = ChrW$(273) & ChrW$(7891) & "ng" & Space(1)
Else
Chu = Space(0)
End If
Case ".00"
Chu = "ch" & ChrW$(7861) & "n"
Case Else
S1 = Left(Nhom, 1)
S2 = Mid(Nhom, 2, 1)
S3 = Right(Nhom, 1)
Chu = Space(0)
Hang(3) = Doc(i)
For J = 1 To 3
Dich = Space(0)
S = Val(Mid(Nhom, J, 1))
If S > 0 Then
Dich = Dem(S) & Space(1) & Hang(J) & Space(1)
End If
Select Case J
Case 2 And S = 1
Dich = "m" & ChrW$(432) & ChrW$(7901) & "i" & Space(1)
Case 3 And S = 0 And Nhom <> Space(2) & "0"
Dich = Hang(J) & Space(1)
Case 3 And S = 5 And S2 <> Space(1) And S2 <> "0"
Dich = "l" & Mid(Dich, 2)
Case 2 And S = 0 And S3 <> "0"
If (S1 >= "1" And S1 <= "9") Or (S1 = "0" And i = 4) Then
Dich = "l" & ChrW$(7867) & Space(1)
End If
End Select
Chu = Chu & Dich
Next J
End Select
ViTri = InStr(1, Chu, "m" & ChrW$(432) & ChrW$(417) & "i m" & ChrW$(7897) & "t", 1)
If ViTri > 0 Then Mid(Chu, ViTri, 9) = "m" & ChrW$(432) & ChrW$(417) & "i m" & ChrW$(7889) & "t"
KetQua = KetQua & Chu
End If
Next i
End If
End If
VNDUni = UCase(Left(KetQua, 1)) & Mid(KetQua, 2)
End Function

Sub ttt()
Dim Iiii As Integer
' tt Macro
'Application.Range("c2").Value
'Application.Range("J3").Value
If Application.Range("L3").Value <> "" And Application.Range("L4").Value <> "" And Application.Range("L5").Value Then
    For Iiii = Application.Range("L3").Value To Application.Range("L4").Value
        Application.Range("P" & Iiii).Select
        'SendKeys ("{f4}")
        Application.Range("J1").Value = Iiii
        ActiveWindow.SelectedSheets.PrintOut Copies:=Application.Range("L5").Value
    Next
End If
'MsgBox " Phai nhap tu dong... den dong...so trang in..."
If Application.Range("L3").Value = "" Then
    MsgBox " Phai nhap tu dong... "
End If
If Application.Range("L4").Value = "" Then
    MsgBox " Phai nhap den dong... "
End If
If Application.Range("L5").Value = "" Then
    MsgBox " Phai nhap so trang in... "
End If
End Sub

'Sub copyright()
    
    'Sheets("PB_KH").Select
    'Range("A3").Select
    'ActiveCell.FormulaR1C1 = _
        "=+IF(or(MID(CELL(""filename""),IF(ISERROR(FIND(""1TS-KH"",CELL(""filename"")))=TRUE,1,FIND(""1TS-KH"",CELL(""filename""))),6)=""1TS-KH"",MID(CELL(""filename""),IF(ISERROR(FIND(""TS-K"",CELL(""filename"")))=TRUE,1,FIND(""TS-K"",CELL(""filename""))),4)=""TS-K""),0,nam)"
    'If Range("A3") = 2018 Then
    'Range("B3").Select
    'ActiveCell.FormulaR1C1 = "=+DateKC"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=+CDSPS!R418C2"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=+nam*1000*2"
    'Range("G3").Select
    'ActiveCell.FormulaR1C1 = "6428"
    'Range("H3").Select
    'ActiveCell.FormulaR1C1 = "1111"
    'Range("A3: H3").Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    'End If

'End Sub


