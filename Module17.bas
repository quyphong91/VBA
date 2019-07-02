Attribute VB_Name = "Module17"
Sub SortN()
   
   If Range("nam") = 2018 Then
    Sheets("N").Select
    Range("N_Vsort").Sort Key1:=Range("C12"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "=""PN""&TEXT(thang,""00"")&""-""&TEXT(PN!R6C14,""0000"")"
    Range("L12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC2<>"""",RC2<>R[-1]C2),R[-1]C12+1,R[-1]C12)"
    Range("M12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(RC2<>"""",""PN""&TEXT(thang,""00"")&""-""&TEXT(RC12,""0000""),"""")"
    Range("N12").Select
    ActiveCell.FormulaR1C1 = "=+IF(ISERROR(RC2-1)=FALSE,RC2,RC13)"
    Range("L12:N12").Select
    Selection.Copy
    Application.Goto Reference:="N_Vdanhphieu"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("N_VlocSH1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N10:N11"), Unique:=False
    Application.Goto Reference:="N_VlocSH"
    Selection.ClearContents
    Range("A11:K11").AutoFilter
    Application.Goto Reference:="N_VlocSH"
    Selection.Copy
    Range("B12").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Range("E1").Select
    Range("E8").Select
Exit Sub
  Else
  Range("E8").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If

End Sub
Sub SortX()

If Range("nam") = 2018 Then

    Sheets("X").Select
    Range("X_Vsort").Sort Key1:=Range("C12"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "=""PX""&TEXT(thang,""00"")&""-""&TEXT(PX!R6C14,""0000"")"
    Range("L12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(AND(RC2<>"""",RC2<>R[-1]C2),R[-1]C12+1,R[-1]C12)"
    Range("M12").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(RC2<>"""",""PX""&TEXT(thang,""00"")&""-""&TEXT(RC12,""0000""),"""")"
    Range("N12").Select
    ActiveCell.FormulaR1C1 = "=+IF(ISERROR(RC2-1)=FALSE,RC2,RC13)"
    Range("L12:N12").Select
    Selection.Copy
    Application.Goto Reference:="X_Vdanhphieu"
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("X_VlocSH1").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("N10:N11"), Unique:=False
    Application.Goto Reference:="X_VlocSH"
    Selection.ClearContents
    Range("A11:K11").AutoFilter
    Application.Goto Reference:="X_VlocSH"
    Selection.Copy
    Range("B12").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Range("E1").Select
    Range("E8").Select
Exit Sub
  Else
  Range("E8").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If

End Sub

Sub INAN()

Dim k As Integer
For k = Range("P3").Value To Range("P4").Value
Range("N6").Value = k
If Range("A1") = "XUAT" Then
xuat
Else
nhap
End If
Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(Range("N7").Value - 1, 0)).EntireRow.Hidden = False

ActiveWindow.SelectedSheets.PrintOut
Next k
End Sub
Sub xem()

Dim k As Integer
For k = Range("P3").Value To Range("P4").Value
Range("N6").Value = k
If Range("A1") = "XUAT" Then
xuat
Else
nhap
End If
Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(Range("N7").Value - 1, 0)).EntireRow.Hidden = False

ActiveWindow.SelectedSheets.PrintPreview
Next k

End Sub
Sub xuat()

If Range("nam") = 2018 Then

SortX
    
    Range("X_Vfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("M10:M11"), Unique:=False

    Sheets("PX").Select
    Range("A12:K12").AutoFilter
    Range("12:500").EntireRow.Hidden = False
    Range("A:P").EntireColumn.Hidden = False
    Application.Goto Reference:="PX_data"
    Selection.ClearContents
    
    Range("X!X_DATA1").Copy
    Range("a13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("X!X_DATA2").Copy
    Range("E13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("X!X_DATA3").Copy
    Range("G13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("X!X_DATA4").Copy
    Range("J13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("X!X_DATA5").Copy
    Range("O13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("X!X_Vdanhphieu").Copy
    Range("L13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Range("I13").Select
    ActiveCell.FormulaR1C1 = "=+RC[-1]"
    Range("PX_thucX").Select
    Selection.FillDown
    Range("PX_Vfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("D10:D11"), Unique:=False
    Application.CutCopyMode = False
    Range("A:C").EntireColumn.Hidden = True
    Range("L:N").EntireColumn.Hidden = True
    Sheets("X").Select
    Range("D11:P11").AutoFilter
    Range("E8").Select
    Sheets("PX").Select
    Range("P6").Select

Exit Sub
  Else
    Sheets("PX").Select
  Range("P6").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
   End If
    
    
End Sub
Sub nhap()
   
   If Range("nam") = 2018 Then
    
    SortN
    Range("N_Vfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("M10:M11"), Unique:=False
    
    Sheets("PN").Select
    Range("A12:K12").AutoFilter
    Range("12:500").EntireRow.Hidden = False
    Range("A:P").EntireColumn.Hidden = False
    Application.Goto Reference:="PN_data"
    Selection.ClearContents
    
    Range("N!N_DATA1").Copy
    Range("a13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("N!N_DATA2").Copy
    Range("E13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("N!N_DATA3").Copy
    Range("G13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("N!N_DATA4").Copy
    Range("J13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("N!N_DATA5").Copy
    Range("O13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("N!N_Vdanhphieu").Copy
    Range("L13").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("I13").Select
    ActiveCell.FormulaR1C1 = "=+RC[-1]"
    Range("PN_thucN").Select
    Selection.FillDown
    Range("PN_Vfilter").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("D10:D11"), Unique:=False
    Application.CutCopyMode = False
    Range("A:C").EntireColumn.Hidden = True
    Range("L:N").EntireColumn.Hidden = True
    Sheets("N").Select
    Range("D11:P11").AutoFilter
    Range("E8").Select
    Sheets("PN").Select
    Range("P6").Select

 Exit Sub
  Else

  Range("P6").Select
  MsgBox " No no. So nay chi duoc su dung cho Nam 2018! OK ? "
End If

End Sub



