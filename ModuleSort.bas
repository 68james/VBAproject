Attribute VB_Name = "ModuleSort"
Option Explicit

Sub �f�n�S���ħ�����Ƨ�()
Attribute �f�n�S���ħ�����Ƨ�.VB_Description = "�f�n�Ƨ�"
Attribute �f�n�S���ħ�����Ƨ�.VB_ProcData.VB_Invoke_Func = "a\n14"
' �f�n�S���ħ�����Ƨ� ����
' �ǥͽЦۦ楴�f�n�Ƨ�
' �ֳt��: Ctrl+q
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("G2").Select
End Sub
Sub �f�n�S���ħ����W()
Attribute �f�n�S���ħ����W.VB_Description = "���W"
Attribute �f�n�S���ħ����W.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�S���ħ����W ����
' ���W
'
' �ֳt��: Ctrl+n
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    Selection.ClearContents
    Range("G1").Select
    Selection.ClearContents
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3],R[413]C[-3])"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("E1").Select
End Sub
