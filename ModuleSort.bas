Attribute VB_Name = "ModuleSort"
Option Explicit

Sub 口罩特約藥局遞減排序()
Attribute 口罩特約藥局遞減排序.VB_Description = "口罩排序"
Attribute 口罩特約藥局遞減排序.VB_ProcData.VB_Invoke_Func = "a\n14"
' 口罩特約藥局遞減排序 巨集
' 學生請自行打口罩排序
' 快速鍵: Ctrl+q
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
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
Sub 口罩特約藥局遞增()
Attribute 口罩特約藥局遞增.VB_Description = "遞增"
Attribute 口罩特約藥局遞增.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩特約藥局遞增 巨集
' 遞增
'
' 快速鍵: Ctrl+n
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
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
