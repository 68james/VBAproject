Attribute VB_Name = "ModuleAscSort"
Option Explicit

Sub ascSort()
Attribute ascSort.VB_Description = "遞增"
Attribute ascSort.VB_ProcData.VB_Invoke_Func = "q\n14"
'' ascSort 巨集
' 遞增
'' 快速鍵: Ctrl+q
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
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("E2").Select
End Sub
