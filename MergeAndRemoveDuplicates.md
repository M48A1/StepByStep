```
Sub MergeColumnsToDiffere()
    Dim wsERP As Worksheet, wsLegacy As Worksheet, wsDiffere As Worksheet
    Dim lastRowERP As Long, lastRowLegacy As Long, lastRowDiffere As Long
    Dim i As Long

    ' 定义工作表
    Set wsERP = ThisWorkbook.Sheets("ERP") ' 数据来源 sheet ERP
    Set wsLegacy = ThisWorkbook.Sheets("legacy") ' 数据来源 sheet legacy
    Set wsDiffere = ThisWorkbook.Sheets("differe") ' 目标 sheet differe

    ' 获取 ERP 和 legacy 的 B 列最后一行
    lastRowERP = wsERP.Cells(wsERP.Rows.Count, "B").End(xlUp).Row
    lastRowLegacy = wsLegacy.Cells(wsLegacy.Rows.Count, "B").End(xlUp).Row

    ' 将 ERP 的 B 列数据复制到 differe 的 B 列
    For i = 1 To lastRowERP
        wsDiffere.Cells(i, 2).Value = wsERP.Cells(i, 2).Value
    Next i

    ' 将 legacy 的 B 列数据追加到 differe 的 B 列
    For i = 1 To lastRowLegacy
        wsDiffere.Cells(lastRowERP + i, 2).Value = wsLegacy.Cells(i, 2).Value
    Next i

    ' 获取 differe 的 B 列最后一行
    lastRowDiffere = wsDiffere.Cells(wsDiffere.Rows.Count, "B").End(xlUp).Row

    ' 删除 differe 的 B 列中的重复项
    wsDiffere.Range("B1:B" & lastRowDiffere).RemoveDuplicates Columns:=1, Header:=xlNo

    ' 提示操作完成
    MsgBox "数据合并并去重已完成！", vbInformation
End Sub
```
