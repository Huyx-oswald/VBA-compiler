Attribute VB_Name = "Module4"
' ============================================================
' 库龄表期初变化分析
' 输入：当前活动工作簿的已清洗库龄表（当期，如 2026.05）+ 记住式选定的历史期库龄表（期初，如 2025.12）
' 输出：在当前工作簿新增【分析结果】表，按物料号汇总并对比期初/当期呆滞
' 依赖 ModuleCommon：解析伴生文件 / Nz / 校验表头 / 取工作表
' ============================================================
Sub 库龄表期初变化分析()
    Dim wb2025 As Workbook
    Dim wb2026 As Workbook
    Dim ws2025 As Worksheet
    Dim ws2026 As Worksheet
    Dim wsResult As Worksheet
    Dim filePath2025 As String
    Dim errMsg As String

    On Error GoTo ErrorHandler

    Set wb2026 = ActiveWorkbook
    Set ws2026 = ActiveSheet

    ' 记住式定位历史期库龄表（首次弹框，之后记住；取消则中止）
    filePath2025 = 解析伴生文件("库龄表历史期文件", "2025年12月库龄表.XLSX", wb2026.Path)
    If Len(filePath2025) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 表头快照校验：当期与历史期都应是已清洗库龄表（B分类 C物料号）
    errMsg = 校验表头(ws2026, Array("B", "C"), Array("分类", "物料号"))
    If Len(errMsg) > 0 Then
        MsgBox "当期库龄表表头校验未通过（是否已先做库龄表清洗？）：" & vbCrLf & errMsg, vbCritical, "表头异常"
        GoTo CleanExit
    End If

    Set wb2025 = Workbooks.Open(filePath2025, ReadOnly:=True)
    Set ws2025 = wb2025.Sheets(1)

    errMsg = 校验表头(ws2025, Array("B", "C"), Array("分类", "物料号"))
    If Len(errMsg) > 0 Then
        MsgBox "历史期库龄表表头校验未通过：" & vbCrLf & errMsg, vbCritical, "表头异常"
        GoTo CleanExit
    End If

    ' 重建结果表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb2026.Sheets("分析结果").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Set wsResult = wb2026.Sheets.Add(After:=wb2026.Sheets(wb2026.Sheets.Count))
    wsResult.Name = "分析结果"

    Call 设置标题行(wsResult)

    Dim dict2026 As Object
    Set dict2026 = CreateObject("Scripting.Dictionary")
    Call 汇总数据(ws2026, dict2026)

    Dim dict2025 As Object
    Set dict2025 = CreateObject("Scripting.Dictionary")
    Call 汇总数据(ws2025, dict2025)

    Call 填充结果表(wsResult, dict2026, dict2025)
    Call 按差异金额排序(wsResult)
    Call 格式化结果表(wsResult)

    wsResult.Activate
    wb2025.Close SaveChanges:=False
    wb2026.Save

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "分析完成！结果已生成在当前工作簿的【分析结果】工作表中。", vbInformation, "完成"
    Exit Sub

CleanExit:
    If Not wb2025 Is Nothing Then wb2025.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    If Not wb2025 Is Nothing Then wb2025.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' 设置标题行
Private Sub 设置标题行(ws As Worksheet)
    With ws.Range("A1")
        .Value = "物料号"
        .Offset(0, 1).Value = "分类"
        .Offset(0, 2).Value = "物料描述"
        .Offset(0, 3).Value = "物料组描述"
        .Offset(0, 4).Value = "单位"
        .Offset(0, 5).Value = "当期数量"
        .Offset(0, 6).Value = "当期金额"
        .Offset(0, 7).Value = "当期呆滞数量"
        .Offset(0, 8).Value = "当期呆滞金额"
        .Offset(0, 9).Value = "期初数量"
        .Offset(0, 10).Value = "期初金额"
        .Offset(0, 11).Value = "期初呆滞数量"
        .Offset(0, 12).Value = "期初呆滞金额"
        .Offset(0, 13).Value = "数量差异"
        .Offset(0, 14).Value = "金额差异"
        .Offset(0, 15).Value = "数量差异幅度"
        .Offset(0, 16).Value = "金额差异幅度"
    End With
End Sub

' 按物料号汇总数据
' 性能优化：整块读入内存数组后遍历，避免逐格 COM 往返（2万行原需 ~27万次读取）
Private Sub 汇总数据(ws As Worksheet, dict As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim data As Variant
    Dim materialCode As Variant
    Dim category As String
    Dim materialDesc As String
    Dim unit As String
    Dim qty As Double
    Dim amt As Double
    Dim groupDesc As String
    Dim stagnantQty As Double
    Dim stagnantAmt As Double
    Dim key As String
    Dim arr As Variant

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' A..O = 列 1..15，正好覆盖本过程用到的最右列(O)
    data = ws.Range("A1:O" & lastRow).Value

    For i = 2 To lastRow
        materialCode = data(i, 3)   ' C 物料号
        If materialCode <> "" Then
            key = CStr(materialCode)
            category = CStr(data(i, 2))    ' B 分类
            materialDesc = CStr(data(i, 6))  ' F 物料描述
            unit = CStr(data(i, 11))         ' K 单位
            qty = CDbl(Nz(data(i, 10), 0))   ' J 数量
            amt = CDbl(Nz(data(i, 12), 0))   ' L 金额
            groupDesc = CStr(data(i, 9))     ' I 物料组描述
            stagnantQty = CDbl(Nz(data(i, 14), 0)) ' N 呆滞数量
            stagnantAmt = CDbl(Nz(data(i, 15), 0)) ' O 呆滞金额

            If dict.Exists(key) Then
                arr = dict(key)
                arr(3) = arr(3) + qty
                arr(4) = arr(4) + amt
                arr(6) = arr(6) + stagnantQty
                arr(7) = arr(7) + stagnantAmt
                dict(key) = arr
            Else
                dict.Add key, Array(category, materialDesc, unit, qty, amt, groupDesc, stagnantQty, stagnantAmt)
            End If
        End If
    Next i
End Sub

' 填充结果表
' 性能优化：结果先在内存二维数组中拼好（差异/幅度用内存变量，不再回读单元格），最后一次写回
Private Sub 填充结果表(ws As Worksheet, dict2026 As Object, dict2025 As Object)
    Dim key As Variant
    Dim totalRows As Long
    Dim nOnly2025 As Long
    Dim outRow As Long
    Dim arr2026 As Variant
    Dim arr2025 As Variant
    Dim hasBegin As Boolean
    Dim beginStagQty As Double
    Dim beginStagAmt As Double
    Dim qtyDiff As Double
    Dim amtDiff As Double
    Dim qtyDiffRate As Variant
    Dim amtDiffRate As Variant
    Dim outArr As Variant

    ' 预统计行数：dict2026 全部 + 仅存在于 dict2025 的
    nOnly2025 = 0
    For Each key In dict2025.Keys
        If Not dict2026.Exists(key) Then nOnly2025 = nOnly2025 + 1
    Next key
    totalRows = dict2026.Count + nOnly2025
    If totalRows <= 0 Then Exit Sub

    ReDim outArr(1 To totalRows, 1 To 17)
    outRow = 0

    For Each key In dict2026.Keys
        arr2026 = dict2026(key)
        hasBegin = dict2025.Exists(key)
        If hasBegin Then
            arr2025 = dict2025(key)
            beginStagQty = arr2025(6)
            beginStagAmt = arr2025(7)
        Else
            beginStagQty = 0
            beginStagAmt = 0
        End If

        outRow = outRow + 1
        outArr(outRow, 1) = key
        outArr(outRow, 2) = arr2026(0)
        outArr(outRow, 3) = arr2026(1)
        outArr(outRow, 4) = arr2026(5)
        outArr(outRow, 5) = arr2026(2)
        outArr(outRow, 6) = arr2026(3)
        outArr(outRow, 7) = arr2026(4)
        outArr(outRow, 8) = arr2026(6)
        outArr(outRow, 9) = arr2026(7)
        If hasBegin Then
            outArr(outRow, 10) = arr2025(3)
            outArr(outRow, 11) = arr2025(4)
        End If
        outArr(outRow, 12) = beginStagQty
        outArr(outRow, 13) = beginStagAmt

        qtyDiff = arr2026(6) - beginStagQty
        amtDiff = arr2026(7) - beginStagAmt
        outArr(outRow, 14) = qtyDiff
        outArr(outRow, 15) = amtDiff

        If beginStagQty <> 0 Then
            qtyDiffRate = qtyDiff / beginStagQty
        Else
            qtyDiffRate = IIf(qtyDiff = 0, 0, "#DIV/0!")
        End If
        If beginStagAmt <> 0 Then
            amtDiffRate = amtDiff / beginStagAmt
        Else
            amtDiffRate = IIf(amtDiff = 0, 0, "#DIV/0!")
        End If
        outArr(outRow, 16) = qtyDiffRate
        outArr(outRow, 17) = amtDiffRate
    Next key

    For Each key In dict2025.Keys
        If Not dict2026.Exists(key) Then
            arr2025 = dict2025(key)
            outRow = outRow + 1
            outArr(outRow, 1) = key
            outArr(outRow, 2) = arr2025(0)
            outArr(outRow, 3) = arr2025(1)
            outArr(outRow, 4) = arr2025(5)
            outArr(outRow, 5) = arr2025(2)
            outArr(outRow, 6) = 0
            outArr(outRow, 7) = 0
            outArr(outRow, 8) = 0
            outArr(outRow, 9) = 0
            outArr(outRow, 10) = arr2025(3)
            outArr(outRow, 11) = arr2025(4)
            outArr(outRow, 12) = arr2025(6)
            outArr(outRow, 13) = arr2025(7)
            outArr(outRow, 14) = -arr2025(6)
            outArr(outRow, 15) = -arr2025(7)
            outArr(outRow, 16) = -1
            outArr(outRow, 17) = -1
        End If
    Next key

    ws.Range("A2:Q" & (outRow + 1)).Value = outArr
End Sub

' 按差异金额倒序排序
Private Sub 按差异金额排序(ws As Worksheet)
    Dim lastRow As Long
    Dim dataRange As Range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        Set dataRange = ws.Range("A1:Q" & lastRow)
        dataRange.Sort Key1:=ws.Range("O1"), Order1:=xlDescending, _
                       Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom
    End If
End Sub

' 格式化结果表
Private Sub 格式化结果表(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 17

    With ws.Range("A1:Q1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If lastRow > 1 Then
        ws.Range("F2:F" & lastRow).NumberFormat = "#,##0"
        ws.Range("H2:H" & lastRow).NumberFormat = "#,##0"
        ws.Range("J2:J" & lastRow).NumberFormat = "#,##0"
        ws.Range("L2:L" & lastRow).NumberFormat = "#,##0"
        ws.Range("N2:N" & lastRow).NumberFormat = "#,##0"
        ws.Range("G2:G" & lastRow).NumberFormat = "#,##0.00"
        ws.Range("I2:I" & lastRow).NumberFormat = "#,##0.00"
        ws.Range("K2:K" & lastRow).NumberFormat = "#,##0.00"
        ws.Range("M2:M" & lastRow).NumberFormat = "#,##0.00"
        ws.Range("O2:O" & lastRow).NumberFormat = "#,##0.00"
        ws.Range("P2:Q" & lastRow).NumberFormat = "0.00%"
    End If

    For i = 1 To lastCol
        ws.Columns(i).AutoFit
    Next i

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    If lastRow > 1 Then
        With ws.Range("A1:Q" & lastRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
End Sub
