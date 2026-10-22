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
Private Sub 汇总数据(ws As Worksheet, dict As Object)
    Dim lastRow As Long
    Dim i As Long
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

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = 2 To lastRow
        materialCode = ws.Cells(i, "C").Value
        If materialCode <> "" Then
            key = CStr(materialCode)
            category = CStr(ws.Cells(i, "B").Value)
            materialDesc = CStr(ws.Cells(i, "F").Value)
            unit = CStr(ws.Cells(i, "K").Value)
            qty = CDbl(Nz(ws.Cells(i, "J").Value, 0))
            amt = CDbl(Nz(ws.Cells(i, "L").Value, 0))
            groupDesc = CStr(ws.Cells(i, "I").Value)
            stagnantQty = CDbl(Nz(ws.Cells(i, "N").Value, 0))
            stagnantAmt = CDbl(Nz(ws.Cells(i, "O").Value, 0))

            If dict.Exists(key) Then
                Dim arr As Variant
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
Private Sub 填充结果表(ws As Worksheet, dict2026 As Object, dict2025 As Object)
    Dim key As Variant
    Dim rowNum As Long
    Dim arr2026 As Variant
    Dim arr2025 As Variant
    Dim qtyDiff As Double
    Dim amtDiff As Double
    Dim qtyDiffRate As Variant
    Dim amtDiffRate As Variant

    rowNum = 2

    For Each key In dict2026.Keys
        arr2026 = dict2026(key)
        ws.Cells(rowNum, 1).Value = key
        ws.Cells(rowNum, 2).Value = arr2026(0)
        ws.Cells(rowNum, 3).Value = arr2026(1)
        ws.Cells(rowNum, 4).Value = arr2026(5)
        ws.Cells(rowNum, 5).Value = arr2026(2)
        ws.Cells(rowNum, 6).Value = arr2026(3)
        ws.Cells(rowNum, 7).Value = arr2026(4)
        ws.Cells(rowNum, 8).Value = arr2026(6)
        ws.Cells(rowNum, 9).Value = arr2026(7)

        If dict2025.Exists(key) Then
            arr2025 = dict2025(key)
            ws.Cells(rowNum, 10).Value = arr2025(3)
            ws.Cells(rowNum, 11).Value = arr2025(4)
            ws.Cells(rowNum, 12).Value = arr2025(6)
            ws.Cells(rowNum, 13).Value = arr2025(7)
        Else
            ws.Cells(rowNum, 10).Value = 0
            ws.Cells(rowNum, 11).Value = 0
            ws.Cells(rowNum, 12).Value = 0
            ws.Cells(rowNum, 13).Value = 0
        End If

        qtyDiff = arr2026(6) - ws.Cells(rowNum, 12).Value
        amtDiff = arr2026(7) - ws.Cells(rowNum, 13).Value
        ws.Cells(rowNum, 14).Value = qtyDiff
        ws.Cells(rowNum, 15).Value = amtDiff

        If ws.Cells(rowNum, 12).Value <> 0 Then
            qtyDiffRate = qtyDiff / ws.Cells(rowNum, 12).Value
        Else
            qtyDiffRate = IIf(qtyDiff = 0, 0, "#DIV/0!")
        End If

        If ws.Cells(rowNum, 13).Value <> 0 Then
            amtDiffRate = amtDiff / ws.Cells(rowNum, 13).Value
        Else
            amtDiffRate = IIf(amtDiff = 0, 0, "#DIV/0!")
        End If

        ws.Cells(rowNum, 16).Value = qtyDiffRate
        ws.Cells(rowNum, 17).Value = amtDiffRate
        rowNum = rowNum + 1
    Next key

    For Each key In dict2025.Keys
        If Not dict2026.Exists(key) Then
            arr2025 = dict2025(key)
            ws.Cells(rowNum, 1).Value = key
            ws.Cells(rowNum, 2).Value = arr2025(0)
            ws.Cells(rowNum, 3).Value = arr2025(1)
            ws.Cells(rowNum, 4).Value = arr2025(5)
            ws.Cells(rowNum, 5).Value = arr2025(2)
            ws.Cells(rowNum, 6).Value = 0
            ws.Cells(rowNum, 7).Value = 0
            ws.Cells(rowNum, 8).Value = 0
            ws.Cells(rowNum, 9).Value = 0
            ws.Cells(rowNum, 10).Value = arr2025(3)
            ws.Cells(rowNum, 11).Value = arr2025(4)
            ws.Cells(rowNum, 12).Value = arr2025(6)
            ws.Cells(rowNum, 13).Value = arr2025(7)
            ws.Cells(rowNum, 14).Value = -arr2025(6)
            ws.Cells(rowNum, 15).Value = -arr2025(7)
            ws.Cells(rowNum, 16).Value = -1
            ws.Cells(rowNum, 17).Value = -1
            rowNum = rowNum + 1
        End If
    Next key
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
