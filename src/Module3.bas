Attribute VB_Name = "Module3"
' ============================================================
' 清洗库龄表数据
' 输入：当前活动工作表的原始库龄表（A工厂 B物料号 …，约30列，未含分类/呆滞列）
' 输出：插入「分类」列(B)与「呆滞数量/呆滞金额」列，呆滞 = 账龄≥180天分段之和
' 依赖 ModuleCommon：物料大类 / 呆滞豁免 / SafeCDbl / 校验表头
' ============================================================
Sub 清洗库龄表数据()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stage As String
    Dim firstChar As String

    Set ws = ActiveSheet
    On Error GoTo ErrorHandler

    ' 幂等守卫：已处理过则中止
    If ws.Cells(1, 2).Value = "分类" Or ws.Cells(1, 14).Value = "呆滞数量" Then
        MsgBox "该工作表已处理过，请使用原始库龄表数据！", vbExclamation, "提示"
        Exit Sub
    End If

    ' 列结构校验
    Dim colCount As Long
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If colCount < 30 Then
        MsgBox "列数不足！当前 " & colCount & " 列，预期至少 30 列。" & vbCrLf & _
               "请确认是否为标准库龄表数据。", vbCritical, "列结构异常"
        Exit Sub
    End If

    ' 表头快照校验（原始库龄表：A工厂 B物料号）
    Dim errMsg As String
    errMsg = 校验表头(ws, Array("A", "B"), Array("工厂", "物料号"))
    If Len(errMsg) > 0 Then
        MsgBox "表头校验未通过，已中止：" & vbCrLf & errMsg, vbCritical, "表头异常"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "未检测到数据行！", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 1. 在B列前插入分类列（原物料号 → C列）
    stage = "插入分类列"
    ws.Columns("B:B").Insert Shift:=xlToRight

    ' 2. 按首字符填分类（原B列现在是C列）
    stage = "填充分类"
    For i = 2 To lastRow
        firstChar = Left$(CStr(ws.Cells(i, 3).Value), 1)
        ws.Cells(i, 2).Value = 物料大类(firstChar)
    Next i
    ws.Cells(1, 2).Value = "分类"

    ' 3. 在N列前插入两列：呆滞数量、呆滞金额
    stage = "插入呆滞列"
    ws.Columns("N:N").Insert Shift:=xlToRight
    ws.Columns("N:N").Insert Shift:=xlToRight
    ws.Cells(1, 14).Value = "呆滞数量"
    ws.Cells(1, 15).Value = "呆滞金额"

    ' 4. 设置分段列标题（两次插入后账龄段整体右移到 16~33）
    stage = "更新分段标题"
    ws.Cells(1, 16).Value = "0-60数量"
    ws.Cells(1, 17).Value = "0-60金额"
    ws.Cells(1, 18).Value = "60-90数量"
    ws.Cells(1, 19).Value = "60-90金额"
    ws.Cells(1, 20).Value = "90-120数量"
    ws.Cells(1, 21).Value = "90-120金额"
    ws.Cells(1, 22).Value = "120-150数量"
    ws.Cells(1, 23).Value = "120-150金额"
    ws.Cells(1, 24).Value = "150-180数量"
    ws.Cells(1, 25).Value = "150-180金额"
    ws.Cells(1, 26).Value = "180-360数量"
    ws.Cells(1, 27).Value = "180-360金额"
    ws.Cells(1, 28).Value = "360-720数量"
    ws.Cells(1, 29).Value = "360-720金额"
    ws.Cells(1, 30).Value = "720-1080数量"
    ws.Cells(1, 31).Value = "720-1080金额"
    ws.Cells(1, 32).Value = "1080以上数量"
    ws.Cells(1, 33).Value = "1080以上金额"

    ' 5. 判定呆滞物料（账龄≥180天的四段之和；半成品-镍版 仅取1080以上段；辅料/备品备件豁免）
    stage = "判定呆滞物料"
    Dim sGroupDesc As String
    Dim seg6Qty As Double, seg6Amt As Double
    Dim seg7Qty As Double, seg7Amt As Double
    Dim seg8Qty As Double, seg8Amt As Double
    Dim seg9Qty As Double, seg9Amt As Double
    Dim totalQty As Double, totalAmt As Double
    ' 性能优化：整块读入 A..AG(列33) 内存数组后遍历，结果先在内存二维数组拼好再一次性写回
    ' （2万行原约 24万次逐格 COM 往返 -> 1 次整块读 + 1 次整块写）
    Dim data As Variant
    Dim outArr As Variant
    data = ws.Range("A1:AG" & lastRow).Value
    ReDim outArr(2 To lastRow, 1 To 2)
    For i = 2 To lastRow
        sGroupDesc = CStr(data(i, 9))   ' 物料组描述
        firstChar = Left$(CStr(data(i, 3)), 1)

        seg6Qty = SafeCDbl(data(i, 26))
        seg6Amt = SafeCDbl(data(i, 27))
        seg7Qty = SafeCDbl(data(i, 28))
        seg7Amt = SafeCDbl(data(i, 29))
        seg8Qty = SafeCDbl(data(i, 30))
        seg8Amt = SafeCDbl(data(i, 31))
        seg9Qty = SafeCDbl(data(i, 32))
        seg9Amt = SafeCDbl(data(i, 33))

        ' 严格保留原语义：仅当对应合计 >0 才写入，否则写 0
        If 呆滞豁免(firstChar) Then
            totalQty = 0: totalAmt = 0
        ElseIf InStr(sGroupDesc, "半成品-镍版") > 0 Then
            If seg9Qty > 0 Or seg9Amt > 0 Then
                totalQty = seg9Qty: totalAmt = seg9Amt
            Else
                totalQty = 0: totalAmt = 0
            End If
        Else
            totalQty = seg6Qty + seg7Qty + seg8Qty + seg9Qty
            totalAmt = seg6Amt + seg7Amt + seg8Amt + seg9Amt
            If Not (totalQty > 0 Or totalAmt > 0) Then
                totalQty = 0: totalAmt = 0
            End If
        End If
        outArr(i, 1) = totalQty
        outArr(i, 2) = totalAmt
    Next i

    ' 呆滞数量=第14列(N)、呆滞金额=第15列(O)，整块写回
    ws.Range("N2:O" & lastRow).Value = outArr

    Application.ScreenUpdating = True
    MsgBox "库龄表数据清洗完成！共处理 " & (lastRow - 1) & " 行数据。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清洗失败！阶段: " & stage & vbCrLf & _
           "错误行: " & i & vbCrLf & _
           "错误: " & Err.Description & " (代码 " & Err.Number & ")", _
           vbCritical, "清洗失败"
End Sub
