Attribute VB_Name = "Module2"
' ============================================================
' 库存价值表数据清洗
' 输入：当前活动工作簿的 Sheet1（库存价值表 SAP 导出，14 列）
' 伴生文件（记住式定位，见 ModuleCommon.解析伴生文件）：
'   - 物料组模板：col1=物料组编码 col2=物料组文本（无表头，首行即数据）
'   - 期初库存价值：col2=物料 col11=总计库存价值
' 处理：新增「物料组文本」「期初库存价值」「分类」三列，并派生原材料页/成品页
' ============================================================
Sub 库存价值表数据清洗()
    Dim wb As Workbook
    Dim dataWS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim errMsg As String
    Dim firstChar As String
    Dim response As VbMsgBoxResult

    Set wb = ActiveWorkbook

    ' 确认处理的工作簿
    response = MsgBox("即将处理工作簿: " & wb.Name & vbCrLf & "是否继续？", vbYesNo + vbQuestion, "确认工作簿")
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If

    ' 定位主数据表
    Set dataWS = 取工作表(wb, "Sheet1")
    If dataWS Is Nothing Then
        MsgBox "错误：在工作簿 '" & wb.Name & "' 中找不到名为 'Sheet1' 的工作表。", vbCritical, "缺工作表"
        Exit Sub
    End If

    ' 幂等守卫：若已含派生列表头，说明处理过，中止以免二次插列错位
    If 已处理过(dataWS) Then
        MsgBox "该表已清洗过（检测到派生列），请使用原始库存价值表数据后再运行！", vbExclamation, "已处理"
        Exit Sub
    End If

    ' 表头快照校验（在插列前，按原始列位置锚定，不符即中止）
    errMsg = 校验表头(dataWS, Array("B", "D"), Array("物料", "物料组"))
    If Len(errMsg) > 0 Then
        MsgBox "表头校验未通过，已中止（避免读错列）：" & vbCrLf & errMsg, vbCritical, "表头异常"
        Exit Sub
    End If

    ' lastRow 以物料编码主列(B)为唯一基准，只求一次，后续各段复用
    lastRow = dataWS.Cells(dataWS.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "错误：工作表 '" & dataWS.Name & "' 中没有找到足够的数据。", vbExclamation, "无数据"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ---------- 载入物料组模板 ----------
    Set dict = CreateObject("Scripting.Dictionary")
    Dim templateFilePath As String
    templateFilePath = 解析伴生文件("物料组模板", "物料组模板.xlsx", wb.Path)
    If Len(templateFilePath) = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim templateWB As Workbook
    On Error Resume Next
    Set templateWB = Workbooks.Open(templateFilePath, ReadOnly:=True)
    On Error GoTo 0
    If templateWB Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "错误：无法打开物料组模板文件，请检查是否被占用或损坏。" & vbCrLf & templateFilePath, vbCritical, "打开失败"
        Exit Sub
    End If

    Dim templateWS As Worksheet, templateLastRow As Long
    Set templateWS = templateWB.Sheets(1)
    templateLastRow = templateWS.Cells(templateWS.Rows.Count, 1).End(xlUp).Row
    ' 模板无表头，首行即数据 → 从第 1 行读（旧代码从第 2 行起会漏掉首条映射）
    For i = 1 To templateLastRow
        If Not IsEmpty(templateWS.Cells(i, 1)) And Not IsEmpty(templateWS.Cells(i, 2)) Then
            dict(CStr(templateWS.Cells(i, 1).Value)) = CStr(templateWS.Cells(i, 2).Value)
        End If
    Next i
    templateWB.Close SaveChanges:=False

    ' ---------- 1. 新增物料组文本列（E列，物料组后）----------
    dataWS.Cells(1, 5).EntireColumn.Insert
    dataWS.Cells(1, 5) = "物料组文本"
    Dim materialGroupCode As String, foundMatch As Boolean, key As Variant
    For i = 2 To lastRow
        materialGroupCode = CStr(dataWS.Cells(i, 4).Value)
        If dict.Exists(materialGroupCode) Then
            dataWS.Cells(i, 5) = dict(materialGroupCode)
        Else
            foundMatch = False
            For Each key In dict.keys
                If Len(CStr(key)) > 0 Then
                    If Left(materialGroupCode, Len(CStr(key))) = CStr(key) Then
                        dataWS.Cells(i, 5) = dict(key)
                        foundMatch = True
                        Exit For
                    End If
                End If
            Next key
            If Not foundMatch Then dataWS.Cells(i, 5) = "未找到匹配"
        End If
    Next i

    ' ---------- 2. 新增期初库存价值列 ----------
    dataWS.Cells(1, 12).EntireColumn.Insert
    dataWS.Cells(1, 12) = "期初库存价值"

    Dim prevDataFilePath As String
    prevDataFilePath = 解析伴生文件("期初库存价值", "库存价值表25年12月EXPORT.XLSX", wb.Path)
    If Len(prevDataFilePath) = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim prevDataWB As Workbook
    On Error Resume Next
    Set prevDataWB = Workbooks.Open(prevDataFilePath, ReadOnly:=True)
    On Error GoTo 0
    If prevDataWB Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "错误：无法打开期初数据文件。" & vbCrLf & prevDataFilePath, vbCritical, "打开失败"
        Exit Sub
    End If

    Dim prevDataWS As Worksheet, prevLastRow As Long
    Set prevDataWS = prevDataWB.Sheets(1)
    prevLastRow = prevDataWS.Cells(prevDataWS.Rows.Count, "B").End(xlUp).Row

    Dim prevDataDict As Object
    Set prevDataDict = CreateObject("Scripting.Dictionary")
    For i = 2 To prevLastRow
        prevDataDict(CStr(prevDataWS.Cells(i, 2).Value)) = SafeCDbl(prevDataWS.Cells(i, 11).Value)
    Next i
    prevDataWB.Close SaveChanges:=False

    ' col2 仍是物料编码（前面只在 E、L 处插列，未触及 A~D）
    Dim currentMaterialNo As String
    For i = 2 To lastRow
        currentMaterialNo = CStr(dataWS.Cells(i, 2).Value)
        If prevDataDict.Exists(currentMaterialNo) Then
            dataWS.Cells(i, 12) = prevDataDict(currentMaterialNo)
        Else
            dataWS.Cells(i, 12) = 0
        End If
    Next i

    ' ---------- 3. 新增分类列（B列，物料编码前）----------
    dataWS.Cells(1, 2).EntireColumn.Insert
    dataWS.Cells(1, 2) = "分类"
    ' 此时：col2=分类, col3=物料编码, col6=物料组文本
    For i = 2 To lastRow
        firstChar = Left$(CStr(dataWS.Cells(i, 3).Value), 1)
        dataWS.Cells(i, 2) = 物料大类(firstChar)
    Next i

    ' ---------- 4. 派生原材料页 ----------
    Dim rawMatWS As Worksheet
    Set rawMatWS = 取工作表(wb, "原材料")
    If rawMatWS Is Nothing Then
        Set rawMatWS = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        rawMatWS.Name = "原材料"
    Else
        rawMatWS.Cells.Clear
    End If
    rawMatWS.Range("A1").Resize(1, 17).Value = dataWS.Range("A1").Resize(1, 17).Value
    Dim rawMatRow As Long: rawMatRow = 2
    For i = 2 To lastRow
        firstChar = Left$(CStr(dataWS.Cells(i, 3).Value), 1)
        If 属于原材料页(firstChar) Then
            rawMatWS.Range("A" & rawMatRow & ":Q" & rawMatRow).Value = dataWS.Range("A" & i & ":Q" & i).Value
            rawMatRow = rawMatRow + 1
        End If
    Next i

    ' ---------- 5. 派生成品页（代码=3 且 物料组文本命中 电化铝/纸品/膜品）----------
    Dim finishedProdWS As Worksheet
    Set finishedProdWS = 取工作表(wb, "成品")
    If finishedProdWS Is Nothing Then
        Set finishedProdWS = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        finishedProdWS.Name = "成品"
    Else
        finishedProdWS.Cells.Clear
    End If
    finishedProdWS.Range("A1").Resize(1, 17).Value = dataWS.Range("A1").Resize(1, 17).Value
    Dim finishedProdRow As Long: finishedProdRow = 2
    Dim groupText As String
    For i = 2 To lastRow
        firstChar = Left$(CStr(dataWS.Cells(i, 3).Value), 1)
        groupText = CStr(dataWS.Cells(i, 6).Value)
        If 属于成品页(firstChar) And _
           (InStr(groupText, "电化铝") > 0 Or InStr(groupText, "纸品") > 0 Or InStr(groupText, "膜品") > 0) Then
            finishedProdWS.Range("A" & finishedProdRow & ":Q" & finishedProdRow).Value = dataWS.Range("A" & i & ":Q" & i).Value
            finishedProdRow = finishedProdRow + 1
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "数据清洗完成！" & vbCrLf & _
           "处理的工作簿: " & wb.Name & vbCrLf & _
           "处理的工作表: " & dataWS.Name & vbCrLf & _
           "数据行数: " & (lastRow - 1), vbInformation, "完成"
End Sub

' 幂等判定：行1 是否已存在派生列表头
Private Function 已处理过(ws As Worksheet) As Boolean
    Dim lastCol As Long, c As Long, h As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If h = "物料组文本" Or h = "期初库存价值" Or h = "分类" Then
            已处理过 = True
            Exit Function
        End If
    Next c
    已处理过 = False
End Function
