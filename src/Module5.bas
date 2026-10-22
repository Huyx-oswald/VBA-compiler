Attribute VB_Name = "Module5"
' ============================================================
' 采购发票信息数据清洗
' 1. 收集特征集合：H列="H" 且 W列="否" 的行，取其 AK列非空值
' 2. 删除所有 AK列值命中特征集合的行
' 3. 对剩余 W列="是" 的行，将 Q列金额取为负数
'
' 列对应（基于实际发票导出文件）：
'   H列 = 凭证类型(D/C)   W列 = 退货项目   Q列 = 金额   AK列 = 参考
'
' BUG-3（防误删）：第2步删除前弹预览确认，并额外统计“命中特征但 W<>'否'”的行数
'   作为风险提示（这些行原本是可能不该删的正常行）。用户取消则不删除。
' ============================================================
Sub 采购发票信息数据清洗()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有数据行，请检查！", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- 第一步：收集特征集合 ---
    Dim featureDict As Object
    Set featureDict = CreateObject("Scripting.Dictionary")
    Dim i As Long, akVal As String
    For i = 2 To lastRow
        If UCase(Trim$(CStr(ws.Cells(i, "H").Value))) = "H" And _
           Trim$(CStr(ws.Cells(i, "W").Value)) = "否" Then
            akVal = Trim$(CStr(ws.Cells(i, "AK").Value))
            If akVal <> "" And Not featureDict.Exists(akVal) Then
                featureDict.Add akVal, True
            End If
        End If
    Next i
    Dim step1Count As Long
    step1Count = featureDict.Count

    ' --- 预扫描：将要删除的行数 + 风险行数（命中特征但 W<>"否"）---
    Dim willDelete As Long, riskRows As Long
    willDelete = 0: riskRows = 0
    For i = 2 To lastRow
        akVal = Trim$(CStr(ws.Cells(i, "AK").Value))
        If akVal <> "" And featureDict.Exists(akVal) Then
            willDelete = willDelete + 1
            If Trim$(CStr(ws.Cells(i, "W").Value)) <> "否" Then riskRows = riskRows + 1
        End If
    Next i

    ' --- 第二步：删除（先弹预览确认）---
    Dim deleteCount As Long
    deleteCount = 0
    If willDelete > 0 Then
        Dim confirm As VbMsgBoxResult
        confirm = MsgBox("即将按参考号(AK)删除 " & willDelete & " 行。" & vbCrLf & _
                         "其中 " & riskRows & " 行的「退货项目」不为“否”（可能是正常行）。" & vbCrLf & vbCrLf & _
                         "确认删除？", vbYesNo + vbExclamation, "删除预览")
        If confirm = vbYes Then
            For i = lastRow To 2 Step -1
                akVal = Trim$(CStr(ws.Cells(i, "AK").Value))
                If akVal <> "" And featureDict.Exists(akVal) Then
                    ws.Rows(i).Delete
                    deleteCount = deleteCount + 1
                End If
            Next i
        Else
            deleteCount = 0
        End If
    End If

    ' --- 第三步：W="是" 的行，Q 金额取负 ---
    Dim negCount As Long
    negCount = 0
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, "W").Value)) = "是" Then
            If IsNumeric(ws.Cells(i, "Q").Value) And ws.Cells(i, "Q").Value <> "" Then
                If CDbl(ws.Cells(i, "Q").Value) >= 0 Then
                    ws.Cells(i, "Q").Value = -CDbl(ws.Cells(i, "Q").Value)
                    negCount = negCount + 1
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "处理完成！" & vbCrLf & _
           "Step1 收集特征数：" & step1Count & vbCrLf & _
           "Step2 删除行数：" & deleteCount & vbCrLf & _
           "Step3 Q列取反行数：" & negCount, vbInformation, "执行结果"

    Set featureDict = Nothing
End Sub
