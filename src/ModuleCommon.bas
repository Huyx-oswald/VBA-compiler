Attribute VB_Name = "ModuleCommon"
' ============================================================
' ModuleCommon.bas —— 公共工具集中地
' 说明：
'   - 本模块只放可复用函数，不含被工具栏 OnAction 直接调用的 Sub。
'   - SafeCDbl / Nz 为 Public，其余模块不要再定义同名过程（避免重复定义编译错误）。
'   - 物料分类以首字符为唯一真相（权威映射见 物料大类）；
'     下游归属判定一律用 属于原材料页 / 属于成品页 / 呆滞豁免（按首字符集合），
'     不再用中文标签子串匹配，杜绝“半成品含成品”这类隐性误收。
' ============================================================

' 安全转换为 Double：错误值 / 空 / 非数字 → 0
Public Function SafeCDbl(v As Variant) As Double
    If IsError(v) Then
        SafeCDbl = 0
    ElseIf IsEmpty(v) Then
        SafeCDbl = 0
    ElseIf IsNull(v) Then
        SafeCDbl = 0
    ElseIf IsNumeric(v) Then
        SafeCDbl = CDbl(v)
    Else
        SafeCDbl = 0
    End If
End Function

' 空 / 错误 / 空串兜底
Public Function Nz(v As Variant, Optional defaultValue As Variant = 0) As Variant
    If IsError(v) Then
        Nz = defaultValue
    ElseIf IsEmpty(v) Then
        Nz = defaultValue
    ElseIf IsNull(v) Then
        Nz = defaultValue
    ElseIf v = "" Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

' ---------- 权威物料分类（按物料编码首字符） ----------
Public Function 物料大类(firstChar As String) As String
    Select Case firstChar
        Case "1": 物料大类 = "原材料"
        Case "2": 物料大类 = "半成品"
        Case "3": 物料大类 = "成品"
        Case "4": 物料大类 = "辅料"
        Case "5": 物料大类 = "备品备件"
        Case "6": 物料大类 = "商品"
        Case "7": 物料大类 = "无价值物料"
        Case "9": 物料大类 = "单元板"
        Case ""
            物料大类 = "未知"
        Case Else
            物料大类 = "其他"
    End Select
End Function

' ---------- 下游归属判定（按首字符集合，Q16 定稿） ----------
Public Function 属于原材料页(firstChar As String) As Boolean
    属于原材料页 = (firstChar = "1" Or firstChar = "4" Or firstChar = "5")
End Function

Public Function 属于成品页(firstChar As String) As Boolean
    属于成品页 = (firstChar = "3")
End Function

Public Function 呆滞豁免(firstChar As String) As Boolean
    呆滞豁免 = (firstChar = "4" Or firstChar = "5")
End Function

' ---------- 表头快照校验（Q9-B：不符即中止，列出差异） ----------
' cols / expected 为等长数组：cols 传列字母或列号，expected 传期望表头文字。
' 返回 "" 表示通过；否则返回多行差异描述。
Public Function 校验表头(ws As Worksheet, cols As Variant, expected As Variant) As String
    Dim i As Long, got As String, msg As String
    If UBound(cols) <> UBound(expected) Then
        校验表头 = "校验表头：列数与期望数不一致"
        Exit Function
    End If
    For i = LBound(cols) To UBound(cols)
        got = Trim$(CStr(ws.Cells(1, cols(i)).Value))
        If got <> CStr(expected(i)) Then
            If Len(msg) > 0 Then msg = msg & vbCrLf
            msg = msg & "列 " & CStr(cols(i)) & "：期望「" & CStr(expected(i)) & "」 实际「" & got & "」"
        End If
    Next i
    校验表头 = msg
End Function

' ---------- 记住式伴生文件定位（P1-1，分发到他机不写死路径） ----------
' 逻辑名做注册表键；命中且文件存在→直接用；否则弹 FileDialog 引导，
' 初始目录用 startDir（通常是待清洗表所在目录），选定后回写注册表。
' 用户取消返回 ""（调用方据此中止）。
Public Function 解析伴生文件(logicalName As String, suggestedName As String, startDir As String) As String
    Dim sh As Object, regBase As String, cached As String, picked As String
    regBase = "HKCU\Software\FinanceRPA\Paths\"

    Set sh = CreateObject("WScript.Shell")
    cached = ""
    On Error Resume Next
    cached = sh.RegRead(regBase & logicalName)
    On Error GoTo 0

    If Len(cached) > 0 Then
        If Dir(cached) <> "" Then
            解析伴生文件 = cached
            Exit Function
        End If
    End If

    picked = 选择文件(suggestedName, startDir)
    If Len(picked) = 0 Then
        解析伴生文件 = ""
        Exit Function
    End If

    On Error Resume Next
    sh.RegWrite regBase & logicalName, picked, "REG_SZ"
    On Error GoTo 0

    解析伴生文件 = picked
End Function

' 弹出单选文件框；取消返回 ""
' 使用 Application.GetOpenFilename —— WPS ET / Excel 均支持（FileDialog 在 WPS 不可用）
' 注意：GetOpenFilename 没有 FileName/InitialFileName 参数；起始目录用 ChDir 临时切换。
' 位置参数调用（不用命名参数），最大化 WPS 兼容性。
Public Function 选择文件(suggestedName As String, startDir As String) As String
    Dim guessDir As String, oldDir As String, result As Variant

    guessDir = ""
    If Len(startDir) > 0 Then
        If Dir(startDir, vbDirectory) <> "" Then guessDir = startDir
    End If
    If guessDir = "" Then
        On Error Resume Next
        guessDir = ThisWorkbook.Path
        On Error GoTo 0
    End If

    oldDir = CurDir
    If Len(guessDir) > 0 Then
        On Error Resume Next
        ChDir guessDir
        On Error GoTo 0
    End If

    result = Application.GetOpenFilename( _
        "Excel 文件 (*.xls*),*.xlsx;*.xls;*.xlsm;*.xlam,所有文件 (*.*),*.*", _
        , _
        "请选择伴生文件：" & suggestedName)

    On Error Resume Next
    If Len(oldDir) > 0 Then ChDir oldDir
    On Error GoTo 0

    If VarType(result) = vbBoolean Then   ' 用户取消 → False (Boolean)
        选择文件 = ""
    Else
        选择文件 = CStr(result)
    End If
End Function

' 判断工作表是否存在（不存在返回 Nothing）
Public Function 取工作表(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    Set 取工作表 = ws
End Function
