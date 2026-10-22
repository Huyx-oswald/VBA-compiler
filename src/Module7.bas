Attribute VB_Name = "Module7"
Sub AutoSetupToolbar()
    Dim cb As Object
    Dim btn As Object

    On Error Resume Next

    Application.CommandBars("FinanceRPA").Delete
    Err.Clear

    Set cb = Application.CommandBars.Add("FinanceRPA", , , True)
    If Err.Number <> 0 Then Exit Sub
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "库存价值表清洗"
    btn.OnAction = "库存价值表数据清洗"
    btn.Style = 3
    btn.FaceId = 210
    btn.Width = 120
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "库龄表清洗"
    btn.OnAction = "清洗库龄表数据"
    btn.Style = 3
    btn.FaceId = 210
    btn.Width = 120
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "库龄期初变化分析"
    btn.OnAction = "库龄表期初变化分析"
    btn.Style = 3
    btn.FaceId = 129
    btn.Width = 120
    btn.BeginGroup = True
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "采购发票清洗"
    btn.OnAction = "采购发票信息数据清洗"
    btn.Style = 3
    btn.FaceId = 210
    btn.Width = 120
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "固定资产隐藏列"
    btn.OnAction = "btnHideFixedAssetColumns"
    btn.Style = 3
    btn.FaceId = 210
    btn.Width = 120
    btn.BeginGroup = True
    Err.Clear

    Set btn = cb.Controls.Add(1)
    btn.Caption = "关于"
    btn.OnAction = "OnBtnAbout"
    btn.Style = 3
    btn.FaceId = 487
    btn.Width = 120
    btn.BeginGroup = True
    Err.Clear

    cb.Visible = True
    On Error GoTo 0
End Sub

Sub OnBtnAbout()
    MsgBox "财务RPA 加载宏 v1.0" & vbCrLf & vbCrLf & _
           "功能模块：" & vbCrLf & _
           "  - 库存价值表清洗" & vbCrLf & _
           "  - 库龄表清洗" & vbCrLf & _
           "  - 库龄期初变化分析" & vbCrLf & _
           "  - 采购发票清洗" & vbCrLf & _
           "  - 固定资产隐藏列" & vbCrLf & vbCrLf & _
           "维护：编辑 src/*.bas → 关闭WPS → 运行 build.vbs", _
           vbInformation, "关于 财务RPA"
End Sub
