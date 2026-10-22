'==========================================================================
' build.vbs - Build FinanceRPA.xlam from src/*.bas (WPS/Excel COM automation)
'
' Source-of-truth rules:
'   - src/*.bas are UTF-8 (no BOM). NEVER rewritten by this script.
'   - This script converts each .bas to a TEMP GBK copy under _build_gbk/,
'     imports from the temp copies (WPS VBA import needs GBK), then deletes
'     the temp dir. Git diff of src stays clean.
'   - Build HARD-FAILS (exit 1) if any expected module is missing after import,
'     or the toolbar / ThisWorkbook checks fail, or isAddin patch fails.
'==========================================================================
Option Explicit

Dim fso, app, wb, vbp, srcDir, scriptDir, xlamPath, xlsmPath, gbkDir
Dim startupPath, ai, file
Dim fsoFile, basName

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlamPath  = scriptDir & "\FinanceRPA.xlam"
xlsmPath  = scriptDir & "\FinanceRPA_temp.xlsm"
srcDir    = scriptDir & "\src"
gbkDir    = scriptDir & "\_build_gbk"

If Not fso.FolderExists(srcDir) Then
    WScript.Echo "ERROR: src/ not found"
    WScript.Quit 1
End If

WScript.Echo "=== Build FinanceRPA (UTF-8 source -> xlam) ==="

' --- Step 0: create WPS/Excel app ---
On Error Resume Next
Set app = CreateObject("Ket.Application")
If app Is Nothing Then Set app = CreateObject("ET.Application")
If app Is Nothing Then Set app = CreateObject("Excel.Application")
On Error GoTo 0
If app Is Nothing Then
    WScript.Echo "ERROR: Cannot create WPS/Excel - close all WPS windows first"
    WScript.Quit 1
End If
app.Visible = False
app.DisplayAlerts = False

' --- Step 1: uninstall existing add-in (release file lock) ---
On Error Resume Next
For Each ai In app.AddIns
    If InStr(ai.Name, "FinanceRPA") > 0 Then
        If ai.Installed Then ai.Installed = False
    End If
Next
On Error GoTo 0

' --- Step 2: new workbook, keep a single sheet ---
Set wb = app.Workbooks.Add()
On Error Resume Next
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
On Error GoTo 0

' --- Step 3: VBProject access ---
On Error Resume Next
Set vbp = wb.VBProject
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot access VBA project - enable 'Trust access to the VBA project object model'"
    wb.Close False
    app.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- Step 4: remove all non-ThisWorkbook components (sheets/modules/userforms) ---
Dim comps, compSnapshot(), n, i
comps = vbp.VBComponents.Count
ReDim compSnapshot(comps - 1)
For i = 1 To comps
    Set compSnapshot(i - 1) = vbp.VBComponents.Item(i)
Next
For i = comps - 1 To 0 Step -1
    If LCase(compSnapshot(i).Name) <> "thisworkbook" Then
        On Error Resume Next
        vbp.VBComponents.Remove compSnapshot(i)
        On Error GoTo 0
    End If
Next

' --- Step 5: reset ThisWorkbook startup stub ---
Dim twComp, twCode
Set twComp = vbp.VBComponents("ThisWorkbook")
Set twCode = twComp.CodeModule
If twCode.CountOfLines > 0 Then twCode.DeleteLines 1, twCode.CountOfLines
twCode.AddFromString "Private Sub Workbook_Open()" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    AutoSetupToolbar" & vbCrLf & _
    "    On Error GoTo 0" & vbCrLf & _
    "End Sub"

' --- Step 6: convert src/*.bas (UTF-8) -> temp GBK copies, import from temp ---
If fso.FolderExists(gbkDir) Then fso.DeleteFolder gbkDir, True
fso.CreateFolder gbkDir

Dim expected()      ' module names that MUST end up in the project
Dim expectedCnt
expectedCnt = 0

Dim rStream, wStream, txt, gbkPath
For Each fsoFile In fso.GetFolder(srcDir).Files
    If LCase(fso.GetExtensionName(fsoFile.Name)) = "bas" Then
        ' read UTF-8
        Set rStream = CreateObject("ADODB.Stream")
        rStream.Type = 2
        rStream.Charset = "utf-8"
        rStream.Open
        rStream.LoadFromFile fsoFile.Path
        txt = rStream.ReadText
        rStream.Close
        Set rStream = Nothing

        ' write GBK temp copy
        gbkPath = gbkDir & "\" & fsoFile.Name
        Set wStream = CreateObject("ADODB.Stream")
        wStream.Type = 2
        wStream.Charset = "gbk"
        wStream.Open
        wStream.WriteText txt
        wStream.SaveToFile gbkPath, 2
        wStream.Close
        Set wStream = Nothing

        ' record expected module name (from Attribute VB_Name, fallback to file stem)
        ReDim Preserve expected(expectedCnt)
        expected(expectedCnt) = ModuleNameOf(txt, fso.GetBaseName(fsoFile.Name))
        expectedCnt = expectedCnt + 1

        ' import (hard-fail later via presence check; surface error here too)
        On Error Resume Next
        vbp.VBComponents.Import gbkPath
        If Err.Number <> 0 Then
            WScript.Echo "  IMPORT ERROR: " & fsoFile.Name & " -> " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
Next
WScript.Echo "Step: converted+imported " & expectedCnt & " module(s) from UTF-8 temp copies"

' --- Step 7: save as xlsm (persist VBA), then SaveAs xlam (format 55) ---
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True
On Error Resume Next
wb.SaveAs xlsmPath, 52
If Err.Number <> 0 Then
    WScript.Echo "ERROR: SaveAs xlsm failed: " & Err.Description
    wb.Close False
    app.Quit
    cleanupGbk
    WScript.Quit 1
End If
On Error GoTo 0

If fso.FileExists(xlamPath) Then fso.DeleteFile xlamPath, True
On Error Resume Next
wb.IsAddin = True
On Error GoTo 0

Dim saveAsOk, fmtOk
saveAsOk = False
On Error Resume Next
wb.SaveAs xlamPath, 55
If Err.Number = 0 Then
    saveAsOk = True
Else
    Err.Clear
    wb.SaveAs xlamPath, 18
    If Err.Number = 0 Then saveAsOk = True
End If
On Error GoTo 0
If Not saveAsOk Then
    WScript.Echo "ERROR: Cannot SaveAs xlam"
    wb.Close False
    app.Quit
    cleanupGbk
    WScript.Quit 1
End If
wb.Close False
WScript.Sleep 800

' --- Step 8: patch isAddin into workbook.xml (WPS SaveAs omits it) + HARD-FAIL ---
Dim shell, retCode
Set shell = CreateObject("WScript.Shell")
retCode = shell.Run("powershell -NoProfile -ExecutionPolicy Bypass -File """ & scriptDir & "\fix_isaddin.ps1"" """ & xlamPath & """", 0, True)
If retCode <> 0 Then
    WScript.Echo "ERROR: fix_isaddin.ps1 failed with exit code " & retCode
    app.Quit
    cleanupGbk
    WScript.Quit 1
End If
WScript.Echo "Step: isAddin patched OK (exit " & retCode & ")"

' --- Step 9: VERIFY against src manifest + toolbar structural checks; HARD-FAIL ---
Dim verifyFail, missing, vbp2, comp, foundToolbar, foundCommon
verifyFail = False
missing = ""

Set wb = app.Workbooks.Open(xlamPath)
Set vbp2 = wb.VBProject
For i = 0 To expectedCnt - 1
    Dim present
    present = False
    For Each comp In vbp2.VBComponents
        If LCase(comp.Name) = LCase(expected(i)) Then present = True
    Next
    If Not present Then
        verifyFail = True
        missing = missing & vbCrLf & "  MISSING module: " & expected(i)
    End If
Next

foundToolbar = ComponentCodeContains(vbp2, "Sub AutoSetupToolbar")
foundCommon  = ComponentCodeContains(vbp2, "SafeCDbl")
If Not foundToolbar Then
    verifyFail = True
    missing = missing & vbCrLf & "  MISSING: AutoSetupToolbar (toolbar module not loaded)"
End If
If Not foundCommon Then
    verifyFail = True
    missing = missing & vbCrLf & "  MISSING: ModuleCommon helpers (SafeCDbl)"
End If

wb.Close False

If verifyFail Then
    WScript.Echo "VERIFY FAILED:" & missing
    app.Quit
    cleanupGbk
    WScript.Quit 1
End If
WScript.Echo "Step: verification passed (" & expectedCnt & " modules + toolbar + common)"

' --- Step 10: deploy to startup folder ---
On Error Resume Next
For Each ai In app.AddIns
    If InStr(ai.Name, "FinanceRPA") > 0 Then ai.Installed = True
Next
On Error GoTo 0
startupPath = app.StartupPath
app.Quit
WScript.Sleep 1500

If Len(startupPath) > 0 And fso.FolderExists(startupPath) Then
    If fso.FileExists(startupPath & "\FinanceRPA.xlam") Then
        fso.DeleteFile startupPath & "\FinanceRPA.xlam", True
    End If
    fso.CopyFile xlamPath, startupPath & "\FinanceRPA.xlam", True
    WScript.Echo "Deployed to startup: " & startupPath
End If

' --- cleanup temp files ---
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True
cleanupGbk

WScript.Echo ""
WScript.Echo "=== DONE - ready to test ==="
WScript.Quit 0

' ================= helpers =================
Function ModuleNameOf(codeText, fallback)
    Dim lines, j, ln, p, namePart
    codeText = Replace(codeText, vbCr, "")   ' normalize CRLF: strip trailing CR so module names match
    lines = Split(codeText, vbLf)
    For j = 0 To UBound(lines)
        ln = lines(j)
        If InStr(ln, "Attribute VB_Name") > 0 Then
            p = InStr(ln, "=")
            If p > 0 Then
                namePart = Trim(Mid(ln, p + 1))
                namePart = Replace(namePart, """", "")
                ModuleNameOf = namePart
                Exit Function
            End If
        End If
    Next
    ModuleNameOf = fallback
End Function

Function ComponentCodeContains(vbProj, needle)
    Dim c, cm, k
    ComponentCodeContains = False
    For Each c In vbProj.VBComponents
        Set cm = c.CodeModule
        If cm.CountOfLines > 0 Then
            If InStr(cm.Lines(1, cm.CountOfLines), needle) > 0 Then
                ComponentCodeContains = True
                Exit Function
            End If
        End If
    Next
End Function

Sub cleanupGbk
    On Error Resume Next
    If fso.FolderExists(gbkDir) Then fso.DeleteFolder gbkDir, True
    On Error GoTo 0
End Sub
