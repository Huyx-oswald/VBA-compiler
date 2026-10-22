'==========================================================================
' pack.vbs - Build the distribution package (pure VBScript, no Node.js)
'
' Double-click to run.
' Assembles FinanceRPA安装包/ from:
'   - FinanceRPA.xlam            (copied as-is)
'   - templates/install.vbs.txt   (UTF-8) -> install.vbs (GBK, runs on target)
'   - templates/readme.txt        (UTF-8) -> 使用说明.txt  (UTF-8 BOM, human-readable)
'
' All editable text lives in templates/ so there is no giant inline string.
' Prerequisite: run build.vbs first to refresh FinanceRPA.xlam.
'==========================================================================
Option Explicit

Dim fso, scriptDir, xlamPath, packDir, tplDir
Dim resp, stream, txt

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlamPath  = scriptDir & "\FinanceRPA.xlam"
packDir   = scriptDir & "\FinanceRPA安装包"
tplDir    = scriptDir & "\templates"

If Not fso.FileExists(xlamPath) Then
    MsgBox "FinanceRPA.xlam not found!" & vbCrLf & "Please run build.vbs first.", vbCritical, "Error"
    WScript.Quit 1
End If
If Not fso.FolderExists(tplDir) Then
    MsgBox "templates/ folder not found.", vbCritical, "Error"
    WScript.Quit 1
End If

resp = MsgBox("Generate installation package?" & vbCrLf & vbCrLf & _
              "Source: FinanceRPA.xlam + templates/" & vbCrLf & _
              "Output: " & packDir & vbCrLf & vbCrLf & _
              "Click Yes to continue.", vbYesNo + vbQuestion, "Pack")
If resp = vbNo Then WScript.Quit 0

If fso.FolderExists(packDir) Then fso.DeleteFolder packDir, True
fso.CreateFolder packDir

' 1) xlam copy
fso.CopyFile xlamPath, packDir & "\FinanceRPA.xlam", True

' 2) install.vbs : read UTF-8 template -> write GBK (cscript on target reads ANSI/GBK)
WriteReencoded tplDir & "\install.vbs.txt", "utf-8", packDir & "\install.vbs", "gbk"

' 3) 使用说明.txt : read UTF-8 template -> write UTF-8 with BOM (Notepad-friendly)
Dim readmePath
readmePath = packDir & "\使用说明.txt"
WriteReencoded tplDir & "\readme.txt", "utf-8", readmePath, "utf-8"

' done: list files
Dim fileList, fsoFile
fileList = ""
For Each fsoFile In fso.GetFolder(packDir).Files
    fileList = fileList & "  " & fsoFile.Name & " (" & fsoFile.Size & " bytes)" & vbCrLf
Next

MsgBox "Package generated!" & vbCrLf & vbCrLf & _
       "Location: " & packDir & vbCrLf & vbCrLf & _
       "Files:" & vbCrLf & fileList & vbCrLf & _
       "Copy the entire folder to another computer, then double-click install.vbs", _
       vbInformation, "Done"

WScript.Quit 0

' ---------------- helpers ----------------
Sub WriteReencoded(srcPath, srcCharset, dstPath, dstCharset)
    Dim rs, content, ws
    If Not fso.FileExists(srcPath) Then
        MsgBox "Template missing: " & srcPath, vbCritical, "Error"
        WScript.Quit 1
    End If
    Set rs = CreateObject("ADODB.Stream")
    rs.Type = 2
    rs.Charset = srcCharset
    rs.Open
    rs.LoadFromFile srcPath
    content = rs.ReadText
    rs.Close
    Set rs = Nothing

    Set ws = CreateObject("ADODB.Stream")
    ws.Type = 2
    ws.Charset = dstCharset
    ws.Open
    ws.WriteText content
    ws.SaveToFile dstPath, 2
    ws.Close
    Set ws = Nothing
End Sub
