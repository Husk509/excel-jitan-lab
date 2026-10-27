Option Explicit

Dim fso, shell, appDataPath, targetFolder
Dim scriptDir, xlamSrc, jsonSrc, xlamDst, jsonDst
Dim excelApp, addinObj, addinName

' オブジェクトの初期化
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' 実行中のスクリプトが存在するフォルダパスを取得
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 転送元ファイルのパス
xlamSrc = fso.BuildPath(scriptDir, "CustomAddin.xlam")
jsonSrc = fso.BuildPath(scriptDir, "config.json")

' 転送先の Excel アドイン標準フォルダ (%AppData%\Microsoft\AddIns)
appDataPath = shell.ExpandEnvironmentStrings("%APPDATA%")
targetFolder = fso.BuildPath(appDataPath, "Microsoft\AddIns")

' アドインフォルダが存在しない場合は作成
If Not fso.FolderExists(targetFolder) Then
    fso.CreateFolder(targetFolder)
End If

' 転送先ファイルのパス
xlamDst = fso.BuildPath(targetFolder, "CustomAddin.xlam")
jsonDst = fso.BuildPath(targetFolder, "config.json")

' 1. ファイルのコピー処理
On Error Resume Next
If fso.FileExists(xlamSrc) Then
    fso.CopyFile xlamSrc, xlamDst, True
    If Err.Number <> 0 Then
        MsgBox "アドインのコピーに失敗しました。" & vbCrLf & "Excel をすべて閉じてから、もう一度、実行してください。" & vbCrLf & "(" & Err.Description & ")", vbExclamation, "ご注意"
        WScript.Quit
    End If
Else
    MsgBox "CustomAddin.xlam が見つかりません。" & vbCrLf & "同一フォルダにファイルが存在するか確認してください。", vbCritical, "エラー"
    WScript.Quit
End If

If fso.FileExists(jsonSrc) Then
    fso.CopyFile jsonSrc, jsonDst, True
End If
On Error GoTo 0

' 1-b. ダウンロードの印(Mark of the Web)を外す。外さないと、Excel がマクロを止める
UnblockFile xlamDst
If fso.FileExists(jsonDst) Then UnblockFile jsonDst

' 2. Excel へのアドイン登録＆有効化処理
On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False

If Err.Number = 0 Then
    ' 既存のアドインから同名ものを取得、なければ新規追加
    Dim item, targetAddin
    For Each item In excelApp.AddIns
        If UCase(item.Name) = "CUSTOMADDIN.XLAM" Then
            Set targetAddin = item
            Exit For
        End If
    Next

    If targetAddin Is Nothing Then
        Set targetAddin = excelApp.AddIns.Add(xlamDst, True)
    End If

    ' 強制的に有効化
    targetAddin.Installed = True
    
    excelApp.Quit
    Set excelApp = Nothing
    
    MsgBox "Excel アドインのインストールと有効化が完了しました！" & vbCrLf & vbCrLf & _
           "配置先: " & targetFolder & vbCrLf & vbCrLf & _
           "Excel を開いて、「便利ツール」タブを見てください。出ないときは、「はじめにお読みください」を見てください。", vbInformation, "インストール完了"
Else
    MsgBox "Excel の自動操作に失敗しました。Excelを完全に閉じた状態で再試行してください。", vbExclamation, "ご注意"
End If
On Error GoTo 0

' ファイルの「ダウンロードの印」を外す(PowerShell の Unblock-File を使う。Microsoft の手順と同じ処理)
Sub UnblockFile(path)
    On Error Resume Next
    shell.Run "powershell.exe -NoProfile -Command ""Unblock-File -LiteralPath '" & Replace(path, "'", "''") & "'""", 0, True
    On Error GoTo 0
End Sub
