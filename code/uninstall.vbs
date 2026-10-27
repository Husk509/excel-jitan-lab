Option Explicit

Dim fso, shell, targetFolder, xlamDst, jsonDst, excelApp, item, ans, failed

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

targetFolder = fso.BuildPath(shell.ExpandEnvironmentStrings("%APPDATA%"), "Microsoft\AddIns")
xlamDst = fso.BuildPath(targetFolder, "CustomAddin.xlam")
jsonDst = fso.BuildPath(targetFolder, "config.json")

ans = MsgBox("Excel の「便利ツール」アドインを、削除します。" & vbCrLf & vbCrLf & _
             "削除するファイル: " & xlamDst & vbCrLf & _
             "(あなたの Excel ファイルには、触りません)" & vbCrLf & vbCrLf & _
             "Excel を開いている場合は、先にすべて閉じてください。続けますか?", _
             vbYesNo + vbQuestion + vbDefaultButton2, "アンインストール")
If ans <> vbYes Then WScript.Quit

' 1. Excel のアドイン登録を外す
On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number = 0 Then
    excelApp.Visible = False
    For Each item In excelApp.AddIns
        If UCase(item.Name) = "CUSTOMADDIN.XLAM" Then item.Installed = False
    Next
    excelApp.Quit
    Set excelApp = Nothing
End If
Err.Clear

' 2. ファイルを削除
failed = False
If fso.FileExists(xlamDst) Then
    fso.DeleteFile xlamDst, True
    If Err.Number <> 0 Then failed = True
End If
Err.Clear
If fso.FileExists(jsonDst) Then fso.DeleteFile jsonDst, True
Err.Clear
On Error GoTo 0

If failed Then
    MsgBox "ファイルを削除できませんでした。Excel をすべて閉じてから、もう一度、実行してください。", vbExclamation, "ご注意"
Else
    MsgBox "アンインストールが完了しました。Excel を開き直すと、「便利ツール」タブは消えています。", vbInformation, "アンインストール完了"
End If