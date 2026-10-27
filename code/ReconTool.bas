Attribute VB_Name = "ReconTool"
Option Explicit

' 表の突合(2つの表を、キーの列で照らし合わせて、差を一覧にする)
' リボンのボタン(customUI14.xml)の onAction から呼ばれる。元の表は書き換えない。結果は、新しいシートに出す。
' 例: 請求書と発注書 / 売上と入金 / 名簿の新旧 などの照合
Private Const TITLE As String = "突合ツール"

' ---------------------------------------------------------------
' 共通
' ---------------------------------------------------------------
' セルの値を、比べるための文字にする(前後の空白と、全角の空白をなくす)
Private Function Rc_Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Rc_Norm = "#ERR": Exit Function
    If IsEmpty(v) Then Exit Function
    ' 日付の値は、46336 のような数字ではなく、日付の文字にする(時刻が入っているときだけ、時刻も付ける)
    If VarType(v) = vbDate Then
        If v = Int(v) Then
            Rc_Norm = Format$(v, "yyyy/mm/dd")
        Else
            Rc_Norm = Format$(v, "yyyy/mm/dd hh:nn")
        End If
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, ChrW(&H3000), " ")
    Rc_Norm = Trim$(s)
End Function

' 値が同じか。数字として読めるものは、数字で比べる(tol 以内の差なら、同じ)
Private Function Rc_Same(ByVal a As Variant, ByVal b As Variant, ByVal tol As Double) As Boolean
    Dim sa As String, sb As String
    sa = Rc_Norm(a)
    sb = Rc_Norm(b)
    If sa = sb Then Rc_Same = True: Exit Function
    If Len(sa) > 0 And Len(sb) > 0 And Len(sa) <= 15 And Len(sb) <= 15 Then
        If IsNumeric(sa) And IsNumeric(sb) Then
            Rc_Same = (Abs(CDbl(sa) - CDbl(sb)) <= tol + 0.0000000001)
        End If
    End If
End Function

' キーの文字(大文字小文字は区別しない)。キーがすべて空白なら、blank を True にする
Private Function Rc_Key(d As Variant, ByVal r As Long, cols() As Long, ByRef blank As Boolean) As String
    Dim k As Long, s As String, part As String, anyVal As Boolean
    For k = 1 To UBound(cols)
        part = Rc_Norm(d(r, cols(k)))
        If Len(part) > 0 Then anyVal = True
        s = s & LCase$(part) & Chr(1)
    Next k
    blank = Not anyVal
    Rc_Key = s
End Function

Private Function Rc_Show(d As Variant, ByVal r As Long, cols() As Long) As String
    Dim k As Long, s As String
    For k = 1 To UBound(cols)
        If k > 1 Then s = s & " / "
        s = s & Rc_Norm(d(r, cols(k)))
    Next k
    Rc_Show = s
End Function

' 1行ぶんの中身を、「列名=値; 列名=値」の文字にする(長いときは、途中まで)
Private Function Rc_Desc(d As Variant, ByVal r As Long, hdr() As String, ByVal nc As Long) As String
    Dim c As Long, s As String, v As String
    For c = 1 To nc
        v = Rc_Norm(d(r, c))
        If Len(v) > 0 Then
            If Len(s) > 0 Then s = s & "; "
            s = s & hdr(c) & "=" & v
        End If
        If Len(s) > 250 Then s = Left$(s, 250) & "...": Exit For
    Next c
    Rc_Desc = s
End Function

' 結果の文字が、数式として動かないようにする
Private Function Rc_Safe(ByVal s As String) As String
    Dim c1 As String
    If Len(s) = 0 Then Exit Function
    c1 = Left$(s, 1)
    If c1 = "=" Or c1 = "+" Or c1 = "@" Or c1 = "-" Then Rc_Safe = "'" & s Else Rc_Safe = s
End Function

Private Sub Rc_Put(res() As Variant, ByRef cnt As Long, ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As Variant, ByVal e As Variant)
    cnt = cnt + 1
    res(cnt, 1) = a
    res(cnt, 2) = Rc_Safe(b)
    res(cnt, 3) = Rc_Safe(c)
    res(cnt, 4) = d
    res(cnt, 5) = e
End Sub

Private Function Rc_UniqueSheetName(ByVal baseName As String) As String
    Dim nm As String, n As Long, ws As Object, hit As Boolean
    nm = baseName
    n = 1
    Do
        hit = False
        For Each ws In ActiveWorkbook.Sheets
            If StrComp(ws.Name, nm, vbTextCompare) = 0 Then hit = True: Exit For
        Next ws
        If Not hit Then Exit Do
        n = n + 1
        nm = baseName & "(" & n & ")"
    Loop
    Rc_UniqueSheetName = nm
End Function

' ---------------------------------------------------------------
' ボタン
' ---------------------------------------------------------------
Sub Ribbon_Reconcile(control As IRibbonControl)
    Dim rA As Range, rB As Range, hdr As String, i As Long, keySpec As String, tolS As String
    Dim tol As Double, ws As Worksheet, msg As String
    If Workbooks.Count = 0 Then MsgBox "ブックを開いてから、押してください。", vbExclamation, TITLE: Exit Sub
    On Error Resume Next
    Set rA = Application.InputBox("1つ目の表を、見出しの行も入れて、ドラッグで選んでください。" & vbCrLf & _
                                  "(別のシートや、別のブックを選んでもかまいません。1つのセルだけ選ぶと、まわりの表全体を使います)", TITLE & " (1/4)", Type:=8)
    On Error GoTo 0
    If rA Is Nothing Then Exit Sub
    If rA.Cells.CountLarge = 1 Then Set rA = rA.CurrentRegion
    On Error Resume Next
    Set rB = Application.InputBox("2つ目の表を、見出しの行も入れて、選んでください。", TITLE & " (2/4)", Type:=8)
    On Error GoTo 0
    If rB Is Nothing Then Exit Sub
    If rB.Cells.CountLarge = 1 Then Set rB = rB.CurrentRegion
    If rA.Areas.Count > 1 Or rB.Areas.Count > 1 Then MsgBox "表は、1つの四角い範囲で選んでください。", vbExclamation, TITLE: Exit Sub

    For i = 1 To rA.Columns.Count
        If i > 20 Then hdr = hdr & "...": Exit For
        hdr = hdr & i & "=" & Left$(Rc_Norm(rA.Cells(1, i).Value2), 10) & "  "
    Next i
    keySpec = InputBox("照らし合わせるキー(伝票番号、顧客コードなど)の列を、見出しの名前か番号で入れてください。" & vbCrLf & _
                       "2つ以上は、カンマで区切ります。見出しの名前は、2つの表で同じである必要があります。" & vbCrLf & vbCrLf & _
                       "1つ目の表の列: " & hdr, TITLE & " (3/4)", Rc_Norm(rA.Cells(1, 1).Value2))
    If Len(Trim$(keySpec)) = 0 Then Exit Sub
    tolS = InputBox("数字の差が、いくつ以内なら「同じ」とみなしますか?" & vbCrLf & _
                    "0 なら、1円の違いも、差として出します。端数の違いを許すなら、1 などを入れます。", TITLE & " (4/4)", "0")
    If Len(Trim$(tolS)) = 0 Then Exit Sub
    tolS = StrConv(tolS, vbNarrow)
    If Not IsNumeric(tolS) Then MsgBox "数字で入れてください。", vbExclamation, TITLE: Exit Sub
    tol = Abs(CDbl(tolS))

    Set ws = Recon_Do(rA, rB, keySpec, tol, msg)
    MsgBox msg, IIf(ws Is Nothing, vbExclamation, vbInformation), TITLE
End Sub

' ---------------------------------------------------------------
' 本体。成功したら、結果のシートを返す(失敗したら Nothing)。msg に、ユーザーへの報告を入れる
' ---------------------------------------------------------------
Private Function Recon_Do(ByVal rA As Range, ByVal rB As Range, ByVal keySpec As String, ByVal tol As Double, ByRef msg As String) As Worksheet
    Dim dA As Variant, dB As Variant, nA As Long, nB As Long, cA As Long, cB As Long
    Dim hA() As String, hB() As String, dictHA As Object, dictHB As Object
    Dim toks() As String, i As Long, c As Long, r As Long, t As String, tn As String, idx As Long
    Dim nk As Long, kA() As Long, kB() As Long, kName() As String, bad As String
    Dim pA() As Long, pB() As Long, pName() As String, np As Long, isKey As Boolean, p As Long
    Dim onlyA As String, onlyB As String, nm As String
    Dim dictKB As Object, seenA As Object, bState() As Long, usedB() As Boolean
    Dim k As String, blank As Boolean, br As Long, diffs As String, nd As Long
    Dim res() As Variant, cnt As Long, cap As Long
    Dim nSame As Long, nDiff As Long, nOnlyA As Long, nOnlyB As Long, dupA As Long, dupB As Long, emptyA As Long, emptyB As Long
    Dim ws As Worksheet, last As Long, hd(1 To 1, 1 To 5) As Variant, sm(1 To 12, 1 To 2) As Variant
    Dim okA As Boolean, okB As Boolean

    On Error GoTo EH
    If rA.Rows.Count < 2 Or rB.Rows.Count < 2 Then
        msg = "表は、見出しの行のほかに、データが1行以上必要です。"
        Exit Function
    End If
    ' .Value2 だと、日付が数字(シリアル値)になるので、日付の型が残る .Value で読む
    dA = rA.Value
    dB = rB.Value
    nA = UBound(dA, 1): cA = UBound(dA, 2)
    nB = UBound(dB, 1): cB = UBound(dB, 2)

    ' 見出し
    Set dictHA = CreateObject("Scripting.Dictionary"): dictHA.CompareMode = vbTextCompare
    Set dictHB = CreateObject("Scripting.Dictionary"): dictHB.CompareMode = vbTextCompare
    ReDim hA(1 To cA): ReDim hB(1 To cB)
    For c = 1 To cA
        hA(c) = Rc_Norm(dA(1, c))
        If Len(hA(c)) > 0 Then If Not dictHA.Exists(hA(c)) Then dictHA.Add hA(c), c
    Next c
    For c = 1 To cB
        hB(c) = Rc_Norm(dB(1, c))
        If Len(hB(c)) > 0 Then If Not dictHB.Exists(hB(c)) Then dictHB.Add hB(c), c
    Next c

    ' キーの列
    keySpec = Replace(Replace(keySpec, "、", ","), "，", ",")
    toks = Split(keySpec, ",")
    ReDim kA(1 To UBound(toks) + 1): ReDim kB(1 To UBound(toks) + 1): ReDim kName(1 To UBound(toks) + 1)
    For i = 0 To UBound(toks)
        t = Trim$(toks(i))
        If Len(t) > 0 Then
            tn = StrConv(t, vbNarrow)
            If Not dictHA.Exists(t) And IsNumeric(tn) Then
                idx = CLng(Val(tn))
                If idx >= 1 And idx <= cA Then t = hA(idx)
            End If
            If Len(t) > 0 And dictHA.Exists(t) And dictHB.Exists(t) Then
                nk = nk + 1
                kA(nk) = dictHA(t): kB(nk) = dictHB(t): kName(nk) = t
            Else
                bad = bad & "「" & t & "」"
            End If
        End If
    Next i
    If Len(bad) > 0 Or nk = 0 Then
        msg = "キーの列が見つかりませんでした: " & bad & vbCrLf & "2つの表に、同じ名前の見出しがあるか、確認してください。"
        Exit Function
    End If
    ReDim Preserve kA(1 To nk): ReDim Preserve kB(1 To nk): ReDim Preserve kName(1 To nk)

    ' 比べる列(キー以外で、見出しの名前が同じ列)
    ReDim pA(1 To cA): ReDim pB(1 To cA): ReDim pName(1 To cA)
    For c = 1 To cA
        nm = hA(c)
        If Len(nm) > 0 Then
            If dictHA(nm) = c Then
                isKey = False
                For i = 1 To nk
                    If StrComp(kName(i), nm, vbTextCompare) = 0 Then isKey = True
                Next i
                If Not isKey Then
                    If dictHB.Exists(nm) Then
                        np = np + 1
                        pA(np) = c: pB(np) = dictHB(nm): pName(np) = nm
                    Else
                        onlyA = onlyA & "「" & nm & "」"
                    End If
                End If
            End If
        End If
    Next c
    For c = 1 To cB
        nm = hB(c)
        If Len(nm) > 0 Then
            If dictHB(nm) = c And Not dictHA.Exists(nm) Then onlyB = onlyB & "「" & nm & "」"
        End If
    Next c

    ' 表Bのキーを、先に集める
    Set dictKB = CreateObject("Scripting.Dictionary")
    Set seenA = CreateObject("Scripting.Dictionary")
    ReDim bState(2 To nB): ReDim usedB(2 To nB)
    For r = 2 To nB
        k = Rc_Key(dB, r, kB, blank)
        If blank Then
            bState(r) = 2
            emptyB = emptyB + 1
        ElseIf dictKB.Exists(k) Then
            bState(r) = 1
            dupB = dupB + 1
        Else
            dictKB.Add k, r
        End If
    Next r

    cap = (nA - 1) + (nB - 1)
    ReDim res(1 To cap, 1 To 5)

    ' 表Aを順に見る
    For r = 2 To nA
        k = Rc_Key(dA, r, kA, blank)
        If blank Then
            emptyA = emptyA + 1
            Rc_Put res, cnt, "キー空(A)", "(キーが空白)", Rc_Desc(dA, r, hA, cA), rA.Row + r - 1, ""
        ElseIf seenA.Exists(k) Then
            dupA = dupA + 1
            Rc_Put res, cnt, "重複(A)", Rc_Show(dA, r, kA), "同じキーが、" & (rA.Row + seenA(k) - 1) & "行目にもあります", rA.Row + r - 1, ""
        Else
            seenA.Add k, r
            If dictKB.Exists(k) Then
                br = dictKB(k)
                usedB(br) = True
                diffs = ""
                nd = 0
                For p = 1 To np
                    If Not Rc_Same(dA(r, pA(p)), dB(br, pB(p)), tol) Then
                        nd = nd + 1
                        If nd > 1 Then diffs = diffs & "; "
                        diffs = diffs & pName(p) & ": " & Rc_Norm(dA(r, pA(p))) & " → " & Rc_Norm(dB(br, pB(p)))
                    End If
                Next p
                If nd = 0 Then
                    nSame = nSame + 1
                Else
                    nDiff = nDiff + 1
                    Rc_Put res, cnt, "差異", Rc_Show(dA, r, kA), diffs, rA.Row + r - 1, rB.Row + br - 1
                End If
            Else
                nOnlyA = nOnlyA + 1
                Rc_Put res, cnt, "Aのみ", Rc_Show(dA, r, kA), Rc_Desc(dA, r, hA, cA), rA.Row + r - 1, ""
            End If
        End If
    Next r

    ' 表Bの、残り
    For r = 2 To nB
        Select Case bState(r)
            Case 2
                Rc_Put res, cnt, "キー空(B)", "(キーが空白)", Rc_Desc(dB, r, hB, cB), "", rB.Row + r - 1
            Case 1
                k = Rc_Key(dB, r, kB, blank)
                Rc_Put res, cnt, "重複(B)", Rc_Show(dB, r, kB), "同じキーが、" & (rB.Row + dictKB(k) - 1) & "行目にもあります", "", rB.Row + r - 1
            Case Else
                If Not usedB(r) Then
                    nOnlyB = nOnlyB + 1
                    Rc_Put res, cnt, "Bのみ", Rc_Show(dB, r, kB), Rc_Desc(dB, r, hB, cB), "", rB.Row + r - 1
                End If
        End Select
    Next r

    ' 検算
    okA = ((nA - 1) = nSame + nDiff + nOnlyA + dupA + emptyA)
    okB = ((nB - 1) = nSame + nDiff + nOnlyB + dupB + emptyB)

    ' 結果のシート
    Application.ScreenUpdating = False
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = Rc_UniqueSheetName("突合結果")
    hd(1, 1) = "結果"
    hd(1, 2) = "キー(" & Join(kName, " / ") & ")"
    hd(1, 3) = "内容"
    hd(1, 4) = "Aの行"
    hd(1, 5) = "Bの行"
    ws.Range("A1:E1").Value2 = hd
    last = cnt + 1
    If cnt > 0 Then
        ws.Range("A2").Resize(cap, 5).Value2 = res
    Else
        ws.Range("A2").Value2 = "(すべて一致しました)"
        last = 2
    End If
    sm(1, 1) = "表Aの行数(見出し除く)": sm(1, 2) = nA - 1
    sm(2, 1) = "表Bの行数(見出し除く)": sm(2, 2) = nB - 1
    sm(3, 1) = "一致": sm(3, 2) = nSame
    sm(4, 1) = "差異": sm(4, 2) = nDiff
    sm(5, 1) = "Aのみ": sm(5, 2) = nOnlyA
    sm(6, 1) = "Bのみ": sm(6, 2) = nOnlyB
    sm(7, 1) = "重複(A)": sm(7, 2) = dupA
    sm(8, 1) = "重複(B)": sm(8, 2) = dupB
    sm(9, 1) = "キー空(A)": sm(9, 2) = emptyA
    sm(10, 1) = "キー空(B)": sm(10, 2) = emptyB
    sm(11, 1) = "検算(A)": sm(11, 2) = IIf(okA, "OK", "不一致")
    sm(12, 1) = "検算(B)": sm(12, 2) = IIf(okB, "OK", "不一致")
    ws.Range("G1:H12").Value2 = sm
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(217, 217, 217)
    ws.Range("G1:G12").Font.Bold = True
    With ws.Range("A2:E" & last).FormatConditions
        .Delete
        .Add(2, , "=$A2=""差異""").Interior.Color = RGB(252, 228, 214)
        .Add(2, , "=$A2=""Aのみ""").Interior.Color = RGB(255, 242, 204)
        .Add(2, , "=$A2=""Bのみ""").Interior.Color = RGB(221, 235, 247)
        .Add(2, , "=LEFT($A2,2)=""重複""").Interior.Color = RGB(237, 237, 237)
        .Add(2, , "=LEFT($A2,2)=""キー""").Interior.Color = RGB(237, 237, 237)
    End With
    ws.Columns("A:H").AutoFit
    If ws.Columns(3).ColumnWidth > 70 Then ws.Columns(3).ColumnWidth = 70
    On Error Resume Next
    ws.Range("A1:E" & last).AutoFilter
    ws.Activate
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo EH
    Application.ScreenUpdating = True

    msg = "突合しました。新しいシート「" & ws.Name & "」を見てください。" & vbCrLf & vbCrLf & _
          "表A " & (nA - 1) & " 行 / 表B " & (nB - 1) & " 行 / キー: " & Join(kName, " + ") & vbCrLf & _
          "一致 " & nSame & " / 差異 " & nDiff & " / Aのみ " & nOnlyA & " / Bのみ " & nOnlyB & vbCrLf & _
          "重複 A " & dupA & ", B " & dupB & " / キー空白 A " & emptyA & ", B " & emptyB & vbCrLf
    msg = msg & "検算: " & IIf(okA And okB, "行数がすべて合いました。", "行数が合いません。結果を確認してください。")
    If np = 0 Then msg = msg & vbCrLf & "※ 比べる列がありません(キーの有無だけ見ました)。"
    If Len(onlyA) > 0 Then msg = msg & vbCrLf & "※ Aにしかない列(比べていません): " & onlyA
    If Len(onlyB) > 0 Then msg = msg & vbCrLf & "※ Bにしかない列(比べていません): " & onlyB
    If tol > 0 Then msg = msg & vbCrLf & "※ 数字の差が " & tol & " 以内は、同じとしました。"
    Set Recon_Do = ws
    Exit Function
EH:
    Application.ScreenUpdating = True
    msg = "うまくいきませんでした。" & vbCrLf & Err.Description
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Function
