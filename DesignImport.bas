Attribute VB_Name = "modDesignImport"
Option Explicit
'==========================================================================
' サービス設計書 取込マクロ
'   取込ファイル（拡張子なしテキスト）を読み、部品シートのブロックを
'   下へ下へコピーして設計書ブック（1プログラム=1ファイル）を作成する。
'
'   公開マクロ
'     設計書取込_フォルダ一括   … フォルダ内の全取込ファイルを処理
'     設計書取込_ファイル指定   … ファイルを選んで処理（複数選択可）
'     設計書取込_フォルダ一括_フロー図なし / 設計書取込_ファイル指定_フロー図なし
'                               … 上と同じだが概要フローの図形を作らない
'     取込ファイル_検証         … 書式チェックのみ（ブックは作らない）
'     取込ファイル_ソース突合   … Javaソースと突き合わせ、抜け・まとめ書きを「突合結果」シートに一覧
'     パレット初期化            … フローオブジェクトシートに既定図形を作成
'==========================================================================

Private Const SH_PARTS As String = "部品"
Private Const SH_CONF As String = "設定"
Private Const SH_LOG As String = "取込ログ"
Private Const SH_PALETTE As String = "フローオブジェクト"
Private Const LEVEL_KEY As String = "階層"

' 処理定義シートのレイアウト
Private Const PROC_TEXT_COL As Long = 4      ' 本文の開始列（D列）
Private Const PROC_INDENT As Long = 2        ' 階層1つあたりのずらし列数
Private Const PROC_NOTE_COL As Long = 55     ' 備考の開始列（BC列）
Private Const LINE_PT As Double = 13.5       ' 改行1行あたりの行高

Private gLastCol As Long
Private gConf As Object       ' 設定値
Private gCache As Object      ' 部品ごとのプレースホルダ位置
Private gDoc As Object        ' DOCブロックの値
Private gWarn As String       ' 1ファイル分の警告
Private gStage As String      ' 処理中の位置（エラー時にログへ出す）
Private gSeq As Long          ' 図形名の連番
Private gNoFlow As Boolean    ' True=フロー図を作らない

'==========================================================================
' 公開マクロ
'==========================================================================
Public Sub 設計書取込_フォルダ一括()
    ImportFolder False
End Sub

Public Sub 設計書取込_ファイル指定()
    ImportSelected False
End Sub

' フロー図（概要フロー）を作らない版。図形を作らないので速く、確実に動く
Public Sub 設計書取込_フォルダ一括_フロー図なし()
    ImportFolder True
End Sub

Public Sub 設計書取込_ファイル指定_フロー図なし()
    ImportSelected True
End Sub

Private Sub ImportFolder(ByVal noFlow As Boolean)
    Dim inDir As String, f As String, files As New Collection
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "取込ファイルのあるフォルダを選択"
        If .Show <> -1 Then Exit Sub
        inDir = .SelectedItems(1)
    End With
    f = Dir(inDir & "\*")
    Do While f <> ""
        If IsTargetFile(f) Then files.Add inDir & "\" & f
        f = Dir()
    Loop
    If files.Count = 0 Then
        MsgBox "取込対象のファイルがありません（拡張子なし または .txt）", vbExclamation
        Exit Sub
    End If
    RunImport files, inDir, noFlow
End Sub

Private Sub ImportSelected(ByVal noFlow As Boolean)
    Dim sel As Variant, files As New Collection, i As Long
    sel = Application.GetOpenFilename("すべてのファイル (*.*),*.*", , "取込ファイルを選択", , True)
    If Not IsArray(sel) Then Exit Sub
    For i = LBound(sel) To UBound(sel)
        files.Add CStr(sel(i))
    Next
    RunImport files, ParentDir(CStr(sel(LBound(sel)))), noFlow
End Sub

Public Sub 取込ファイル_検証()
    Dim sel As Variant, i As Long, secs As Collection, doc As Object, msg As String, sec As Object, sb As Object, n As Long, j As Long
    sel = Application.GetOpenFilename("すべてのファイル (*.*),*.*", , "検証するファイルを選択", , True)
    If Not IsArray(sel) Then Exit Sub
    InitRun
    For i = LBound(sel) To UBound(sel)
        gWarn = ""
        Set secs = ParseFile(CStr(sel(i)), doc)
        msg = ""
        For Each sec In secs
            n = 0
            For j = 1 To sec("SUBS").Count
                Set sb = sec("SUBS")(j)
                n = n + sb("ROWS").Count
                If Not BlockExists(sec("TYPE"), sb("NAME")) And sec("TYPE") <> "DOC" Then
                    AddWarn "部品がありません: " & BlockBase(sec("TYPE"), sb("NAME"))
                End If
            Next
            msg = msg & sec("TYPE") & "(" & n & "行) "
        Next
        WriteLog CStr(sel(i)), IIf(gWarn = "", "OK", "警告"), "[検証] " & msg & gWarn
    Next
    EndRun
    ThisWorkbook.Worksheets(SH_LOG).Activate
    MsgBox "検証が終わりました。取込ログシートを確認してください。", vbInformation
End Sub

'==========================================================================
' 実行制御
'==========================================================================
Private Sub RunImport(files As Collection, ByVal inDir As String, ByVal noFlow As Boolean)
    Dim outDir As String, f As Variant, ok As Long, ng As Long, t As Single
    InitRun
    gNoFlow = noFlow
    outDir = Trim$(ConfStr("出力フォルダ", ""))
    If outDir = "" Then outDir = inDir & "\出力"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    t = Timer
    For Each f In files
        Application.StatusBar = "取込中: " & f
        If ImportOne(CStr(f), outDir) Then ok = ok + 1 Else ng = ng + 1
    Next
    EndRun
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(SH_LOG).Activate
    MsgBox "完了しました。" & IIf(noFlow, "（フロー図なし）", "") & vbLf & "成功: " & ok & " 件 / 失敗: " & ng & " 件" & vbLf & _
           "出力先: " & outDir & vbLf & "処理時間: " & Format(Timer - t, "0.0") & " 秒", vbInformation
End Sub

Private Sub InitRun()
    LoadConf
    gLastCol = CLng(ConfNum("最終列", 90))
    Set gCache = NewDict()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndRun()
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function ImportOne(ByVal path As String, ByVal outDir As String) As Boolean
    Dim secs As Collection, doc As Object, wb As Workbook, ws As Worksheet
    Dim sec As Object, nextRow As Object, shName As String, r As Long, fname As String
    On Error GoTo EH
    gWarn = ""
    gStage = "読込"
    Set secs = ParseFile(path, doc)
    Set gDoc = doc
    If secs.Count = 0 Then Err.Raise vbObjectError + 1, , "#BEGIN ブロックがありません"

    Set wb = NewOutputBook()
    Set nextRow = NewDict()
    For Each ws In wb.Worksheets
        FillSheetHeader ws
        nextRow(ws.Name) = LastRow(ws) + 2
    Next

    For Each sec In secs
        If sec("TYPE") = "FLOW" And gNoFlow Then
            ' フロー図なし版: FLOWブロックは読み飛ばす
        ElseIf sec("TYPE") <> "DOC" Then
            shName = ConfStr(sec("TYPE"), "")
            If shName = "" Then
                AddWarn "種別「" & sec("TYPE") & "」の出力シートが設定にないためスキップ"
            ElseIf Not nextRow.Exists(shName) Then
                AddWarn "出力シート「" & shName & "」がないためスキップ"
            Else
                Set ws = wb.Worksheets(shName)
                r = RenderSection(sec, ws, nextRow(shName))
                nextRow(shName) = r + CLng(ConfNum("ブロック間の空行", 1))
            End If
        End If
    Next

    For Each ws In wb.Worksheets
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.ScrollRow = 1
    Next
    wb.Worksheets(1).Activate

    fname = SafeFileName(ReplaceHolders(ConfStr("ファイル名", "サービス設計書_{{処理ID}}.xlsx"), MakeCtx(gDoc)))
    If LCase$(Right$(fname, 5)) <> ".xlsx" Then fname = fname & ".xlsx"
    wb.SaveAs outDir & "\" & fname, xlOpenXMLWorkbook
    wb.Close False
    WriteLog path, IIf(gWarn = "", "OK", "警告"), fname & " " & gWarn
    ImportOne = True
    Exit Function
EH:
    WriteLog path, "NG", Err.Description & "（処理位置: " & gStage & "）" & gWarn
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

Private Function NewOutputBook() As Workbook
    Dim nm As Variant, arr() As Variant, i As Long
    nm = Split(ConfStr("出力シート", ""), ",")
    ReDim arr(0 To UBound(nm))
    For i = 0 To UBound(nm)
        arr(i) = Trim$(nm(i))
    Next
    ThisWorkbook.Worksheets(arr).Copy
    Set NewOutputBook = ActiveWorkbook
End Function

'==========================================================================
' 取込ファイルの解析
'   戻り値: Collection of section(Dictionary)
'     section("TYPE")  種別
'     section("SUBS")  Collection of sub(Dictionary)  ※1番目はメイン
'     sub("NAME") / sub("ATTR") / sub("COLS") / sub("ROWS")
'==========================================================================
Private Function ParseFile(ByVal path As String, ByRef doc As Object) As Collection
    Dim txt As String, lines() As String, i As Long, ln As String, p As Long
    Dim secs As New Collection, sec As Object, sb As Object, key As String, arr As Variant

    Set doc = NewDict()
    txt = ReadText(path, ConfStr("文字コード", "UTF-8"))
    txt = Replace(txt, ChrW(&HFEFF), "")
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If ln = "" Then GoTo NextLine
        If Left$(ln, 2) = "--" Or Left$(ln, 3) = "```" Then GoTo NextLine

        If UCase$(Left$(ln, 7)) = "#BEGIN " Then
            Set sec = NewSection(UCase$(Trim$(Mid$(ln, 8))))
            Set sb = sec("SUBS")(1)
            secs.Add sec
        ElseIf UCase$(ln) = "#END" Then
            Set sec = Nothing
            Set sb = Nothing
        ElseIf sec Is Nothing Then
            AddWarn (i + 1) & "行目: ブロック外の行を無視"
        ElseIf Left$(ln, 1) = "@" Then
            p = InStr(ln, "=")
            If p = 0 Then
                AddWarn (i + 1) & "行目: @行に = がありません"
            Else
                key = Trim$(Mid$(ln, 2, p - 2))
                If sec("TYPE") = "DOC" Then
                    doc(key) = Unescape(Mid$(ln, p + 1))
                Else
                    sb("ATTR")(key) = Unescape(Mid$(ln, p + 1))
                End If
            End If
        ElseIf Left$(ln, 1) = ">" Then
            Set sb = NewSub(UCase$(Trim$(Mid$(ln, 2))))
            sec("SUBS").Add sb
        Else
            arr = SplitRow(ln)
            If IsEmpty(sb("COLS")) Then
                sb("COLS") = arr
            Else
                sb("ROWS").Add FitRow(arr, UBound(sb("COLS")), i + 1)
            End If
        End If
NextLine:
    Next
    Set ParseFile = secs
End Function

Private Function NewSection(ByVal typ As String) As Object
    Dim d As Object, subs As New Collection
    Set d = NewDict()
    d("TYPE") = typ
    subs.Add NewSub("")
    Set d("SUBS") = subs
    Set NewSection = d
End Function

Private Function NewSub(ByVal nm As String) As Object
    Dim d As Object, rows As New Collection
    Set d = NewDict()
    d("NAME") = nm
    Set d("ATTR") = NewDict()
    d("COLS") = Empty
    Set d("ROWS") = rows
    Set NewSub = d
End Function

Private Function SplitRow(ByVal s As String) As Variant
    Dim a As Variant, i As Long
    s = Replace(s, "\\", Chr$(2))
    s = Replace(s, "\|", Chr$(1))
    a = Split(s, "|")
    For i = 0 To UBound(a)
        a(i) = Unesc(Trim$(a(i)))
    Next
    SplitRow = a
End Function

Private Function Unescape(ByVal s As String) As String
    s = Replace(s, "\\", Chr$(2))
    s = Replace(s, "\|", Chr$(1))
    Unescape = Unesc(Trim$(s))
End Function

Private Function Unesc(ByVal s As String) As String
    s = Replace(s, "\n", vbLf)
    s = Replace(s, Chr$(1), "|")
    Unesc = Replace(s, Chr$(2), "\")
End Function

' 列数に合わせる（足りなければ空、多ければ最終列に連結）
Private Function FitRow(a As Variant, ByVal last As Long, ByVal lineNo As Long) As Variant
    Dim r() As Variant, i As Long, extra As String
    ReDim r(0 To last)
    For i = 0 To UBound(a)
        If i <= last Then
            r(i) = a(i)
        ElseIf a(i) <> "" Then
            extra = extra & "|" & a(i)
        End If
    Next
    For i = UBound(a) + 1 To last
        r(i) = ""
    Next
    If extra <> "" Then
        r(last) = r(last) & extra
        AddWarn lineNo & "行目: 列数超過（最終列に連結）"
    End If
    FitRow = r
End Function

Private Function ReadText(ByVal path As String, ByVal cs As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = cs
    st.Open
    st.LoadFromFile path
    ReadText = st.ReadText(-1)
    st.Close
End Function

'==========================================================================
' セクション描画
'==========================================================================
Private Function RenderSection(sec As Object, ws As Worksheet, ByVal r As Long) As Long
    Dim typ As String, main As Object, sb As Object, ctx As Collection, i As Long
    Dim cnt As Object, nm As String, base As String, dNo As Object

    typ = sec("TYPE")
    Set main = sec("SUBS")(1)
    gStage = typ & " " & GetVal(MakeCtx(main("ATTR")), "処理No") & " " & GetVal(MakeCtx(main("ATTR")), "編集名") & _
             GetVal(MakeCtx(main("ATTR")), "アクセス名") & GetVal(MakeCtx(main("ATTR")), "論理名称")
    Set ctx = MakeCtx(main("ATTR"), gDoc)

    If HasName("TPL_" & typ & "_H") Then r = r + PlaceBlock(ws, "TPL_" & typ & "_H", r, ctx)

    Select Case typ
        Case "PROC": r = RenderProc(ws, r, main)
        Case "FLOW": r = RenderFlow(ws, r, main)
        Case Else:   r = RenderRows(ws, "TPL_" & typ & "_R", r, main, ctx)
    End Select

    Set cnt = NewDict()
    For i = 2 To sec("SUBS").Count
        Set sb = sec("SUBS")(i)
        nm = sb("NAME")
        cnt(nm) = cnt(nm) + 1
        Set dNo = NewDict()
        dNo("_N") = CStr(cnt(nm))
        dNo("_NO") = Circled(cnt(nm))
        base = BlockBase(typ, nm)
        If Not BlockExists(typ, nm) Then
            AddWarn "部品がありません: " & base
        Else
            Set ctx = MakeCtx(sb("ATTR"), dNo, main("ATTR"), gDoc)
            If HasName(base & "_H") Then r = r + PlaceBlock(ws, base & "_H", r, ctx)
            r = RenderRows(ws, base & "_R", r, sb, ctx)
            If HasName(base & "_F") Then r = r + PlaceBlock(ws, base & "_F", r, ctx)
        End If
    Next

    If HasName("TPL_" & typ & "_F") Then r = r + PlaceBlock(ws, "TPL_" & typ & "_F", r, MakeCtx(main("ATTR"), gDoc))
    RenderSection = r
End Function

Private Function RenderRows(ws As Worksheet, ByVal nm As String, ByVal r As Long, sb As Object, ctx As Collection) As Long
    Dim n As Long, h As Long, k As Long, rowCtx As Collection
    n = sb("ROWS").Count
    If n = 0 Then RenderRows = r: Exit Function
    If Not HasName(nm) Then
        AddWarn "明細の部品がありません: " & nm
        RenderRows = r: Exit Function
    End If
    h = CopyBlock(ws, nm, r, n)
    For k = 1 To n
        Set rowCtx = PrependCtx(RowDict(sb("COLS"), sb("ROWS")(k)), ctx)
        FillBlock ws, nm, r + (k - 1) * h, rowCtx
    Next
    SolidBottom ws, r + n * h - 1
    RenderRows = r + n * h
End Function

' 表の最終行: 点線の下罫線を実線にする（締めの線）
Private Sub SolidBottom(ws As Worksheet, ByVal r As Long)
    Dim c As Long
    For c = 1 To gLastCol
        With ws.Cells(r, c).Borders(xlEdgeBottom)
            If .LineStyle <> xlNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End If
        End With
    Next
End Sub

Private Function PlaceBlock(ws As Worksheet, ByVal nm As String, ByVal r As Long, ctx As Collection) As Long
    Dim h As Long
    h = CopyBlock(ws, nm, r, 1)
    FillBlock ws, nm, r, ctx
    PlaceBlock = h
End Function

' 部品の行を times 回コピー（倍々でコピーして高速化）。戻り値=部品の行数
Private Function CopyBlock(ws As Worksheet, ByVal nm As String, ByVal r As Long, ByVal times As Long) As Long
    Dim src As Range, h As Long, done As Long, n As Long, i As Long
    Set src = BlockRange(nm)
    h = src.Rows.Count
    src.Copy Destination:=ws.Cells(r, 1)
    done = 1
    Do While done < times
        n = done
        If n > times - done Then n = times - done
        ws.Range(ws.Cells(r, 1), ws.Cells(r + n * h - 1, gLastCol)).Copy Destination:=ws.Cells(r + done * h, 1)
        done = done + n
    Loop
    For i = 0 To h * times - 1
        ws.Rows(r + i).RowHeight = src.Rows((i Mod h) + 1).RowHeight
    Next
    CopyBlock = h
End Function

' {{キー}} を値に置換。{{>キー}} は「階層」の分だけ右へずらす
Private Sub FillBlock(ws As Worksheet, ByVal nm As String, ByVal r As Long, ctx As Collection)
    Dim items As Collection, it As Variant, s As String, key As String, v As String
    Dim lvl As Long, rr As Long, lines As Object, k As Variant, cur As Double, need As Double
    Set items = GetHolders(nm)
    Set lines = NewDict()
    For Each it In items
        rr = r + it(0)
        s = it(2)
        If Left$(s, 3) = "{{>" And Right$(s, 2) = "}}" Then
            key = Mid$(s, 4, Len(s) - 5)
            v = GetVal(ctx, key)
            lvl = CLng(Val(GetVal(ctx, LEVEL_KEY)))
            If lvl < 1 Then lvl = 1
            ws.Cells(rr, it(1)).Value = Empty
            PutValue ws.Cells(rr, it(1) + lvl - 1), v
        Else
            v = ReplaceHolders(s, ctx)
            PutValue ws.Cells(rr, it(1)), v
        End If
        If InStr(v, vbLf) > 0 Then
            If UBound(Split(v, vbLf)) + 1 > lines(rr) Then lines(rr) = UBound(Split(v, vbLf)) + 1
        End If
    Next
    For Each k In lines.Keys
        cur = ws.Rows(k).RowHeight
        need = lines(k) * LINE_PT + 3
        If need > cur Then ws.Rows(k).RowHeight = need
    Next
End Sub

'==========================================================================
' 処理定義（種別|階層|値1|値2|備考）
'   T=文章  SQL=SQL箱  TH/TD=表（コード｜メッセージ）  SH/SD=設定項目表（項目｜設定値）  BR=空行
'==========================================================================
Private Function RenderProc(ws As Worksheet, ByVal r As Long, sb As Object) As Long
    Dim n As Long, k As Long, a As Variant, cols As Variant, r0 As Long
    Dim iKind As Long, iLvl As Long, iV1 As Long, iV2 As Long, iNote As Long
    Dim kind As String, lvl As Long, c As Long, note As String, v1 As String, v2 As String
    Dim nl As Variant, j As Long, rr As Long

    n = sb("ROWS").Count
    If n = 0 Then RenderProc = r: Exit Function
    cols = sb("COLS")
    iKind = ColIndex(cols, "種別", 0)
    iLvl = ColIndex(cols, "階層", 1)
    iV1 = ColIndex(cols, "値1", 2)
    iV2 = ColIndex(cols, "値2", 3)
    iNote = ColIndex(cols, "備考", 4)

    CopyBlock ws, "TPL_PROC_R", r, n
    r0 = r
    For k = 1 To n
        a = sb("ROWS")(k)
        gStage = "PROC " & k & "行目"
        kind = UCase$(Fld(a, iKind))
        lvl = CLng(Val(Fld(a, iLvl)))
        If lvl < 1 Then lvl = 1
        c = PROC_TEXT_COL + (lvl - 1) * PROC_INDENT
        v1 = Fld(a, iV1)
        v2 = Fld(a, iV2)
        Select Case kind
            Case "BR"
            Case "SQL"
                DrawBox ws, r, c, 4, "SQL", "STY_SQL_LABEL"
                DrawBox ws, r, c + 4, 8, v1, "STY_SQL_ID"
                DrawBox ws, r, c + 12, 28, v2, "STY_SQL_NAME"
            Case "TH"
                DrawBox ws, r, c, 5, v1, "STY_TH"
                DrawBox ws, r, c + 5, 32, v2, "STY_TH"
            Case "TD"
                DrawBox ws, r, c, 5, v1, "STY_TD"
                DrawBox ws, r, c + 5, 32, v2, "STY_TD"
            Case "SH"      ' 設定項目表の見出し（項目｜設定値）
                DrawBox ws, r, c, 16, v1, "STY_TH"
                DrawBox ws, r, c + 16, 24, v2, "STY_TH"
            Case "SD"      ' 設定項目表の明細
                DrawBox ws, r, c, 16, v1, "STY_TD"
                DrawBox ws, r, c + 16, 24, v2, "STY_TD"
            Case Else
                PutValue ws.Cells(r, c), v1
        End Select
        ' 備考は改行ごとに下の行へ書く（行の高さは変えない）
        note = Fld(a, iNote)
        If note <> "" Then
            nl = Split(note, vbLf)
            For j = 0 To UBound(nl)
                rr = r + j
                If rr > r0 + n - 1 Then rr = r0 + n - 1
                If CStr(ws.Cells(rr, PROC_NOTE_COL).Value) = "" Then
                    PutValue ws.Cells(rr, PROC_NOTE_COL), CStr(nl(j))
                Else
                    ws.Cells(rr, PROC_NOTE_COL).Value = ws.Cells(rr, PROC_NOTE_COL).Value & " " & nl(j)
                End If
            Next
        End If
        r = r + 1
    Next
    With ws.Range(ws.Cells(r0, 1), ws.Cells(r - 1, 3))
        .Merge
        .Value = "処理内容"
        .Orientation = xlVertical
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Font.Bold = True
    End With
    RenderProc = r
End Function

Private Sub DrawBox(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal w As Long, ByVal txt As String, ByVal styName As String)
    Dim rg As Range, sty As Range, b As Variant
    If c + w - 1 >= PROC_NOTE_COL Then w = PROC_NOTE_COL - c
    If w < 1 Then Exit Sub
    Set rg = ws.Range(ws.Cells(r, c), ws.Cells(r, c + w - 1))
    rg.Merge
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rg.Borders(b)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next
    If HasName(styName) Then
        Set sty = ThisWorkbook.Names(styName).RefersToRange.Cells(1, 1)
        If sty.Interior.ColorIndex <> xlNone Then rg.Interior.Color = sty.Interior.Color
        rg.Font.Bold = sty.Font.Bold
        rg.Font.Color = sty.Font.Color
        rg.HorizontalAlignment = sty.HorizontalAlignment
    End If
    rg.VerticalAlignment = xlCenter
    PutValue rg.Cells(1, 1), txt
End Sub

'==========================================================================
' フロー図（ID|種別|テキスト|次|列|...）
'   図形はフローオブジェクトシートの「FO_種別」をコピー。なければ標準図形。
'   図形内の {{テキスト}} {{ID}} {{処理No}} 等は行の値で置換。
'==========================================================================
Private Function RenderFlow(ws As Worksheet, ByVal r As Long, sb As Object) As Long
    Dim n As Long, k As Long, a As Variant, cols As Variant, d As Object
    Dim iId As Long, iKind As Long, iNext As Long, iCol As Long, iSlot As Long, slot As Long, dashed As Boolean
    Dim shp As Shape, shps As Object, pos As Object, ids() As String
    Dim gapY As Double, gapX As Double, x0 As Double, top0 As Double, maxBottom As Double
    Dim nx As String, tok As Variant, tid As String, lbl As String, p As Long, colNo As Long

    n = sb("ROWS").Count
    If n = 0 Then RenderFlow = r: Exit Function
    cols = sb("COLS")
    iId = ColIndex(cols, "ID", 0)
    iKind = ColIndex(cols, "種別", 1)
    iNext = ColIndex(cols, "次", 3)
    iCol = ColIndex(cols, "列", 4)
    iSlot = ColIndex(cols, "段", 99)
    gapY = ConfNum("フロー縦間隔", 60)
    gapX = ConfNum("フロー横間隔", 170)
    x0 = ConfNum("フロー左端", 120)

    top0 = ws.Cells(r, 1).Top + 10
    maxBottom = top0
    Set shps = NewDict()
    Set pos = NewDict()
    ReDim ids(1 To n)

    For k = 1 To n
        a = sb("ROWS")(k)
        Set d = RowDict(cols, a)
        ids(k) = Fld(a, iId)
        If ids(k) = "" Then ids(k) = "N" & k
        colNo = CLng(Val(Fld(a, iCol)))
        slot = CLng(Val(Fld(a, iSlot)))
        If slot < 1 Then slot = k
        gStage = "FLOW " & k & "行目（" & ids(k) & " " & Fld(a, iKind) & "）"
        Set shp = MakeFlowShape(ws, Fld(a, iKind), d, r)
        shp.Left = x0 + colNo * gapX - shp.Width / 2
        shp.Top = top0 + (slot - 1) * gapY + (40 - shp.Height) / 2
        On Error Resume Next
        shp.Name = "FL_" & ids(k)
        On Error GoTo 0
        Set shps(ids(k)) = shp
        pos(ids(k)) = Array(slot, colNo)
        If shp.Top + shp.Height > maxBottom Then maxBottom = shp.Top + shp.Height
    Next

    For k = 1 To n
        a = sb("ROWS")(k)
        nx = Fld(a, iNext)
        If nx = "" And k < n And Fld(a, iKind) <> "終了" Then nx = ids(k + 1)
        If nx <> "" And nx <> "-" Then
            For Each tok In Split(nx, ",")
                tid = Trim$(tok)
                lbl = ""
                dashed = (Left$(tid, 1) = "~")
                If dashed Then tid = Mid$(tid, 2)
                p = InStr(tid, ":")
                If p > 0 Then
                    lbl = Trim$(Mid$(tid, p + 1))
                    tid = Trim$(Left$(tid, p - 1))
                End If
                If shps.Exists(tid) Then
                    ConnectShapes ws, shps(ids(k)), shps(tid), pos(ids(k)), pos(tid), lbl, _
                                  IIf(dashed, "FO_データ線", "FO_線"), HasBetween(pos, pos(ids(k)), pos(tid))
                Else
                    AddWarn "フロー: 遷移先IDなし " & ids(k) & "→" & tid
                End If
            Next
        End If
    Next

    Do While ws.Cells(r, 1).Top < maxBottom + 15
        r = r + 1
    Loop
    RenderFlow = r
End Function

Private Function MakeFlowShape(ws As Worksheet, ByVal kind As String, d As Object, ByVal r As Long) As Shape
    ' クリップボードは使わない: 図形を新規作成し、パレット図形の書式を写す
    Dim pal As Worksheet, src As Shape, shp As Shape, i As Long, nm() As Variant, gi As Shape
    Set pal = PaletteSheet(ws.Parent)
    If Not pal Is Nothing Then Set src = FindShape(pal, "FO_" & kind)
    If src Is Nothing Then
        Set shp = AddDefaultShape(ws, kind, 0, 0)
        ReplaceShapeText shp, d
        Set MakeFlowShape = shp
        Exit Function
    End If
    If src.Type = msoGroup Then
        ReDim nm(1 To src.GroupItems.Count)
        For i = 1 To src.GroupItems.Count
            Set gi = src.GroupItems(i)
            Set shp = CloneShape(ws, gi, gi.Left - src.Left, gi.Top - src.Top, d, kind)
            gSeq = gSeq + 1
            shp.Name = "FLX_" & gSeq
            nm(i) = shp.Name
        Next
        Set MakeFlowShape = ws.Shapes.Range(nm).Group
    Else
        Set MakeFlowShape = CloneShape(ws, src, 0, 0, d, kind)
    End If
End Function

Private Function CloneShape(ws As Worksheet, ByVal src As Shape, ByVal x As Double, ByVal y As Double, _
                            ByVal d As Object, ByVal kind As String) As Shape
    Dim shp As Shape, t As String, ast As Long
    On Error Resume Next
    ast = src.AutoShapeType
    On Error GoTo 0
    If ast <= 0 Or ast = msoShapeNotPrimitive Then
        Set shp = AddDefaultShape(ws, kind, x, y)     ' 画像・アイコン等は標準図形で代替
        shp.Width = src.Width: shp.Height = src.Height
    Else
        Set shp = ws.Shapes.AddShape(ast, x, y, src.Width, src.Height)
        src.PickUp
        shp.Apply
    End If
    On Error Resume Next
    t = ""
    If src.TextFrame2.HasText Then t = src.TextFrame2.TextRange.Text
    If InStr(t, "{{") = 0 Then t = "{{テキスト}}"
    With shp.TextFrame2
        .TextRange.Text = Replace(ReplaceHolders(t, MakeCtx(d, gDoc)), vbLf, vbCr)
        .TextRange.Font.Size = src.TextFrame2.TextRange.Font.Size
        .TextRange.Font.Name = src.TextFrame2.TextRange.Font.Name
        .TextRange.Font.NameFarEast = src.TextFrame2.TextRange.Font.NameFarEast
        .TextRange.Font.Bold = src.TextFrame2.TextRange.Font.Bold
        .TextRange.Font.Fill.ForeColor.RGB = src.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        .TextRange.ParagraphFormat.Alignment = src.TextFrame2.TextRange.ParagraphFormat.Alignment
        .VerticalAnchor = src.TextFrame2.VerticalAnchor
        .WordWrap = src.TextFrame2.WordWrap
        .MarginLeft = src.TextFrame2.MarginLeft: .MarginRight = src.TextFrame2.MarginRight
        .MarginTop = src.TextFrame2.MarginTop: .MarginBottom = src.TextFrame2.MarginBottom
    End With
    On Error GoTo 0
    Set CloneShape = shp
End Function

Private Function AddDefaultShape(ws As Worksheet, ByVal kind As String, ByVal x As Double, ByVal y As Double) As Shape
    Dim t As Long, w As Double, h As Double, shp As Shape
    w = 130: h = 36
    Select Case kind
        Case "開始", "終了": t = msoShapeFlowchartTerminator
        Case "分岐", "判定": t = msoShapeFlowchartDecision: w = 140: h = 52
        Case "入出力": t = msoShapeFlowchartData
        Case "サブ": t = msoShapeFlowchartPredefinedProcess
        Case "テーブル": t = msoShapeCan: w = 96: h = 44
        Case "ループ始", "ループ終": t = msoShapeFlowchartPreparation
        Case "イベント": t = msoShapeRoundedRectangle
        Case Else: t = msoShapeFlowchartProcess
    End Select
    Set shp = ws.Shapes.AddShape(t, x, y, w, h)
    With shp
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(46, 117, 182)
        .Line.Weight = 1.25
        With .TextFrame2
            .TextRange.Text = "{{テキスト}}"
            .TextRange.Font.Size = 9
            .TextRange.Font.NameFarEast = "Meiryo UI"
            .TextRange.Font.Name = "Meiryo UI"
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
        End With
    End With
    Set AddDefaultShape = shp
End Function

Private Sub ReplaceShapeText(ByVal shp As Shape, ByVal d As Object)
    Dim i As Long, s As String, ctx As Collection
    Set ctx = MakeCtx(d, gDoc)
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            ReplaceShapeText shp.GroupItems(i), d
        Next
        Exit Sub
    End If
    On Error Resume Next
    If shp.TextFrame2.HasText Then
        s = shp.TextFrame2.TextRange.Text
        If InStr(s, "{{") > 0 Then shp.TextFrame2.TextRange.Text = Replace(ReplaceHolders(s, ctx), vbLf, vbCr)
    End If
    On Error GoTo 0
End Sub

Private Sub ConnectShapes(ByVal ws As Worksheet, ByVal a As Shape, ByVal b As Shape, ByVal pa As Variant, ByVal pb As Variant, _
                          ByVal lbl As String, ByVal styleName As String, ByVal skip As Boolean)
    ' 方向: 0=上 1=左 2=下 3=右
    Dim dA As Long, dB As Long, cn As Shape, pal As Worksheet, st As Shape
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, tb As Shape
    If pb(0) > pa(0) Then
        If pb(1) = pa(1) And skip Then
            dA = 3: dB = 3          ' 間に図形がある場合は右側を回る
        ElseIf pb(1) = pa(1) Then
            dA = 2: dB = 0
        ElseIf pb(1) > pa(1) Then
            dA = 3: dB = 0
        Else
            dA = 2: dB = 3
        End If
    ElseIf pb(0) = pa(0) Then
        If pb(1) > pa(1) Then
            dA = 3: dB = 1
        Else
            dA = 1: dB = 3
        End If
    Else
        If pb(1) = pa(1) Then
            dA = 1: dB = 1
        ElseIf pb(1) < pa(1) Then
            dA = 0: dB = 3
        Else
            dA = 3: dB = 1
        End If
    End If
    SitePoint a, dA, x1, y1
    SitePoint b, dB, x2, y2
    Set cn = ws.Shapes.AddConnector(msoConnectorElbow, x1, y1, x2, y2)
    If CanConnect(a) And CanConnect(b) Then
        On Error Resume Next
        cn.ConnectorFormat.BeginConnect a, SiteIndex(a, dA)
        cn.ConnectorFormat.EndConnect b, SiteIndex(b, dB)
        On Error GoTo 0
    End If
    Set pal = PaletteSheet(ws.Parent)
    If Not pal Is Nothing Then Set st = FindShape(pal, styleName)
    If Not st Is Nothing Then
        st.PickUp
        cn.Apply
    Else
        cn.Line.ForeColor.RGB = RGB(46, 117, 182)
        cn.Line.Weight = 1.25
        cn.Line.EndArrowheadStyle = msoArrowheadTriangle
        If styleName = "FO_データ線" Then
            cn.Line.DashStyle = msoLineDash
            cn.Line.BeginArrowheadStyle = msoArrowheadOval
        End If
    End If
    If lbl <> "" Then
        Select Case dA
            Case 2: Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x1 + 3, y1, 70, 14)
            Case 3: Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x1 + 2, y1 - 15, 70, 14)
            Case 0: Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x1 + 3, y1 - 15, 70, 14)
            Case Else: Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x1 - 72, y1 - 15, 70, 14)
        End Select
        With tb
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
            .TextFrame2.TextRange.Text = lbl
            .TextFrame2.TextRange.Font.Size = 8
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(192, 0, 0)
        End With
    End If
End Sub

' 同じ列で、a と b の間の段に別の図形があるか
Private Function HasBetween(ByVal pos As Object, ByVal pa As Variant, ByVal pb As Variant) As Boolean
    Dim k As Variant, p As Variant
    If pa(1) <> pb(1) Then Exit Function
    For Each k In pos.Keys
        p = pos(k)
        If p(1) = pa(1) And p(0) > pa(0) And p(0) < pb(0) Then HasBetween = True: Exit Function
    Next
End Function

Private Function CanConnect(ByVal s As Shape) As Boolean
    On Error Resume Next
    CanConnect = (s.ConnectionSiteCount >= 4 And s.ConnectionSiteCount Mod 4 = 0)
End Function

' 接続点は上から反時計回りに並ぶ（四角=4点、円=8点…）
Private Function SiteIndex(ByVal s As Shape, ByVal dirNo As Long) As Long
    SiteIndex = dirNo * (s.ConnectionSiteCount \ 4) + 1
End Function

Private Sub SitePoint(ByVal s As Shape, ByVal dirNo As Long, ByRef x As Double, ByRef y As Double)
    Select Case dirNo
        Case 0: x = s.Left + s.Width / 2: y = s.Top
        Case 1: x = s.Left: y = s.Top + s.Height / 2
        Case 2: x = s.Left + s.Width / 2: y = s.Top + s.Height
        Case Else: x = s.Left + s.Width: y = s.Top + s.Height / 2
    End Select
End Sub

Private Function PaletteSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set PaletteSheet = wb.Worksheets(SH_PALETTE)
End Function

Private Function FindShape(ByVal ws As Worksheet, ByVal nm As String) As Shape
    On Error Resume Next
    Set FindShape = ws.Shapes(nm)
End Function

'==========================================================================
' パレット初期化（既定図形を FO_種別 の名前で作成）
'   社内の図形を使う場合は、その図形を選んで名前ボックスに FO_開始 等と入力
'==========================================================================
Public Sub パレット初期化()
    Dim pal As Worksheet, kinds As Variant, i As Long, shp As Shape, cn As Shape, x As Double, y As Double
    Set pal = ThisWorkbook.Worksheets(SH_PALETTE)
    kinds = Array("開始", "終了", "処理", "分岐", "入出力", "サブ", "テーブル", "ループ始", "ループ終", "イベント")
    x = pal.Range("C10").Left
    y = pal.Range("C10").Top
    For i = 0 To UBound(kinds)
        If FindShape(pal, "FO_" & kinds(i)) Is Nothing Then
            Set shp = AddDefaultShape(pal, CStr(kinds(i)), x + (i Mod 5) * 150, y + (i \ 5) * 70)
            shp.Name = "FO_" & kinds(i)
            shp.AlternativeText = "種別「" & kinds(i) & "」。{{テキスト}}等が取込値に置換されます"
        End If
    Next
    If FindShape(pal, "FO_線") Is Nothing Then
        Set cn = pal.Shapes.AddConnector(msoConnectorStraight, x, y + 160, x + 120, y + 160)
        cn.Line.ForeColor.RGB = RGB(31, 78, 121)
        cn.Line.Weight = 1.25
        cn.Line.EndArrowheadStyle = msoArrowheadTriangle
        cn.Name = "FO_線"
    End If
    MsgBox "不足していた FO_ 図形を作成しました（既にある図形はそのまま）。" & vbLf & _
           "色や形は自由に変更できます（名前と {{テキスト}} は残してください）。", vbInformation
End Sub

'==========================================================================
' 部品・プレースホルダ
'==========================================================================
Private Function BlockBase(ByVal typ As String, ByVal subName As String) As String
    If subName = "" Then BlockBase = "TPL_" & typ Else BlockBase = "TPL_" & typ & "_" & subName
End Function

Private Function BlockExists(ByVal typ As String, ByVal subName As String) As Boolean
    Dim b As String
    b = BlockBase(typ, subName)
    If subName = "" And (typ = "PROC" Or typ = "FLOW") Then BlockExists = True: Exit Function
    BlockExists = HasName(b & "_H") Or HasName(b & "_R")
End Function

Private Function HasName(ByVal nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nm)
    HasName = Not (n Is Nothing)
End Function

Private Function BlockRange(ByVal nm As String) As Range
    Dim rg As Range, ws As Worksheet
    Set rg = ThisWorkbook.Names(nm).RefersToRange
    Set ws = rg.Worksheet
    Set BlockRange = ws.Range(ws.Cells(rg.Row, 1), ws.Cells(rg.Row + rg.Rows.Count - 1, gLastCol))
End Function

' 部品内の {{ }} 位置をキャッシュ: Array(行オフセット, 列, 文字列)
Private Function GetHolders(ByVal nm As String) As Collection
    Dim rg As Range, v As Variant, i As Long, j As Long, c As New Collection
    If gCache.Exists(nm) Then Set GetHolders = gCache(nm): Exit Function
    Set rg = BlockRange(nm)
    v = rg.Value
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            If VarType(v(i, j)) = vbString Then
                If InStr(v(i, j), "{{") > 0 Then c.Add Array(i - 1, j, CStr(v(i, j)))
            End If
        Next
    Next
    Set gCache(nm) = c
    Set GetHolders = c
End Function

Private Function ReplaceHolders(ByVal s As String, ctx As Collection) As String
    Dim p As Long, q As Long, key As String, out As String
    Do
        p = InStr(s, "{{")
        If p = 0 Then Exit Do
        q = InStr(p + 2, s, "}}")
        If q = 0 Then Exit Do
        key = Mid$(s, p + 2, q - p - 2)
        If Left$(key, 1) = ">" Then key = Mid$(key, 2)
        out = out & Left$(s, p - 1) & GetVal(ctx, key)
        s = Mid$(s, q + 2)
    Loop
    ReplaceHolders = out & s
End Function

Private Sub FillSheetHeader(ws As Worksheet)
    Dim c As Range, ctx As Collection
    Set ctx = MakeCtx(gDoc)
    For Each c In ws.UsedRange.Cells
        If VarType(c.Value) = vbString Then
            If InStr(c.Value, "{{") > 0 Then PutValue c, ReplaceHolders(c.Value, ctx)
        End If
    Next
End Sub

' 数値・日付に化けないよう文字列として書く
Private Sub PutValue(c As Range, ByVal v As String)
    If v = "" Then
        c.Value = Empty
    ElseIf IsNumeric(v) Or IsDate(v) Or InStr("=+-@", Left$(v, 1)) > 0 Then
        c.Value = "'" & v
    Else
        c.Value = v
    End If
End Sub

'==========================================================================
' 値の参照（ctx = 優先順のDictionaryのリスト）
'==========================================================================
Private Function MakeCtx(ParamArray ds() As Variant) As Collection
    Dim c As New Collection, i As Long
    For i = LBound(ds) To UBound(ds)
        c.Add ds(i)
    Next
    Set MakeCtx = c
End Function

Private Function PrependCtx(d As Object, ctx As Collection) As Collection
    Dim c As New Collection, x As Variant
    c.Add d
    For Each x In ctx
        c.Add x
    Next
    Set PrependCtx = c
End Function

Private Function GetVal(ctx As Collection, ByVal key As String) As String
    Dim d As Variant
    For Each d In ctx
        If d.Exists(key) Then GetVal = CStr(d(key)): Exit Function
    Next
    GetVal = ""
End Function

Private Function RowDict(cols As Variant, a As Variant) As Object
    Dim d As Object, i As Long
    Set d = NewDict()
    For i = 0 To UBound(cols)
        If cols(i) <> "" And i <= UBound(a) Then d(cols(i)) = a(i)
    Next
    Set RowDict = d
End Function

Private Function ColIndex(cols As Variant, ByVal nm As String, ByVal def As Long) As Long
    Dim i As Long
    For i = 0 To UBound(cols)
        If cols(i) = nm Then ColIndex = i: Exit Function
    Next
    ColIndex = def
End Function

Private Function Fld(a As Variant, ByVal i As Long) As String
    If i <= UBound(a) Then Fld = CStr(a(i))
End Function

'==========================================================================
' 設定・ログ・小物
'==========================================================================
Private Sub LoadConf()
    Dim ws As Worksheet, r As Long, k As String
    Set gConf = NewDict()
    Set ws = ThisWorkbook.Worksheets(SH_CONF)
    For r = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        k = Trim$(CStr(ws.Cells(r, 2).Value))
        If k <> "" Then gConf(k) = CStr(ws.Cells(r, 3).Value)
    Next
End Sub

Private Function ConfStr(ByVal k As String, ByVal def As String) As String
    If gConf Is Nothing Then LoadConf
    If gConf.Exists(k) Then
        If gConf(k) <> "" Then ConfStr = gConf(k): Exit Function
    End If
    ConfStr = def
End Function

Private Function ConfNum(ByVal k As String, ByVal def As Double) As Double
    Dim s As String
    s = ConfStr(k, "")
    If IsNumeric(s) Then ConfNum = CDbl(s) Else ConfNum = def
End Function

Private Sub WriteLog(ByVal path As String, ByVal result As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = path
    ws.Cells(r, 3).Value = result
    ws.Cells(r, 4).Value = msg
End Sub

Private Sub AddWarn(ByVal s As String)
    If Len(gWarn) < 2000 Then gWarn = gWarn & "[" & s & "] "
End Sub

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Function Circled(ByVal n As Long) As String
    If n >= 1 And n <= 20 Then Circled = ChrW(&H2460 + n - 1) Else Circled = "(" & n & ")"
End Function

Private Function LastRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find("*", , xlValues, , xlByRows, xlPrevious)
    If f Is Nothing Then LastRow = 1 Else LastRow = f.Row
End Function

Private Function IsTargetFile(ByVal f As String) As Boolean
    IsTargetFile = (InStr(f, ".") = 0) Or (LCase$(Right$(f, 4)) = ".txt")
End Function

Private Function ParentDir(ByVal p As String) As String
    ParentDir = Left$(p, InStrRev(p, "\") - 1)
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbLf, vbCr)
        s = Replace(s, ch, "_")
    Next
    SafeFileName = s
End Function

'==========================================================================
' ソース突合: Javaソースと取込ファイルを突き合わせ、抜け・まとめ書きを一覧にする
'==========================================================================
Public Sub 取込ファイル_ソース突合()
    Dim imp As Variant, src As Variant, f As Variant, srcText As String
    Dim methods As Object, pubs As Object, sqlIds As Object
    Dim secs As Collection, doc As Object, sec As Object, sb As Object, a As Variant, cols As Variant
    Dim procM As Object, procSql As Object, accIds As Object, editNames As Object, editRefs As Object
    Dim res As New Collection, k As Variant, i As Long, j As Long
    Dim iKind As Long, iV1 As Long, iV2 As Long, iNote As Long, pno As String, mname As String
    Dim id As String, note As String, ref As String, req As String

    imp = Application.GetOpenFilename("すべてのファイル (*.*),*.*", , "① 取込ファイルを選択", , False)
    If VarType(imp) = vbBoolean Then Exit Sub
    src = Application.GetOpenFilename("Javaソース (*.java),*.java,すべてのファイル (*.*),*.*", , "② 突き合わせるJavaソースを選択（複数可）", , True)
    If Not IsArray(src) Then Exit Sub

    InitRun
    On Error GoTo EH
    gWarn = ""
    For Each f In src
        srcText = srcText & vbLf & ReadText(CStr(f), ConfStr("ソース文字コード", "UTF-8"))
    Next
    srcText = Replace(Replace(srcText, vbCrLf, vbLf), vbCr, vbLf)
    JavaInventory srcText, methods, pubs, sqlIds

    ' --- 取込ファイル側 ---
    Set secs = ParseFile(CStr(imp), doc)
    Set procM = NewDict(): Set procSql = NewDict(): Set accIds = NewDict()
    Set editNames = NewDict(): Set editRefs = NewDict()
    For Each sec In secs
        Set sb = sec("SUBS")(1)
        Select Case sec("TYPE")
        Case "PROC"
            pno = GetVal(MakeCtx(sb("ATTR")), "処理No")
            mname = Trim$(GetVal(MakeCtx(sb("ATTR")), "メソッド"))
            If mname = "" Then
                res.Add Array("処理定義", pno, "@メソッド が未記入", pno, "")
            Else
                procM(mname) = pno
            End If
            If Not IsEmpty(sb("COLS")) Then
                cols = sb("COLS")
                iKind = ColIndex(cols, "種別", 0): iV1 = ColIndex(cols, "値1", 2)
                iV2 = ColIndex(cols, "値2", 3): iNote = ColIndex(cols, "備考", 4)
                For j = 1 To sb("ROWS").Count
                    a = sb("ROWS")(j)
                    If UCase$(Fld(a, iKind)) = "SQL" Then
                        id = Trim$(Fld(a, iV1))
                        If IsBundled(id) Then res.Add Array("まとめ書き", id, "SQL IDが「／」等でまとめられている", pno, "1行1SQLに分ける")
                        If Not procSql.Exists(id) Then procSql(id) = pno
                        If Trim$(Fld(a, iV2)) = "" Or Trim$(Fld(a, iV2)) = id Then
                            res.Add Array("SQL名称", id, "SQL名称が空またはIDと同じ", pno, "Javadoc等から日本語名を付ける")
                        End If
                    End If
                    note = Fld(a, iNote)
                    If InStr(note, "編集定義") > 0 Then
                        ref = RefName(note)
                        If ref <> "" Then editRefs(ref) = pno
                    End If
                Next
            End If
        Case "ACCESS"
            id = Trim$(GetVal(MakeCtx(sb("ATTR")), "SQLID"))
            If id = "" Then id = Trim$(GetVal(MakeCtx(sb("ATTR")), "アクセス名"))
            If IsBundled(id) Then res.Add Array("まとめ書き", id, "アクセス定義でSQLがまとめられている", GetVal(MakeCtx(sb("ATTR")), "処理No"), "SQLごとにブロックを分ける")
            accIds(id) = 1
        Case "EDIT"
            editNames(Trim$(GetVal(MakeCtx(sb("ATTR")), "編集名"))) = 1
        End Select
    Next

    ' --- 突き合わせ ---
    For Each k In methods.Keys
        If Not procM.Exists(k) Then
            res.Add Array("メソッド抜け", k, "処理定義にブロックがない（" & methods(k) & "）", "", "P1xx として追加")
        End If
    Next
    For Each k In procM.Keys
        If Not methods.Exists(k) Then res.Add Array("メソッド不明", k, "ソースに無いメソッド名", procM(k), "メソッド名の誤記を確認")
    Next
    For Each k In sqlIds.Keys
        If Not procSql.Exists(k) Then res.Add Array("SQL抜け(処理)", k, "処理定義にSQL行がない", "", "呼出位置にSQL行を追加")
        If Not accIds.Exists(k) Then res.Add Array("SQL抜け(アクセス)", k, "アクセス定義にブロックがない", "", "ACCESSブロックを追加")
    Next
    For Each k In procSql.Keys
        If Not sqlIds.Exists(k) And Not IsBundled(CStr(k)) Then res.Add Array("SQL不明", k, "ソースに見つからないSQL ID", procSql(k), "IDの誤記を確認")
    Next
    For Each k In editRefs.Keys
        If Not editNames.Exists(k) Then res.Add Array("編集定義抜け", k, "処理定義の備考で参照しているが編集定義がない", editRefs(k), "同じ@編集名のEDITを追加")
    Next
    If pubs.Count > 1 Then
        For Each k In pubs.Keys
            res.Add Array("public確認", k, "publicメソッド（入口は1つのはず）", IIf(procM.Exists(k), procM(k), ""), "入口以外ならprivateの誤りの可能性")
        Next
    End If
    If sqlIds.Count = 0 Then res.Add Array("設定確認", "SQL IDパターン", "ソースからSQL IDが1つも見つからない", "", "設定シートのSQL IDパターンを確認")

    WriteCheckSheet res, CStr(imp), src, methods.Count, sqlIds.Count, procM.Count, procSql.Count
    EndRun
    ThisWorkbook.Worksheets("突合結果").Activate
    MsgBox "突合が終わりました。指摘 " & res.Count & " 件" & vbLf & "「突合結果」シートを確認してください。", vbInformation
    Exit Sub
EH:
    EndRun
    MsgBox "突合中にエラー: " & Err.Description, vbExclamation
End Sub

' Javaソースからメソッド名（可視性）とSQL ID（文字列リテラル）を抜き出す
Private Sub JavaInventory(ByVal t As String, ByRef methods As Object, ByRef pubs As Object, ByRef sqlIds As Object)
    Dim re As Object, m As Object, nm As String
    Set methods = NewDict(): Set pubs = NewDict(): Set sqlIds = NewDict()
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True
    re.Pattern = "^[ \t]*(public|protected|private)[ \t]+(?:(?:static|final|synchronized|abstract)[ \t]+)*[\w<>\[\]\.,? ]+?[ \t]+(\w+)[ \t]*\("
    For Each m In re.Execute(t)
        nm = m.SubMatches(1)
        If InStr("|if|for|while|switch|catch|return|new|class|interface|enum|", "|" & nm & "|") = 0 Then
            If Not methods.Exists(nm) Then methods(nm) = m.SubMatches(0)
            If m.SubMatches(0) = "public" Then pubs(nm) = 1
        End If
    Next
    re.MultiLine = False
    re.Pattern = """(" & ConfStr("SQL IDパターン", "[a-z]+[0-9]+_[A-Za-z0-9_]+") & ")"""
    For Each m In re.Execute(t)
        sqlIds(m.SubMatches(0)) = 1
    Next
End Sub

Private Function IsBundled(ByVal s As String) As Boolean
    IsBundled = (InStr(s, "／") > 0 Or InStr(s, "/") > 0 Or InStr(s, "・") > 0 Or InStr(s, "、") > 0)
End Function

' 備考「シート「編集定義」参照\n　名称」から名称を取り出す
Private Function RefName(ByVal note As String) As String
    Dim p As Variant, s As String
    p = Split(note, vbLf)
    s = p(UBound(p))
    If UBound(p) = 0 Then
        If InStr(s, "参照") > 0 Then s = Mid$(s, InStr(s, "参照") + 2)
    End If
    s = Replace(s, "　", " ")
    RefName = Trim$(s)
End Function

Private Sub WriteCheckSheet(res As Collection, ByVal imp As String, src As Variant, _
                            ByVal nM As Long, ByVal nS As Long, ByVal pM As Long, ByVal pS As Long)
    Dim ws As Worksheet, r As Long, it As Variant, f As Variant, req As String, srcNames As String
    Set ws = ThisWorkbook.Worksheets("突合結果")
    ws.Cells.Clear
    For Each f In src
        srcNames = srcNames & Mid$(f, InStrRev(f, "\") + 1) & " "
    Next
    ws.Range("A1").Value = "ソース突合結果"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "取込ファイル: " & imp
    ws.Range("A3").Value = "ソース: " & srcNames
    ws.Range("A4").Value = "ソース: メソッド " & nM & " / SQL ID " & nS & "　｜　取込ファイル: 処理定義ブロック " & pM & " / SQL行(種類) " & pS & "　｜　指摘 " & res.Count & " 件"
    r = 6
    ws.Range("A6:E6").Value = Array("区分", "対象", "内容", "処理No", "対応")
    ws.Range("A6:E6").Font.Bold = True
    ws.Range("A6:E6").Interior.Color = RGB(183, 222, 232)
    For Each it In res
        r = r + 1
        ws.Cells(r, 1).Resize(1, 5).Value = it
        If it(0) <> "public確認" And it(0) <> "設定確認" Then
            req = req & "- " & it(0) & "：" & it(1) & "（" & it(2) & "）" & vbLf
        End If
    Next
    If r = 6 Then ws.Cells(7, 1).Value = "指摘なし"
    r = r + 2
    ws.Cells(r, 1).Value = "▼ Copilotへの依頼文（セルをコピーして /design-reverse-fix またはチャットに貼る）"
    ws.Cells(r, 1).Font.Bold = True
    If req = "" Then req = "（抜けなし）"
    ws.Cells(r + 1, 1).Value = "取込ファイル " & Mid$(imp, InStrRev(imp, "\") + 1) & " について、ソースとの突合で次の抜け・誤りが見つかりました。" & _
        "取込フォーマットの規則に従い、該当するブロックを追記・修正してください（既存の正しいブロックは変更しない）。" & vbLf & req
    ws.Cells(r + 1, 1).WrapText = False
    ws.Columns("A:E").VerticalAlignment = xlTop
End Sub
