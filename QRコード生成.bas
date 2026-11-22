Attribute VB_Name = "QRCode"
'==============================================================================
'  QRコード生成マクロ（外部ライブラリ・インターネット接続 不要）
'  - 日本語を含む文字列に対応（UTF-8 バイトモード）
'  - バージョン（サイズ）とマスクは自動選択
'  - QRコードは図形（四角形のグループ）としてシートに描画されます
'
'  使い方：
'   1) Alt+F11 で VBE を開き、[ファイル]→[ファイルのインポート] でこの .bas を読み込む
'   2) 下の「設定」を必要に応じて変更
'   3) Alt+F8 から「QRコードを作成」または「選択セルをQRコード化」を実行
'   ※ 保存はマクロ有効ブック（.xlsm）で
'==============================================================================
Option Explicit

'---------------------------- 設定 --------------------------------------------
Private Const SRC_CELL As String = "A1"     ' QR化する文字列が入っているセル
Private Const DEST_CELL As String = "C1"    ' QRコードを置くセル（左上）
Private Const QR_SIZE_PT As Double = 120    ' QRコードの一辺の大きさ（ポイント）
Private Const QR_ECC As String = "M"        ' 誤り訂正レベル L / M / Q / H
'------------------------------------------------------------------------------

Private ECC_CW(0 To 3, 0 To 40) As Long
Private NBLK(0 To 3, 0 To 40) As Long
Private P2(0 To 31) As Long
Private M() As Byte      ' モジュール（1=黒）  M(y, x)
Private F() As Byte      ' 機能パターン領域フラグ
Private QSize As Long
Private QEcl As Long
Private Inited As Boolean

'==============================================================================
'  実行用マクロ
'==============================================================================

' 設定したセル（SRC_CELL）の文字列を QR コード化して DEST_CELL に配置
Public Sub QRコードを作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim txt As String: txt = CStr(ws.Range(SRC_CELL).Value)
    If Len(txt) = 0 Then
        MsgBox SRC_CELL & " に文字列が入力されていません。", vbExclamation
        Exit Sub
    End If
    Dim d As Range: Set d = ws.Range(DEST_CELL)
    MakeQRShape ws, txt, d.Left, d.Top, QR_SIZE_PT, QR_ECC, "QR_" & Replace(SRC_CELL, "$", "")
End Sub

' 選択中の各セルの文字列を QR コード化し、そのセルの右隣に配置
Public Sub 選択セルをQRコード化()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range, ws As Worksheet, n As Long
    Set ws = ActiveSheet
    For Each c In Selection.Cells
        If Len(CStr(c.Value)) > 0 Then
            MakeQRShape ws, CStr(c.Value), c.Offset(0, 1).Left + 2, c.Offset(0, 1).Top + 2, _
                        QR_SIZE_PT, QR_ECC, "QR_" & c.Address(False, False)
            n = n + 1
        End If
    Next
    If n = 0 Then MsgBox "文字列が入ったセルを選択してください。", vbExclamation
End Sub

' シート上のこのマクロで作った QR コードをすべて削除
Public Sub QRコードを全削除()
    Dim i As Long, ws As Worksheet: Set ws = ActiveSheet
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name Like "QR_*" Then ws.Shapes(i).Delete
    Next
End Sub

'==============================================================================
'  QR コードを図形として描画（他のマクロから呼び出し可）
'==============================================================================
Public Function MakeQRShape(ws As Worksheet, ByVal text As String, ByVal leftPt As Double, _
        ByVal topPt As Double, ByVal sizePt As Double, Optional ByVal ecc As String = "M", _
        Optional ByVal shapeName As String = "") As Shape
    Dim e As Long
    Select Case UCase$(ecc)
        Case "L": e = 0
        Case "Q": e = 2
        Case "H": e = 3
        Case Else: e = 1
    End Select

    If Not EncodeQR(text, e) Then
        MsgBox "文字列が長すぎて QR コードにできません。（" & Len(text) & " 文字）", vbExclamation
        Exit Function
    End If

    Dim scrUpd As Boolean: scrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False

    If Len(shapeName) > 0 Then
        On Error Resume Next
        ws.Shapes(shapeName).Delete
        On Error GoTo 0
    End If

    Const QUIET As Long = 4
    Dim px As Double: px = sizePt / (QSize + QUIET * 2)
    Dim names() As Variant, cnt As Long
    ReDim names(0 To QSize * QSize \ 2 + 2)

    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(1, leftPt, topPt, sizePt, sizePt)
    bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.Visible = msoFalse
    names(cnt) = bg.Name: cnt = cnt + 1

    Dim x As Long, y As Long, x0 As Long, s As Shape
    For y = 0 To QSize - 1
        x = 0
        Do While x < QSize
            If M(y, x) = 1 Then
                x0 = x
                Do While x < QSize
                    If M(y, x) = 0 Then Exit Do
                    x = x + 1
                Loop
                ' 隙間防止にわずかに重ねて描画
                Set s = ws.Shapes.AddShape(1, leftPt + (x0 + QUIET) * px, topPt + (y + QUIET) * px, _
                                           (x - x0) * px + 0.05, px + 0.05)
                s.Fill.ForeColor.RGB = RGB(0, 0, 0)
                s.Line.Visible = msoFalse
                If cnt > UBound(names) Then ReDim Preserve names(0 To cnt * 2)
                names(cnt) = s.Name: cnt = cnt + 1
            Else
                x = x + 1
            End If
        Loop
    Next
    ReDim Preserve names(0 To cnt - 1)

    Dim g As Shape
    Set g = ws.Shapes.Range(names).Group
    If Len(shapeName) > 0 Then g.Name = shapeName
    g.LockAspectRatio = msoTrue
    g.AlternativeText = text
    Set MakeQRShape = g

    Application.ScreenUpdating = scrUpd
End Function

'==============================================================================
'  以下、QR コードのエンコード処理
'==============================================================================
Private Sub InitTables()
    If Inited Then Exit Sub
    Dim rowsE, rowsB, i As Long, j As Long, a
    rowsE = Array( _
      "-1,7,10,15,20,26,18,20,24,30,18,20,24,26,30,22,24,28,30,28,28,28,28,30,30,26,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30", _
      "-1,10,16,26,18,24,16,18,22,22,26,30,22,22,24,24,28,28,26,26,26,26,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28", _
      "-1,13,22,18,26,18,24,18,22,20,24,28,26,24,20,30,24,28,28,26,30,28,30,30,30,30,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30", _
      "-1,17,28,22,16,22,28,26,26,24,28,24,28,22,24,24,30,28,28,26,28,30,24,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30")
    rowsB = Array( _
      "-1,1,1,1,1,1,2,2,2,2,4,4,4,4,4,6,6,6,6,7,8,8,9,9,10,12,12,12,13,14,15,16,17,18,19,19,20,21,22,24,25", _
      "-1,1,1,1,2,2,4,4,4,5,5,5,8,9,9,10,10,11,13,14,16,17,17,18,20,21,23,25,26,28,29,31,33,35,37,38,40,43,45,47,49", _
      "-1,1,1,2,2,4,4,6,6,8,8,8,10,12,16,12,17,16,18,21,20,23,23,25,27,29,34,34,35,38,40,43,45,48,51,53,56,59,62,65,68", _
      "-1,1,1,2,4,4,4,5,6,8,8,11,11,16,16,18,16,19,21,25,25,25,34,30,32,35,37,40,42,45,48,51,54,57,60,63,66,70,74,77,81")
    For i = 0 To 3
        a = Split(rowsE(i), ",")
        For j = 0 To 40: ECC_CW(i, j) = CLng(a(j)): Next
        a = Split(rowsB(i), ",")
        For j = 0 To 40: NBLK(i, j) = CLng(a(j)): Next
    Next
    P2(0) = 1
    For i = 1 To 30: P2(i) = P2(i - 1) * 2: Next
    Inited = True
End Sub

Private Function Bit(ByVal v As Long, ByVal i As Long) As Long
    Bit = (v \ P2(i)) And 1
End Function

' 文字列 → UTF-8 バイト列
Private Function Utf8Bytes(ByVal s As String, ByRef n As Long) As Long()
    Dim b() As Long, i As Long, c As Long, lo As Long
    ReDim b(0 To Len(s) * 4 + 1)
    n = 0
    i = 1
    Do While i <= Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < Len(s) Then
            lo = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * &H400& + (lo - &HDC00&)
                i = i + 1
            End If
        End If
        If c < &H80& Then
            b(n) = c: n = n + 1
        ElseIf c < &H800& Then
            b(n) = &HC0& Or (c \ &H40&): b(n + 1) = &H80& Or (c And &H3F&): n = n + 2
        ElseIf c < &H10000 Then
            b(n) = &HE0& Or (c \ &H1000&)
            b(n + 1) = &H80& Or ((c \ &H40&) And &H3F&)
            b(n + 2) = &H80& Or (c And &H3F&): n = n + 3
        Else
            b(n) = &HF0& Or (c \ &H40000)
            b(n + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
            b(n + 2) = &H80& Or ((c \ &H40&) And &H3F&)
            b(n + 3) = &H80& Or (c And &H3F&): n = n + 4
        End If
        i = i + 1
    Loop
    Utf8Bytes = b
End Function

Private Function RawModules(ByVal v As Long) As Long
    Dim r As Long, na As Long
    r = (16 * v + 128) * v + 64
    If v >= 2 Then
        na = v \ 7 + 2
        r = r - ((25 * na - 10) * na - 55)
        If v >= 7 Then r = r - 36
    End If
    RawModules = r
End Function

Private Function DataCW(ByVal v As Long, ByVal e As Long) As Long
    DataCW = RawModules(v) \ 8 - ECC_CW(e, v) * NBLK(e, v)
End Function

Private Function GfMul(ByVal x As Long, ByVal y As Long) As Long
    Dim z As Long, i As Long
    For i = 7 To 0 Step -1
        z = ((z * 2) Xor (((z \ 128) And 1) * &H11D&)) And &HFF&
        z = z Xor (((y \ P2(i)) And 1) * x)
    Next
    GfMul = z
End Function

Private Function RsDivisor(ByVal deg As Long) As Long()
    Dim r() As Long, root As Long, i As Long, j As Long
    ReDim r(0 To deg - 1)
    r(deg - 1) = 1: root = 1
    For i = 0 To deg - 1
        For j = 0 To deg - 1
            r(j) = GfMul(r(j), root)
            If j + 1 < deg Then r(j) = r(j) Xor r(j + 1)
        Next
        root = GfMul(root, 2)
    Next
    RsDivisor = r
End Function

Private Function RsRemainder(data() As Long, ByVal st As Long, ByVal ln As Long, dv() As Long) As Long()
    Dim deg As Long, res() As Long, k As Long, i As Long, f As Long
    deg = UBound(dv) + 1
    ReDim res(0 To deg - 1)
    For k = st To st + ln - 1
        f = data(k) Xor res(0)
        For i = 0 To deg - 2: res(i) = res(i + 1): Next
        res(deg - 1) = 0
        For i = 0 To deg - 1: res(i) = res(i) Xor GfMul(dv(i), f): Next
    Next
    RsRemainder = res
End Function

Private Sub SetF(ByVal x As Long, ByVal y As Long, ByVal dark As Boolean)
    M(y, x) = IIf(dark, 1, 0)
    F(y, x) = 1
End Sub

Private Function EncodeQR(ByVal text As String, ByVal e As Long) As Boolean
    InitTables
    QEcl = e
    Dim data() As Long, nData As Long, v As Long, capBits As Long, ccBits As Long
    data = Utf8Bytes(text, nData)

    For v = 1 To 40
        ccBits = IIf(v <= 9, 8, 16)
        capBits = DataCW(v, e) * 8
        If 4 + ccBits + 8 * nData <= capBits Then Exit For
    Next
    If v > 40 Then EncodeQR = False: Exit Function

    ' --- ビット列作成 ---
    Dim bits() As Byte, nb As Long, i As Long, j As Long, pad As Long
    ReDim bits(0 To capBits - 1)
    AppendBits bits, nb, 4, 4
    AppendBits bits, nb, nData, ccBits
    For i = 0 To nData - 1: AppendBits bits, nb, data(i), 8: Next
    AppendBits bits, nb, 0, IIf(capBits - nb < 4, capBits - nb, 4)
    AppendBits bits, nb, 0, (8 - nb Mod 8) Mod 8
    pad = &HEC&
    Do While nb < capBits
        AppendBits bits, nb, pad, 8
        pad = pad Xor &HFD&
    Loop

    Dim nCW As Long: nCW = capBits \ 8
    Dim cw() As Long: ReDim cw(0 To nCW - 1)
    For i = 0 To nCW - 1
        For j = 0 To 7: cw(i) = cw(i) * 2 + bits(i * 8 + j): Next
    Next

    ' --- ブロック分割・誤り訂正符号付加・インターリーブ ---
    Dim nBlocks As Long, ec As Long, rawCW As Long, nShort As Long, shortLen As Long
    nBlocks = NBLK(e, v): ec = ECC_CW(e, v): rawCW = RawModules(v) \ 8
    nShort = nBlocks - rawCW Mod nBlocks
    shortLen = rawCW \ nBlocks
    Dim dv() As Long: dv = RsDivisor(ec)
    Dim blk() As Long: ReDim blk(0 To nBlocks - 1, 0 To shortLen)
    Dim ecw() As Long, k As Long, dl As Long, p As Long
    For i = 0 To nBlocks - 1
        dl = shortLen - ec + IIf(i < nShort, 0, 1)
        For j = 0 To dl - 1: blk(i, j) = cw(k + j): Next
        ecw = RsRemainder(cw, k, dl, dv)
        k = k + dl
        p = dl + IIf(i < nShort, 1, 0)
        For j = 0 To ec - 1: blk(i, p + j) = ecw(j): Next
    Next
    Dim outCW() As Long, nOut As Long: ReDim outCW(0 To rawCW - 1)
    For i = 0 To shortLen
        For j = 0 To nBlocks - 1
            If i <> shortLen - ec Or j >= nShort Then
                outCW(nOut) = blk(j, i): nOut = nOut + 1
            End If
        Next
    Next

    ' --- 機能パターン描画 ---
    QSize = v * 4 + 17
    ReDim M(0 To QSize - 1, 0 To QSize - 1)
    ReDim F(0 To QSize - 1, 0 To QSize - 1)
    For i = 0 To QSize - 1
        SetF 6, i, (i Mod 2 = 0)
        SetF i, 6, (i Mod 2 = 0)
    Next
    DrawFinder 3, 3
    DrawFinder QSize - 4, 3
    DrawFinder 3, QSize - 4

    If v > 1 Then
        Dim na As Long, stp As Long, pos() As Long, dx As Long, dy As Long
        na = v \ 7 + 2
        If v = 32 Then stp = 26 Else stp = (v * 4 + na * 2 + 1) \ (na * 2 - 2) * 2
        ReDim pos(0 To na - 1)
        pos(0) = 6
        p = QSize - 7
        For i = na - 1 To 1 Step -1: pos(i) = p: p = p - stp: Next
        For i = 0 To na - 1
            For j = 0 To na - 1
                If Not ((i = 0 And j = 0) Or (i = 0 And j = na - 1) Or (i = na - 1 And j = 0)) Then
                    For dy = -2 To 2
                        For dx = -2 To 2
                            SetF pos(i) + dx, pos(j) + dy, (MaxL(Abs(dx), Abs(dy)) <> 1)
                        Next
                    Next
                End If
            Next
        Next
    End If

    DrawFormat 0
    If v >= 7 Then
        Dim r As Long, b As Long, bt As Boolean
        r = v
        For i = 1 To 12: r = (r * 2) Xor ((r \ 2048) * &H1F25&): Next
        b = v * 4096 Or r
        For i = 0 To 17
            bt = (Bit(b, i) = 1)
            SetF QSize - 11 + i Mod 3, i \ 3, bt
            SetF i \ 3, QSize - 11 + i Mod 3, bt
        Next
    End If

    ' --- データ配置 ---
    Dim rightC As Long, vert As Long, x As Long, y As Long, up As Boolean, bi As Long
    rightC = QSize - 1
    Do While rightC >= 1
        If rightC = 6 Then rightC = 5
        For vert = 0 To QSize - 1
            For j = 0 To 1
                x = rightC - j
                up = (((rightC + 1) And 2) = 0)
                If up Then y = QSize - 1 - vert Else y = vert
                If F(y, x) = 0 And bi < nOut * 8 Then
                    M(y, x) = (outCW(bi \ 8) \ P2(7 - (bi And 7))) And 1
                    bi = bi + 1
                End If
            Next
        Next
        rightC = rightC - 2
    Loop

    ' --- マスク選択 ---
    Dim mk As Long, best As Long, bestP As Long, pen As Long
    bestP = -1
    For mk = 0 To 7
        ApplyMask mk
        DrawFormat mk
        pen = Penalty()
        If bestP < 0 Or pen < bestP Then bestP = pen: best = mk
        ApplyMask mk
    Next
    ApplyMask best
    DrawFormat best
    EncodeQR = True
End Function

Private Sub AppendBits(bits() As Byte, nb As Long, ByVal v As Long, ByVal n As Long)
    Dim i As Long
    For i = n - 1 To 0 Step -1
        bits(nb) = (v \ P2(i)) And 1
        nb = nb + 1
    Next
End Sub

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

Private Sub DrawFinder(ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long, d As Long, xx As Long, yy As Long
    For dy = -4 To 4
        For dx = -4 To 4
            d = MaxL(Abs(dx), Abs(dy))
            xx = cx + dx: yy = cy + dy
            If xx >= 0 And xx < QSize And yy >= 0 And yy < QSize Then SetF xx, yy, (d <> 2 And d <> 4)
        Next
    Next
End Sub

Private Sub DrawFormat(ByVal mask As Long)
    Dim fb As Long, d As Long, r As Long, b As Long, i As Long
    fb = Choose(QEcl + 1, 1, 0, 3, 2)
    d = fb * 8 Or mask
    r = d
    For i = 1 To 10: r = (r * 2) Xor ((r \ 512) * &H537&): Next
    b = (d * 1024 Or r) Xor &H5412&
    For i = 0 To 5: SetF 8, i, Bit(b, i) = 1: Next
    SetF 8, 7, Bit(b, 6) = 1
    SetF 8, 8, Bit(b, 7) = 1
    SetF 7, 8, Bit(b, 8) = 1
    For i = 9 To 14: SetF 14 - i, 8, Bit(b, i) = 1: Next
    For i = 0 To 7: SetF QSize - 1 - i, 8, Bit(b, i) = 1: Next
    For i = 8 To 14: SetF 8, QSize - 15 + i, Bit(b, i) = 1: Next
    SetF 8, QSize - 8, True
End Sub

Private Sub ApplyMask(ByVal mk As Long)
    Dim x As Long, y As Long, c As Boolean
    For y = 0 To QSize - 1
        For x = 0 To QSize - 1
            If F(y, x) = 0 Then
                Select Case mk
                    Case 0: c = ((x + y) Mod 2 = 0)
                    Case 1: c = (y Mod 2 = 0)
                    Case 2: c = (x Mod 3 = 0)
                    Case 3: c = ((x + y) Mod 3 = 0)
                    Case 4: c = ((x \ 3 + y \ 2) Mod 2 = 0)
                    Case 5: c = ((x * y) Mod 2 + (x * y) Mod 3 = 0)
                    Case 6: c = (((x * y) Mod 2 + (x * y) Mod 3) Mod 2 = 0)
                    Case 7: c = (((x + y) Mod 2 + (x * y) Mod 3) Mod 2 = 0)
                End Select
                If c Then M(y, x) = M(y, x) Xor 1
            End If
        Next
    Next
End Sub

Private Function Penalty() As Long
    Dim res As Long, a As Long, b As Long, t As Long, s As String, run As Long, p As Long
    Dim buf() As String: ReDim buf(0 To QSize - 1)
    For t = 0 To 1
        For a = 0 To QSize - 1
            For b = 0 To QSize - 1
                If t = 0 Then buf(b) = CStr(M(a, b)) Else buf(b) = CStr(M(b, a))
            Next
            s = Join(buf, "")
            run = 1
            For b = 2 To QSize
                If Mid$(s, b, 1) = Mid$(s, b - 1, 1) Then
                    run = run + 1
                    If run = 5 Then
                        res = res + 3
                    ElseIf run > 5 Then
                        res = res + 1
                    End If
                Else
                    run = 1
                End If
            Next
            p = InStr(1, s, "10111010000")
            Do While p > 0: res = res + 40: p = InStr(p + 1, s, "10111010000"): Loop
            p = InStr(1, s, "00001011101")
            Do While p > 0: res = res + 40: p = InStr(p + 1, s, "00001011101"): Loop
        Next
    Next
    Dim x As Long, y As Long, dark As Long, tot As Long
    For y = 0 To QSize - 2
        For x = 0 To QSize - 2
            If M(y, x) = M(y, x + 1) And M(y, x) = M(y + 1, x) And M(y, x) = M(y + 1, x + 1) Then res = res + 3
        Next
    Next
    For y = 0 To QSize - 1
        For x = 0 To QSize - 1: dark = dark + M(y, x): Next
    Next
    tot = QSize * QSize
    res = res + ((Abs(dark * 20 - tot * 10) + tot - 1) \ tot - 1) * 10
    Penalty = res
End Function
