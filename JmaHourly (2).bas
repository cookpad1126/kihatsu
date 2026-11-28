Attribute VB_Name = "JmaHourly"
'==========================================================================
' 気象庁「過去の気象データ」1時間値（hourly_s1）→ 月別Excelファイル（地点ごとのシート）
'
'   ■ 実行マクロ
'     GetJmaHourly   … 期間シートの開始日～終了日を取得 → 月別ファイルへ書き出し → 書き込み確認
'     CheckJmaData   … 月別ファイルを開いて、期間内の各地点・各時刻のデータが埋まっているか確認
'                      （未書き込みの日があれば、その日だけ再取得するか確認します）
'
'   ■ このブックに必要なシート
'     「気象取得」     … 取得期間（開始日・終了日）を入力するシート（無ければ自動作成）
'     「実績地点設定」 … 地点名 / prec_no / block_no / 代替地点1～3（無ければ自動作成）
'     「取得チェック」 … 確認結果（自動作成）
'
'   ■ 欠損の補完（USE_SUBSTITUTE = True のとき）
'     代替地点1 → 代替地点2 → 前後の時刻から直線補間（INTERP_MAX_HOURS 以内） → 代替地点3
'     降雪・積雪は寒候年ごとの運用期間内の欠測だけ代替地点1→2→3で補完
'     代替地点の値＝オレンジ、補間値＝緑で着色し、備考列に「補完：気温←呉」などを表示
'
'   参照設定不要。HTML解析は文字列処理のみ
'==========================================================================
Option Explicit

'##########################################################################
'#  設定（ここだけ変更してください）
'##########################################################################

'---- ① 取得期間：このブックのシート上のセル ------------------------------
Private Const PERIOD_SHEET As String = "気象取得"    ' 期間を入力するシート名
Private Const START_CELL As String = "B2"            ' 開始日のセル
Private Const END_CELL As String = "B3"              ' 終了日のセル

'---- ② 書き出し先ファイル ------------------------------------------------
Private Const OUTPUT_FOLDER As String = ""           ' 書き出し先フォルダ（空欄＝このブックと同じフォルダ）
Private Const FILE_NAME_PATTERN As String = "気象実績_{yyyy}{mm}.xlsx"   ' {yyyy}{mm} が年・月に置き換わる
Private Const TEMPLATE_FILE As String = ""           ' 新しい月のファイルを作るときのひな形（フルパス。空欄＝白紙）
Private Const ADD_PREV_LAST_DAY As Boolean = True    ' True：各月ファイルの先頭に「前月の最終日」も書く
Private Const ADD_NEXT_FIRST_DAY As Boolean = True   ' True：各月ファイルの最後に「翌月の1日」も書く
Private Const HEADER_ROW As Long = 1                 ' 見出しの行
Private Const DATA_START_ROW As Long = 2             ' データの開始行（1日＝24行、日付順に固定位置）
Private Const WRITE_HEADER As Boolean = True         ' True：見出しセルが空欄なら見出しを書く

'---- ③ 時刻の書き方 ------------------------------------------------------
Private Const TIME_MODE As Long = 1
    ' 1 = 日時を1列（COL_DATETIME）       例 2026/11/23 01:00
    ' 2 = 日付列＋時刻列（COL_DATE, COL_TIME） 例 2026/11/23 | 01:00
    ' 3 = 年・月・日・時・分・秒を別々の列（COL_YEAR～COL_SECOND。不要な列は空欄）
Private Const TIME_BASIS As Long = 1
    ' 1 = 観測時刻（気象庁の表示どおり 1時～24時。値はその時刻までの1時間）
    ' 2 = 開始時刻（0時～23時。0:00～1:00 の値を 0時 と表示）
Private Const USE_24H As Boolean = False
    ' TIME_BASIS=1 のときの24時の書き方。True：当日の「24時」（モード2は 24:00、モード3は 時=24）
    '                                     False：翌日の 0時
    '   ※モード1（日時1列）は Excel の仕様で24時を表せないため常に翌日 0:00
Private Const DATETIME_FORMAT As String = "yyyy/mm/dd hh:mm"
Private Const DATE_FORMAT As String = "yyyy/mm/dd"
Private Const TIME_FORMAT As String = "[h]:mm"        ' [h] にすると 24:00 も表示できる

'---- ④ 列の割り当て（"A","B"…。空欄 "" の項目は書き出さない）------------
'   時刻（TIME_MODE に合う列だけ使われます）
Private Const COL_DATETIME As String = "A"           ' モード1
Private Const COL_DATE As String = ""                ' モード2
Private Const COL_TIME As String = ""                ' モード2
Private Const COL_YEAR As String = ""                ' モード3
Private Const COL_MONTH As String = ""               ' モード3
Private Const COL_DAY As String = ""                 ' モード3
Private Const COL_HOUR As String = ""                ' モード3
Private Const COL_MINUTE As String = ""              ' モード3
Private Const COL_SECOND As String = ""              ' モード3
'   気象項目
Private Const COL_TEMP As String = "B"               ' 気温(℃)
Private Const COL_HUM As String = "C"                ' 湿度(%)
Private Const COL_PREC As String = "D"               ' 降水量(mm)
Private Const COL_WSPD As String = "E"               ' 風速(m/s)
Private Const COL_WDIR As String = "F"               ' 風向
Private Const COL_NW As String = "G"                 ' 北風成分(m/s)（北寄り＋、南寄り－）
Private Const COL_SUN As String = "H"                ' 日照時間(h)
Private Const COL_RAD As String = "I"                ' 全天日射量(MJ/㎡)
Private Const COL_DEW As String = "J"                ' 露点温度(℃)
Private Const COL_VAP As String = "K"                ' 蒸気圧(hPa)
Private Const COL_PLOC As String = "L"               ' 現地気圧(hPa)
Private Const COL_PSEA As String = "M"               ' 海面気圧(hPa)
Private Const COL_SNOWF As String = "N"              ' 降雪(cm)
Private Const COL_SNOWD As String = "O"              ' 積雪(cm)
Private Const COL_WX As String = "P"                 ' 天気
Private Const COL_CLOUD As String = "Q"              ' 雲量
Private Const COL_VIS As String = "R"                ' 視程(km)
Private Const COL_NOTE As String = "S"               ' 備考（補完の内容）
Private Const COL_STATION As String = ""             ' 地点名（シート名と同じ。必要なら列を指定）

'---- ⑤ 補完・確認・通信 --------------------------------------------------
Private Const USE_SUBSTITUTE As Boolean = True       ' 欠損を代替地点・補間で補完する
Private Const INTERP_MAX_HOURS As Double = 3         ' 補間する欠損の最大長さ（時間）
Private Const CHECK_ITEMS As String = "気温,湿度,降水量,風速"   ' 確認で「数値が入っているか」を見る項目（④の項目名）
Private Const WAIT_SEC As Long = 1                   ' 1ページ取得ごとの待機秒（サーバー負荷対策）
Private Const SUB_COLOR As Long = 10086143           ' 代替地点の値の色 RGB(255,230,153)
Private Const INT_COLOR As Long = 13561798           ' 補間値の色       RGB(198,239,206)

'---- ⑥ このブックのシート名 ----------------------------------------------
Private Const SET_SHEET As String = "実績地点設定"
Private Const CHECK_SHEET As String = "取得チェック"

'##########################################################################
'#  ここから下は通常変更不要
'##########################################################################

Private Const INTERP_MARK As String = "補間"

' 気象庁の1時間値の表（17列）の並び
Private Const J_HOUR = 1, J_PLOC = 2, J_PSEA = 3, J_PREC = 4, J_TEMP = 5, J_DEW = 6
Private Const J_VAP = 7, J_HUM = 8, J_WSPD = 9, J_WDIR = 10, J_SUN = 11, J_RAD = 12
Private Const J_SNOWF = 13, J_SNOWD = 14, J_WX = 15, J_CLOUD = 16, J_VIS = 17
Private Const J_COLS = 17
Private Const J_NAMES As String = "時,現地気圧,海面気圧,降水量,気温,露点温度,蒸気圧,湿度,風速,風向,日照時間,全天日射量,降雪,積雪,天気,雲量,視程"
Private Const SRC_NW = -1, SRC_NOTE = -2, SRC_STATION = -3   ' 項目の特殊コード

Private Const DIR16 As String = "北,北北東,北東,東北東,東,東南東,南東,南南東,南,南南西,南西,西南西,西,西北西,北西,北北西"
Private Const N_SUB = 3

Private gStep As String
Private gSubName() As String, gSubPrec() As String, gSubBlock() As String
Private gCache As Object
Private gSnowPass As Boolean
Private gSnowOK() As Boolean

'==========================================================================
' 実行マクロ
'==========================================================================
Public Sub GetJmaHourly()
    Dim d1 As Date, d2 As Date, msg As String
    If Not ReadPeriod(d1, d2) Then Exit Sub
    If Not ValidateSettings() Then Exit Sub
    If MsgBox(Format(d1, "yyyy/mm/dd") & " ～ " & Format(d2, "yyyy/mm/dd") & " の気象実績を取得して、月別ファイルに書き出します。" & _
              vbCrLf & "目安の所要時間：約 " & Format(StationCount() * (d2 - d1 + 1) * (WAIT_SEC + 0.7) / 60, "0.0") & " 分", _
              vbOKCancel + vbInformation) = vbCancel Then Exit Sub

    msg = RunCore(d1, d2, "")
    If msg = "#ERROR" Then Exit Sub
    CheckAndOffer d1, d2, msg
End Sub

Public Sub CheckJmaData()
    Dim d1 As Date, d2 As Date
    If Not ReadPeriod(d1, d2) Then Exit Sub
    If Not ValidateSettings() Then Exit Sub
    CheckAndOffer d1, d2, ""
End Sub

'==========================================================================
' 取得の本体：地点ごとに全期間を集め → 補完 → 月別ファイルへ書き出し
'   onlyStation：""＝全地点、地点名を指定するとその地点だけ
'   戻り値：注意事項のメッセージ（"#ERROR"＝中断）
'==========================================================================
Private Function RunCore(ByVal d1 As Date, ByVal d2 As Date, ByVal onlyStation As String) As String
    Dim names() As String, precs() As String, blocks() As String, weights() As Double, nPt As Long
    Dim nSlot As Long, maxGap As Long, k As Long, i As Long, off As Long, h As Long, c As Long
    Dim d As Date, ser() As Variant, flg() As String, dayRes As Variant
    Dim failed As Collection, fd As Variant, errMsg As String
    Dim serAll() As Variant, flgAll() As Variant, useK() As Boolean

    On Error GoTo ErrHandler
    gStep = "地点設定の読み込み"
    ReadSettings PrepareSettingSheet(), names, precs, blocks, weights, nPt
    If nPt = 0 Then Err.Raise vbObjectError + 10, , "「" & SET_SHEET & "」シートに地点がありません"

    Set gCache = CreateObject("Scripting.Dictionary")
    nSlot = (d2 - d1 + 1) * 24
    maxGap = CLng(INTERP_MAX_HOURS)
    ReDim serAll(1 To nPt): ReDim flgAll(1 To nPt): ReDim useK(1 To nPt)

    Application.ScreenUpdating = False
    For k = 1 To nPt
        If onlyStation = "" Or onlyStation = names(k) Then
            useK(k) = True
            ReDim ser(1 To nSlot, 1 To J_COLS)
            ReDim flg(1 To nSlot, 1 To J_COLS)
            Set failed = New Collection
            For i = 1 To nSlot: ser(i, J_HOUR) = ((i - 1) Mod 24) + 1: Next i

            ' 本来の地点 ＋ 代替1・2
            For d = d1 To d2
                Application.StatusBar = "取得中: " & names(k) & " " & Format(d, "yyyy/mm/dd")
                DoEvents
                gStep = names(k) & " " & Format(d, "yyyy/mm/dd")
                off = (d - d1) * 24
                On Error Resume Next
                dayRes = GetHourlyDay(precs(k), blocks(k), d)
                If Err.Number <> 0 Then
                    failed.Add Array(d, Err.Description)
                    Err.Clear
                    dayRes = Empty
                End If
                On Error GoTo ErrHandler
                If IsArray(dayRes) Then
                    For h = 1 To 24
                        For c = 2 To J_COLS: ser(off + h, c) = dayRes(h, c): Next c
                    Next h
                End If
                If USE_SUBSTITUTE Then FillHourlyAt ser, flg, k, d, 1, 2, off
            Next d

            If USE_SUBSTITUTE Then
                ' 補間 → 代替3
                gStep = names(k) & " 補間"
                InterpSeries ser, flg, nSlot, Array(J_PLOC, J_PSEA, J_TEMP, J_DEW, J_VAP, J_HUM), J_WSPD, J_WDIR, 0, maxGap
                For d = d1 To d2
                    Application.StatusBar = "補完中: " & names(k) & " " & Format(d, "yyyy/mm/dd")
                    DoEvents
                    off = (d - d1) * 24
                    If DayHasGap(ser, off) Then FillHourlyAt ser, flg, k, d, 3, 3, off
                Next d
                ' 降雪・積雪：運用期間内の欠測だけ代替1→2→3
                gStep = names(k) & " 降雪・積雪"
                BuildSnowWindow ser, nSlot, d1
                gSnowPass = True
                For d = d1 To d2
                    off = (d - d1) * 24
                    If AnyTrue(gSnowOK, off + 1, off + 24) Then FillHourlyAt ser, flg, k, d, 1, 3, off
                Next d
                gSnowPass = False
            End If

            serAll(k) = ser
            flgAll(k) = flg
            For Each fd In failed
                off = (fd(0) - d1) * 24
                If DayHasData(ser, off) Then
                    errMsg = errMsg & "・" & names(k) & " " & Format(fd(0), "yyyy/mm/dd") & "：取得できず、代替地点・補間で補完" & vbCrLf
                Else
                    errMsg = errMsg & "・" & names(k) & " " & Format(fd(0), "yyyy/mm/dd") & "：" & fd(1) & vbCrLf
                End If
            Next fd
        End If
    Next k

    gStep = "月別ファイルへの書き出し"
    WriteMonthlyFiles d1, d2, names, nPt, useK, serAll, flgAll

    Application.StatusBar = False
    Application.ScreenUpdating = True
    RunCore = errMsg
    Exit Function

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "処理: " & gStep & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & Err.Description, vbCritical
    RunCore = "#ERROR"
End Function

'==========================================================================
' 月別ファイルへの書き出し
'==========================================================================
' 取得した日 → 書き出すファイル（年月）の一覧。前月末日・翌月1日のオプションを反映
Private Function TargetFiles(ByVal d1 As Date, ByVal d2 As Date) As Object
    Dim dic As Object, d As Date
    Set dic = CreateObject("Scripting.Dictionary")
    For d = d1 To d2
        AddTarget dic, DateSerial(Year(d), Month(d), 1), d
        If ADD_PREV_LAST_DAY And Day(d + 1) = 1 Then AddTarget dic, DateSerial(Year(d), Month(d) + 1, 1), d
        If ADD_NEXT_FIRST_DAY And Day(d) = 1 Then AddTarget dic, DateSerial(Year(d), Month(d) - 1, 1), d
    Next d
    Set TargetFiles = dic
End Function

Private Sub AddTarget(dic As Object, ByVal m1 As Date, ByVal d As Date)
    Dim key As String
    key = Format(m1, "yyyymm")
    If Not dic.Exists(key) Then dic.Add key, New Collection
    dic(key).Add d
End Sub

' 月ファイル内の行番号（日付順に固定。先頭に前月末日を入れる場合は1日分ずれる）
Private Function RowOf(ByVal m1 As Date, ByVal d As Date, ByVal h As Long) As Long
    Dim fileStart As Date
    fileStart = m1
    If ADD_PREV_LAST_DAY Then fileStart = m1 - 1
    RowOf = DATA_START_ROW + CLng(d - fileStart) * 24 + (h - 1)
End Function

Private Sub WriteMonthlyFiles(ByVal d1 As Date, ByVal d2 As Date, names() As String, ByVal nPt As Long, _
                              useK() As Boolean, serAll() As Variant, flgAll() As Variant)
    Dim files As Object, key As Variant, m1 As Date, wb As Workbook, ws As Worksheet, k As Long
    Dim cols As Variant, srcs As Variant, hdrs As Variant, nItem As Long
    Dim dayList As Collection, dv As Variant, r1 As Long, r2 As Long, rr As Long
    Dim x As Long, h As Long, i As Long, colVals As Variant, v As Variant
    Dim ser As Variant, flg As Variant, fl As String, mSub As Collection, mInt As Collection
    Dim wasOpen As Boolean

    GetItemMap cols, srcs, hdrs, nItem
    Set files = TargetFiles(d1, d2)
    Application.DisplayAlerts = False

    For Each key In files.Keys
        m1 = DateSerial(CLng(Left$(key, 4)), CLng(Right$(key, 2)), 1)
        Application.StatusBar = "書き出し中: " & MonthFilePath(m1)
        DoEvents
        gStep = "書き出し " & MonthFilePath(m1)
        wasOpen = IsOpenBook(MonthFilePath(m1))
        Set wb = OpenMonthFile(m1, True)
        Set dayList = files(key)

        For k = 1 To nPt
            If useK(k) Then
                Set ws = GetOrAddSheetIn(wb, names(k))
                WriteHeaderIfEmpty ws, cols, hdrs, nItem
                ser = serAll(k): flg = flgAll(k)

                ' このファイルで書く行の範囲
                r1 = 0: r2 = 0
                For Each dv In dayList
                    rr = RowOf(m1, dv, 1)
                    If r1 = 0 Or rr < r1 Then r1 = rr
                    If rr + 23 > r2 Then r2 = rr + 23
                Next dv

                Set mSub = New Collection: Set mInt = New Collection
                For x = 1 To nItem
                    ' 既存の値を読み込み → 書く行だけ置き換え → まとめて書き戻す（間の行は保持）
                    colVals = ws.Range(cols(x) & r1 & ":" & cols(x) & r2).Value   ' 常に24行以上なので2次元配列
                    For Each dv In dayList
                        For h = 1 To 24
                            i = (dv - d1) * 24 + h
                            rr = RowOf(m1, dv, h)
                            Select Case srcs(x)
                                Case SRC_NW
                                    v = NorthComponent(CStr(ser(i, J_WDIR)), ser(i, J_WSPD)): fl = flg(i, J_WSPD)
                                Case SRC_NOTE
                                    v = NoteOf(flg, i): fl = ""
                                Case SRC_STATION
                                    v = names(k): fl = ""
                                Case Is > 0
                                    v = ser(i, srcs(x)): fl = flg(i, srcs(x))
                                Case Else                                   ' 時刻の列
                                    v = TimeCellValue(srcs(x), dv, h): fl = ""
                            End Select
                            If IsEmpty(v) Then v = ""
                            colVals(rr - r1 + 1, 1) = v
                            If fl <> "" Then
                                If fl = INTERP_MARK Then mInt.Add Array(rr, x) Else mSub.Add Array(rr, x)
                            End If
                        Next h
                    Next dv
                    ws.Range(cols(x) & r1 & ":" & cols(x) & r2).Value = colVals
                    FormatTimeColumn ws, cols(x), srcs(x), r1, r2
                    If srcs(x) > 0 Or srcs(x) = SRC_NW Then
                        ws.Range(cols(x) & r1 & ":" & cols(x) & r2).Interior.ColorIndex = xlNone
                    End If
                Next x
                PaintMarks ws, mSub, cols, SUB_COLOR
                PaintMarks ws, mInt, cols, INT_COLOR
            End If
        Next k

        wb.Save
        If Not wasOpen Then wb.Close SaveChanges:=False   ' もともと開いていたファイルは開いたまま
    Next key
    Application.DisplayAlerts = True
End Sub

' 書き出す項目の一覧（列・元データ・見出し）
'   時刻の列は srcs に負の大きな値（-101～-109）を入れる
Private Sub GetItemMap(cols As Variant, srcs As Variant, hdrs As Variant, nItem As Long)
    Dim c() As String, s() As Long, hd() As String
    ReDim c(1 To 40): ReDim s(1 To 40): ReDim hd(1 To 40)
    nItem = 0
    Select Case TIME_MODE
        Case 1
            AddItem c, s, hd, nItem, COL_DATETIME, -101, "日時"
        Case 2
            AddItem c, s, hd, nItem, COL_DATE, -102, "日付"
            AddItem c, s, hd, nItem, COL_TIME, -103, "時刻"
        Case Else
            AddItem c, s, hd, nItem, COL_YEAR, -104, "年"
            AddItem c, s, hd, nItem, COL_MONTH, -105, "月"
            AddItem c, s, hd, nItem, COL_DAY, -106, "日"
            AddItem c, s, hd, nItem, COL_HOUR, -107, "時"
            AddItem c, s, hd, nItem, COL_MINUTE, -108, "分"
            AddItem c, s, hd, nItem, COL_SECOND, -109, "秒"
    End Select
    AddItem c, s, hd, nItem, COL_STATION, SRC_STATION, "地点"
    AddItem c, s, hd, nItem, COL_TEMP, J_TEMP, "気温(℃)"
    AddItem c, s, hd, nItem, COL_HUM, J_HUM, "湿度(%)"
    AddItem c, s, hd, nItem, COL_PREC, J_PREC, "降水量(mm)"
    AddItem c, s, hd, nItem, COL_WSPD, J_WSPD, "風速(m/s)"
    AddItem c, s, hd, nItem, COL_WDIR, J_WDIR, "風向"
    AddItem c, s, hd, nItem, COL_NW, SRC_NW, "北風成分(m/s)"
    AddItem c, s, hd, nItem, COL_SUN, J_SUN, "日照時間(h)"
    AddItem c, s, hd, nItem, COL_RAD, J_RAD, "全天日射量(MJ/㎡)"
    AddItem c, s, hd, nItem, COL_DEW, J_DEW, "露点温度(℃)"
    AddItem c, s, hd, nItem, COL_VAP, J_VAP, "蒸気圧(hPa)"
    AddItem c, s, hd, nItem, COL_PLOC, J_PLOC, "現地気圧(hPa)"
    AddItem c, s, hd, nItem, COL_PSEA, J_PSEA, "海面気圧(hPa)"
    AddItem c, s, hd, nItem, COL_SNOWF, J_SNOWF, "降雪(cm)"
    AddItem c, s, hd, nItem, COL_SNOWD, J_SNOWD, "積雪(cm)"
    AddItem c, s, hd, nItem, COL_WX, J_WX, "天気"
    AddItem c, s, hd, nItem, COL_CLOUD, J_CLOUD, "雲量"
    AddItem c, s, hd, nItem, COL_VIS, J_VIS, "視程(km)"
    AddItem c, s, hd, nItem, COL_NOTE, SRC_NOTE, "備考"
    cols = c: srcs = s: hdrs = hd
End Sub

Private Sub AddItem(c() As String, s() As Long, hd() As String, nItem As Long, _
                    ByVal colLetter As String, ByVal src As Long, ByVal header As String)
    If Trim$(colLetter) = "" Then Exit Sub
    nItem = nItem + 1
    c(nItem) = UCase$(Trim$(colLetter)): s(nItem) = src: hd(nItem) = header
End Sub

' 時刻の列の値
Private Function TimeCellValue(ByVal code As Long, ByVal d As Date, ByVal h As Long) As Variant
    Dim dt As Date, dPart As Date, hr As Long, is24 As Boolean
    If TIME_BASIS = 2 Then
        dt = d + TimeSerial(h - 1, 0, 0)            ' 開始時刻 0～23時
    Else
        dt = d + TimeSerial(h, 0, 0)                ' 観測時刻 1～24時（24時＝翌日0時）
    End If
    is24 = (TIME_BASIS = 1 And h = 24 And USE_24H)
    If is24 Then dPart = d: hr = 24 Else dPart = Int(dt): hr = Hour(dt)

    Select Case code
        Case -101: TimeCellValue = dt
        Case -102: TimeCellValue = dPart
        Case -103: If is24 Then TimeCellValue = 1# Else TimeCellValue = TimeSerial(hr, 0, 0)
        Case -104: TimeCellValue = Year(dPart)
        Case -105: TimeCellValue = Month(dPart)
        Case -106: TimeCellValue = Day(dPart)
        Case -107: TimeCellValue = hr
        Case -108: TimeCellValue = 0
        Case -109: TimeCellValue = 0
    End Select
End Function

Private Sub FormatTimeColumn(ws As Worksheet, ByVal colLetter As String, ByVal code As Long, ByVal r1 As Long, ByVal r2 As Long)
    Dim rng As Range
    Set rng = ws.Range(colLetter & r1 & ":" & colLetter & r2)
    Select Case code
        Case -101: rng.NumberFormat = DATETIME_FORMAT
        Case -102: rng.NumberFormat = DATE_FORMAT
        Case -103: rng.NumberFormat = TIME_FORMAT
        Case -109 To -104: rng.NumberFormat = "0"
    End Select
End Sub

Private Sub WriteHeaderIfEmpty(ws As Worksheet, cols As Variant, hdrs As Variant, ByVal nItem As Long)
    Dim x As Long
    If Not WRITE_HEADER Then Exit Sub
    For x = 1 To nItem
        If CStr(ws.Range(cols(x) & HEADER_ROW).Value) = "" Then
            ws.Range(cols(x) & HEADER_ROW).Value = hdrs(x)
            ws.Range(cols(x) & HEADER_ROW).Font.Bold = True
        End If
    Next x
End Sub

Private Sub PaintMarks(ws As Worksheet, marks As Collection, cols As Variant, ByVal colorVal As Long)
    Dim rng As Range, m As Variant, n As Long
    For Each m In marks
        If rng Is Nothing Then
            Set rng = ws.Range(cols(m(1)) & m(0))
        Else
            Set rng = Union(rng, ws.Range(cols(m(1)) & m(0)))
        End If
        n = n + 1
        If n Mod 200 = 0 Then rng.Interior.Color = colorVal: Set rng = Nothing
    Next m
    If Not rng Is Nothing Then rng.Interior.Color = colorVal
End Sub

'---- 月ファイル ------------------------------------------------------------
Private Function MonthFilePath(ByVal m1 As Date) As String
    Dim folder As String, fname As String
    folder = OUTPUT_FOLDER
    If folder = "" Then folder = ThisWorkbook.Path
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    fname = Replace(Replace(FILE_NAME_PATTERN, "{yyyy}", Format(m1, "yyyy")), "{mm}", Format(m1, "mm"))
    MonthFilePath = folder & fname
End Function

' 月ファイルを開く（無ければ作成。createIfMissing=False なら Nothing を返す）
Private Function OpenMonthFile(ByVal m1 As Date, ByVal createIfMissing As Boolean, _
                               Optional ByVal readOnly As Boolean = False) As Workbook
    Dim path As String, wb As Workbook, fname As String
    path = MonthFilePath(m1)
    fname = Mid$(path, InStrRev(path, "\") + 1)
    For Each wb In Application.Workbooks           ' すでに開いている場合
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Set OpenMonthFile = wb: Exit Function
    Next wb
    If Dir(path) <> "" Then
        Set OpenMonthFile = Workbooks.Open(path, ReadOnly:=readOnly)
    ElseIf createIfMissing Then
        If TEMPLATE_FILE <> "" Then
            Set wb = Workbooks.Open(TEMPLATE_FILE)
        Else
            Set wb = Workbooks.Add(xlWBATWorksheet)
        End If
        wb.SaveAs path, FileFormat:=xlOpenXMLWorkbook
        Set OpenMonthFile = wb
    Else
        Set OpenMonthFile = Nothing
    End If
End Function

Private Function GetOrAddSheetIn(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        ' 白紙ブックの最初の空シートがあればそれを使う
        If wb.Worksheets.Count = 1 And Application.WorksheetFunction.CountA(wb.Worksheets(1).Cells) = 0 _
           And Left$(wb.Worksheets(1).Name, 5) = "Sheet" Then
            Set ws = wb.Worksheets(1)
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If
        ws.Name = nm
    End If
    Set GetOrAddSheetIn = ws
End Function

'==========================================================================
' 書き込み確認
'==========================================================================
Private Sub CheckAndOffer(ByVal d1 As Date, ByVal d2 As Date, ByVal fetchMsg As String)
    Dim missing As Object, nIssue As Long, nUnwritten As Long, msg As String, key As Variant
    Dim parts As Variant, dd As Variant, ranges As Collection, rg As Variant, again As String

    gStep = "書き込み確認"
    Set missing = CheckCore(d1, d2, nIssue, nUnwritten)

    msg = ""
    If fetchMsg <> "" Then msg = "【取得時の注意】" & vbCrLf & Left$(fetchMsg, 800) & vbCrLf
    If nIssue = 0 Then
        msg = msg & "【確認】期間内の全地点・全時刻のデータが埋まっています。"
        MsgBox msg, vbInformation
        Exit Sub
    End If
    msg = msg & "【確認】問題のある時刻が " & nIssue & " 件あります（「" & CHECK_SHEET & "」シート参照）。" & vbCrLf & _
          "　うち未書き込み（通信エラーなど）：" & nUnwritten & " 件"
    If missing.Count = 0 Then
        msg = msg & vbCrLf & "残りは気象庁側の欠測で、代替地点・補間でも埋まらなかった値です。"
        MsgBox msg, vbExclamation
        Exit Sub
    End If

    If MsgBox(msg & vbCrLf & vbCrLf & "未書き込みの日（" & missing.Count & " 地点・日）を再取得しますか？", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' 地点ごとに、連続した日をまとめて再取得
    again = ""
    For Each key In StationsOf(missing)
        Set ranges = DayRanges(missing, CStr(key))
        For Each rg In ranges
            again = again & RunCore(rg(0), rg(1), CStr(key))
            If Right$(again, 6) = "#ERROR" Then Exit Sub
        Next rg
    Next key
    CheckAndOffer d1, d2, again
End Sub

' 期間内の各地点・各時刻を確認。問題を「取得チェック」シートへ。
'   戻り値：未書き込みのある「地点|yyyymmdd」の一覧
Private Function CheckCore(ByVal d1 As Date, ByVal d2 As Date, nIssue As Long, nUnwritten As Long) As Object
    Dim names() As String, precs() As String, blocks() As String, weights() As Double, nPt As Long
    Dim cols As Variant, srcs As Variant, hdrs As Variant, nItem As Long
    Dim checkX() As Long, nChk As Long, timeX As Long, x As Long, item As Variant
    Dim files As Object, key As Variant, m1 As Date, wb As Workbook, ws As Worksheet, wasOpen As Boolean
    Dim dayList As Collection, dv As Variant, k As Long, h As Long, rr As Long
    Dim rep As Worksheet, outR As Long, missing As Object, problem As String, v As Variant
    Dim colVals() As Variant, r1 As Long, r2 As Long

    ReadSettings PrepareSettingSheet(), names, precs, blocks, weights, nPt
    GetItemMap cols, srcs, hdrs, nItem
    Set missing = CreateObject("Scripting.Dictionary")
    nIssue = 0: nUnwritten = 0

    ' 確認する列：最初の時刻列 ＋ CHECK_ITEMS
    ReDim checkX(1 To nItem)
    For x = 1 To nItem
        If srcs(x) <= -101 And timeX = 0 Then timeX = x
    Next x
    For Each item In Split(CHECK_ITEMS, ",")
        For x = 1 To nItem
            If InStr(hdrs(x), Trim$(item)) = 1 Then nChk = nChk + 1: checkX(nChk) = x
        Next x
    Next item

    Set rep = GetOrAddSheetIn(ThisWorkbook, CHECK_SHEET)
    rep.Cells.Clear
    rep.Range("A1:F1").Value = Array("確認日時", "ファイル", "地点", "日付", "時", "内容")
    rep.Range("A1:F1").Font.Bold = True
    outR = 2

    Application.ScreenUpdating = False
    Set files = TargetFiles(d1, d2)
    For Each key In files.Keys
        m1 = DateSerial(CLng(Left$(key, 4)), CLng(Right$(key, 2)), 1)
        Application.StatusBar = "確認中: " & MonthFilePath(m1)
        Set dayList = files(key)
        wasOpen = IsOpenBook(MonthFilePath(m1))
        Set wb = OpenMonthFile(m1, False, True)

        For k = 1 To nPt
            Set ws = Nothing
            If Not wb Is Nothing Then
                On Error Resume Next
                Set ws = wb.Worksheets(names(k))
                On Error GoTo 0
            End If
            For Each dv In dayList
                problem = ""
                If wb Is Nothing Then
                    problem = "ファイルがありません"
                ElseIf ws Is Nothing Then
                    problem = "シート「" & names(k) & "」がありません"
                End If
                If problem <> "" Then
                    WriteIssue rep, outR, MonthFilePath(m1), names(k), dv, "全時刻", problem
                    nIssue = nIssue + 24: nUnwritten = nUnwritten + 24
                    AddMissing missing, names(k), dv
                Else
                    r1 = RowOf(m1, dv, 1): r2 = r1 + 23
                    For h = 1 To 24
                        rr = r1 + h - 1
                        problem = ""
                        If timeX > 0 Then
                            If CStr(ws.Range(cols(timeX) & rr).Value) = "" Then problem = "未書き込み（時刻が空欄）"
                        End If
                        If problem = "" Then
                            Dim nEmpty As Long, lst As String
                            nEmpty = 0: lst = ""
                            For x = 1 To nChk
                                v = ws.Range(cols(checkX(x)) & rr).Value
                                If Not IsNumeric(v) Or CStr(v) = "" Then
                                    nEmpty = nEmpty + 1
                                    lst = lst & IIf(lst = "", "", "・") & hdrs(checkX(x))
                                End If
                            Next x
                            If nChk > 0 And nEmpty = nChk Then
                                problem = "未書き込み（確認項目がすべて空欄）"
                            ElseIf nEmpty > 0 Then
                                problem = "欠測（補完できず）：" & lst
                            End If
                        End If
                        If problem <> "" Then
                            WriteIssue rep, outR, MonthFilePath(m1), names(k), dv, h, problem
                            nIssue = nIssue + 1
                            If Left$(problem, 4) = "未書き込" Then
                                nUnwritten = nUnwritten + 1
                                AddMissing missing, names(k), dv
                            End If
                        End If
                    Next h
                End If
            Next dv
        Next k
        If Not wb Is Nothing And Not wasOpen Then wb.Close SaveChanges:=False
    Next key

    rep.Columns("A:F").AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Set CheckCore = missing
End Function

Private Sub WriteIssue(rep As Worksheet, outR As Long, ByVal path As String, ByVal nm As String, _
                       ByVal d As Date, ByVal h As Variant, ByVal problem As String)
    rep.Cells(outR, 1).Value = Now
    rep.Cells(outR, 1).NumberFormat = "yyyy/mm/dd hh:mm"
    rep.Cells(outR, 2).Value = Mid$(path, InStrRev(path, "\") + 1)
    rep.Cells(outR, 3).Value = nm
    rep.Cells(outR, 4).Value = d
    rep.Cells(outR, 4).NumberFormat = "yyyy/mm/dd"
    rep.Cells(outR, 5).Value = h
    rep.Cells(outR, 6).Value = problem
    outR = outR + 1
End Sub

Private Sub AddMissing(missing As Object, ByVal nm As String, ByVal d As Date)
    Dim key As String
    key = nm & "|" & Format(d, "yyyymmdd")
    If Not missing.Exists(key) Then missing.Add key, d
End Sub

Private Function StationsOf(missing As Object) As Collection
    Dim col As New Collection, key As Variant, nm As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For Each key In missing.Keys
        nm = Split(key, "|")(0)
        If Not seen.Exists(nm) Then seen.Add nm, 1: col.Add nm
    Next key
    Set StationsOf = col
End Function

' 地点の未書き込み日を、連続した期間（開始日・終了日）にまとめる
Private Function DayRanges(missing As Object, ByVal nm As String) As Collection
    Dim col As New Collection, key As Variant, ds() As Date, n As Long, i As Long, j As Long, t As Date
    Dim s As Date, e As Date
    ReDim ds(1 To missing.Count)
    For Each key In missing.Keys
        If Split(key, "|")(0) = nm Then n = n + 1: ds(n) = missing(key)
    Next key
    For i = 1 To n - 1                               ' 並べ替え
        For j = i + 1 To n
            If ds(j) < ds(i) Then t = ds(i): ds(i) = ds(j): ds(j) = t
        Next j
    Next i
    If n = 0 Then Set DayRanges = col: Exit Function
    s = ds(1): e = ds(1)
    For i = 2 To n
        If ds(i) = e + 1 Then
            e = ds(i)
        ElseIf ds(i) > e Then
            col.Add Array(s, e): s = ds(i): e = ds(i)
        End If
    Next i
    col.Add Array(s, e)
    Set DayRanges = col
End Function

Private Function IsOpenBook(ByVal path As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then IsOpenBook = True: Exit Function
    Next wb
End Function

'==========================================================================
' 期間・設定の読み込み
'==========================================================================
Private Function ReadPeriod(d1 As Date, d2 As Date) As Boolean
    Dim ws As Worksheet, v1 As Variant, v2 As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PERIOD_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PERIOD_SHEET
        ws.Range(START_CELL).Offset(0, -1).Value = "開始日"
        ws.Range(END_CELL).Offset(0, -1).Value = "終了日"
        ws.Range(START_CELL).Value = Date - 2
        ws.Range(END_CELL).Value = Date - 1
        ws.Range(START_CELL & "," & END_CELL).NumberFormat = "yyyy/mm/dd"
        MsgBox "シート「" & PERIOD_SHEET & "」を作成しました。" & START_CELL & "（開始日）と " & END_CELL & _
               "（終了日）を入力してから、もう一度実行してください。", vbInformation
        Exit Function
    End If
    v1 = ws.Range(START_CELL).Value
    v2 = ws.Range(END_CELL).Value
    If Not IsDate(v1) Or Not IsDate(v2) Then
        MsgBox "シート「" & PERIOD_SHEET & "」の " & START_CELL & "・" & END_CELL & " に日付を入力してください。", vbExclamation
        Exit Function
    End If
    d1 = Int(CDate(v1)): d2 = Int(CDate(v2))
    If d1 > d2 Then
        MsgBox "開始日が終了日より後になっています。", vbExclamation
        Exit Function
    End If
    If d2 >= Date Then
        If MsgBox("終了日が今日以降です。今日の分は途中までしかありません。続けますか？", vbOKCancel + vbExclamation) = vbCancel Then Exit Function
    End If
    ReadPeriod = True
End Function

Private Function ValidateSettings() As Boolean
    Dim cols As Variant, srcs As Variant, hdrs As Variant, nItem As Long, x As Long, y As Long, hasTime As Boolean
    GetItemMap cols, srcs, hdrs, nItem
    For x = 1 To nItem
        If srcs(x) <= -101 Then hasTime = True
        For y = x + 1 To nItem
            If cols(x) = cols(y) Then
                MsgBox "列の割り当てが重複しています：" & hdrs(x) & " と " & hdrs(y) & "（" & cols(x) & "列）", vbExclamation
                Exit Function
            End If
        Next y
    Next x
    If Not hasTime Then
        MsgBox "TIME_MODE=" & TIME_MODE & " に対応する時刻の列（COL_…）が空欄です。設定を確認してください。", vbExclamation
        Exit Function
    End If
    If DATA_START_ROW <= HEADER_ROW Then
        MsgBox "DATA_START_ROW は HEADER_ROW より後の行にしてください。", vbExclamation
        Exit Function
    End If
    ValidateSettings = True
End Function

Private Function StationCount() As Long
    Dim names() As String, precs() As String, blocks() As String, weights() As Double, nPt As Long
    ReadSettings PrepareSettingSheet(), names, precs, blocks, weights, nPt
    StationCount = nPt
End Function

'==========================================================================
' 1日分の取得（キャッシュつき）：24行（行番号＝時）
'==========================================================================
Private Function GetHourlyDay(ByVal precNo As String, ByVal blockNo As String, ByVal d As Date) As Variant
    Dim key As String, html As String, dat As Variant, res() As Variant, i As Long, c As Long, h As Long
    Dim msg As String
    key = "H|" & precNo & "|" & blockNo & "|" & Format(d, "yyyymmdd")
    If gCache.Exists(key) Then
        If IsArray(gCache(key)) Then GetHourlyDay = gCache(key): Exit Function
        Err.Raise vbObjectError + 20, , gCache(key)
    End If

    On Error GoTo Fail
    html = FetchPage("hourly_s1", precNo, blockNo, d)
    dat = ParseJmaDay(html)
    ReDim res(1 To 24, 1 To J_COLS)
    For h = 1 To 24: res(h, J_HOUR) = h: Next h
    For i = 1 To UBound(dat, 1)
        h = dat(i, J_HOUR)
        If h >= 1 And h <= 24 Then
            For c = 1 To J_COLS: res(h, c) = dat(i, c): Next c
        End If
    Next i
    gCache.Add key, res
    GetHourlyDay = res
    Exit Function
Fail:
    msg = Err.Description
    gCache(key) = msg
    Err.Raise vbObjectError + 20, , msg
End Function

Private Function SlotHasData(ser As Variant, ByVal i As Long) As Boolean
    Dim c As Long
    For c = 2 To J_COLS
        If Not IsEmpty(ser(i, c)) Then
            If CStr(ser(i, c)) <> "" Then SlotHasData = True: Exit Function
        End If
    Next c
End Function

Private Function DayHasData(ser As Variant, ByVal off As Long) As Boolean
    Dim h As Long
    For h = 1 To 24
        If SlotHasData(ser, off + h) Then DayHasData = True: Exit Function
    Next h
End Function

Private Function DayHasGap(ser As Variant, ByVal off As Long) As Boolean
    Dim h As Long, i As Long
    For h = 1 To 24
        i = off + h
        If VarType(ser(i, J_TEMP)) <> vbDouble Or VarType(ser(i, J_WSPD)) <> vbDouble _
           Or VarType(ser(i, J_PREC)) <> vbDouble Or VarType(ser(i, J_HUM)) <> vbDouble Then DayHasGap = True: Exit Function
    Next h
End Function

'==========================================================================
' 欠損の補完
'==========================================================================
' 欠損セルを代替地点 sFrom～sTo の同じ時刻の値で埋める
Private Sub FillHourlyAt(ser() As Variant, flg() As String, ByVal k As Long, ByVal d As Date, _
                         ByVal sFrom As Long, ByVal sTo As Long, ByVal off As Long)
    Dim subCols As Variant, h As Long, i As Long, x As Long, c As Long, s As Long
    Dim sd As Variant, need As Boolean

    subCols = Array(J_PLOC, J_PSEA, J_PREC, J_TEMP, J_DEW, J_VAP, J_HUM, J_WSPD, J_SUN, J_RAD, J_SNOWF, J_SNOWD)
    For h = 1 To 24
        i = off + h
        For x = 0 To UBound(subCols)
            c = subCols(x)
            If c = J_WSPD Then
                need = IsMissJ(ser(i, J_WSPD), J_WSPD) Or WdirMissing(ser(i, J_WDIR))
            Else
                need = IsMissJ(ser(i, c), c)
            End If
            If Not SnowGate(c = J_SNOWF Or c = J_SNOWD, i) Then need = False
            If need Then
                For s = sFrom To sTo
                    If gSubBlock(k, s) <> "" Then
                        sd = SubDay(k, s, d)
                        If IsArray(sd) Then
                            If c = J_WSPD Then
                                If VarType(sd(h, J_WSPD)) = vbDouble And Not WdirMissing(sd(h, J_WDIR)) Then
                                    ser(i, J_WSPD) = sd(h, J_WSPD): flg(i, J_WSPD) = gSubName(k, s)
                                    ser(i, J_WDIR) = sd(h, J_WDIR): flg(i, J_WDIR) = gSubName(k, s)
                                    Exit For
                                End If
                            ElseIf VarType(sd(h, c)) = vbDouble Then
                                ser(i, c) = sd(h, c): flg(i, c) = gSubName(k, s)
                                Exit For
                            End If
                        End If
                    End If
                Next s
            End If
        Next x
    Next h
End Sub

Private Function SubDay(ByVal k As Long, ByVal s As Long, ByVal d As Date) As Variant
    On Error Resume Next
    SubDay = GetHourlyDay(gSubPrec(k, s), gSubBlock(k, s), d)
    If Err.Number <> 0 Then SubDay = Empty
End Function

Private Function SnowGate(ByVal isSnow As Boolean, ByVal i As Long) As Boolean
    If gSnowPass Then
        If isSnow Then SnowGate = gSnowOK(i)
    Else
        SnowGate = Not isSnow
    End If
End Function

' 寒候年（8/1～7/31）ごとに、本来の地点で降雪・積雪の値が出た最初～最後の時刻を運用期間とする
Private Sub BuildSnowWindow(ser() As Variant, ByVal nSlot As Long, ByVal d1 As Date)
    Dim i As Long, sea As Long, s0 As Long, s1 As Long
    Dim firstV() As Long, lastV() As Long
    ReDim gSnowOK(1 To nSlot)
    s0 = SeasonOf(d1)
    s1 = SeasonOf(d1 + Int((nSlot - 1) / 24))
    ReDim firstV(s0 To s1): ReDim lastV(s0 To s1)
    For i = 1 To nSlot
        If VarType(ser(i, J_SNOWF)) = vbDouble Or VarType(ser(i, J_SNOWD)) = vbDouble Then
            sea = SeasonOf(d1 + Int((i - 1) / 24))
            If firstV(sea) = 0 Then firstV(sea) = i
            lastV(sea) = i
        End If
    Next i
    For i = 1 To nSlot
        sea = SeasonOf(d1 + Int((i - 1) / 24))
        If firstV(sea) > 0 Then gSnowOK(i) = (i > firstV(sea) And i < lastV(sea))
    Next i
End Sub

Private Sub InterpSeries(ser() As Variant, flg() As String, ByVal nSlot As Long, scalarCols As Variant, _
                         ByVal spdCol As Long, ByVal dirCol As Long, ByVal nwCol As Long, ByVal maxGap As Long)
    Dim x As Long
    If maxGap < 1 Then Exit Sub
    For x = 0 To UBound(scalarCols)
        InterpScalar ser, flg, nSlot, CLng(scalarCols(x)), maxGap
    Next x
    InterpWind ser, flg, nSlot, spdCol, dirCol, nwCol, maxGap
End Sub

' 備考：「補完：気温←呉、風速←補間」
Private Function NoteOf(flg As Variant, ByVal r As Long) As String
    Dim nm As Variant, c As Long, s As String
    nm = Split(J_NAMES, ",")
    For c = 2 To J_COLS
        If flg(r, c) <> "" Then
            If s <> "" Then s = s & "、"
            s = s & nm(c - 1) & "←" & flg(r, c)
        End If
    Next c
    If s <> "" Then s = "補完：" & s
    NoteOf = s
End Function

'==========================================================================
' 地点設定シート（代替地点1～3）
'==========================================================================
Private Function PrepareSettingSheet() As Worksheet
    Dim ws As Worksheet, r As Long, lastR As Long, dv As Variant, isNew As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SET_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SET_SHEET
        ws.Range("A2:D2").Value = Array("広島", "67", "47765", 1)
        ws.Range("A3:D3").Value = Array("岡山", "66", "47768", 1)
        ws.Range("A4:D4").Value = Array("山口", "81", "47784", 1)
        ws.Range("A5:D5").Value = Array("松江", "68", "47741", 1)
        ws.Range("A6:D6").Value = Array("鳥取", "69", "47746", 1)
    End If

    If ws.Range("E1").Value <> "代替1 地点名" Then
        isNew = True
        ws.Range("E:N").ClearContents
    ElseIf ws.Range("K1").Value <> "代替3 地点名" Then
        isNew = False
        ws.Range("K:N").ClearContents
    Else
        Set PrepareSettingSheet = ws
        Exit Function
    End If

    ws.Range("A1:M1").Value = Array("地点名（＝書き出し先のシート名）", "prec_no（都府県）", "block_no（地点）", "取得する（1＝取得、0＝しない）", _
        "代替1 地点名", "代替1 prec_no", "代替1 block_no", "代替2 地点名", "代替2 prec_no", "代替2 block_no", _
        "代替3 地点名", "代替3 prec_no", "代替3 block_no")
    ws.Range("A1:M1").Font.Bold = True
    ws.Range("E1:M1").Interior.Color = SUB_COLOR
    ws.Range("B:C,F:G,I:J,L:M").NumberFormat = "@"
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        dv = DefaultSubs(Trim$(ws.Cells(r, 1).Value))
        If IsArray(dv) Then
            If isNew Then
                ws.Cells(r, 5).Resize(1, 9).Value = dv
            Else
                ws.Cells(r, 11).Resize(1, 3).Value = Array(dv(6), dv(7), dv(8))
            End If
        End If
    Next r
    ws.Range("O1").Value = "※気象台・特別地域気象観測所（hourly_s1）の地点のみ対応。番号は気象庁ページのURLの prec_no / block_no。"
    ws.Range("O2").Value = "※D列が0または空欄の地点は取得しません。"
    ws.Range("O3").Value = "※欠損の補完順：代替1 → 代替2 → 前後の時刻から補間（" & INTERP_MAX_HOURS & "時間以内） → 代替3。不要な代替地点は空欄に。"
    ws.Columns("A:M").AutoFit
    Set PrepareSettingSheet = ws
End Function

Private Sub ReadSettings(ws As Worksheet, names() As String, precs() As String, blocks() As String, _
                         weights() As Double, nPt As Long)
    Dim r As Long, lastR As Long, w As Variant, s As Long, col As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim names(1 To 50): ReDim precs(1 To 50): ReDim blocks(1 To 50): ReDim weights(1 To 50)
    ReDim gSubName(1 To 50, 1 To N_SUB): ReDim gSubPrec(1 To 50, 1 To N_SUB): ReDim gSubBlock(1 To 50, 1 To N_SUB)
    nPt = 0
    For r = 2 To lastR
        w = ws.Cells(r, 4).Value
        If Trim$(ws.Cells(r, 1).Value) <> "" And Trim$(ws.Cells(r, 2).Value) <> "" _
           And Trim$(ws.Cells(r, 3).Value) <> "" And IsNumeric(w) And Not IsEmpty(w) Then
            If CDbl(w) > 0 Then
                nPt = nPt + 1
                names(nPt) = Trim$(ws.Cells(r, 1).Value)
                precs(nPt) = Trim$(CStr(ws.Cells(r, 2).Value))
                blocks(nPt) = Trim$(CStr(ws.Cells(r, 3).Value))
                weights(nPt) = CDbl(w)
                For s = 1 To N_SUB
                    col = 5 + (s - 1) * 3
                    If Trim$(CStr(ws.Cells(r, col + 1).Value)) <> "" And Trim$(CStr(ws.Cells(r, col + 2).Value)) <> "" Then
                        gSubName(nPt, s) = Trim$(CStr(ws.Cells(r, col).Value))
                        If gSubName(nPt, s) = "" Then gSubName(nPt, s) = Trim$(CStr(ws.Cells(r, col + 2).Value))
                        gSubPrec(nPt, s) = Trim$(CStr(ws.Cells(r, col + 1).Value))
                        gSubBlock(nPt, s) = Trim$(CStr(ws.Cells(r, col + 2).Value))
                    End If
                Next s
            End If
        End If
    Next r
End Sub

'==========================================================================
' 解析・通信などの部品
'==========================================================================
Private Function BuildUrl(ByVal precNo As String, ByVal blockNo As String, ByVal d As Date, _
                          Optional ByVal page As String = "hourly_s1") As String
    BuildUrl = "https://www.data.jma.go.jp/stats/etrn/view/" & page & ".php" & _
               "?prec_no=" & precNo & "&block_no=" & blockNo & _
               "&year=" & Year(d) & "&month=" & Month(d) & "&day=" & Day(d) & "&view="
End Function

Private Function ParseJmaDay(ByVal html As String) As Variant
    Dim p As Long, q As Long, tbl As String, e As Long, rowHtml As String
    Dim inners As Collection, res() As Variant, n As Long, c As Long, s As String
    Dim rowsFound As New Collection, r As Variant

    ' 表のIDに頼らず、ページ全体から「先頭が時刻(1～24)で17列以上」の行を集める
    tbl = html
    p = InStr(1, tbl, "<tr", vbTextCompare)
    Do While p > 0
        e = InStr(p, tbl, "</tr>", vbTextCompare)
        If e = 0 Then e = Len(tbl) + 1
        rowHtml = Mid$(tbl, p, e - p)
        Set inners = GetTdInners(rowHtml)
        If inners.Count >= J_COLS Then
            s = CleanText(inners(1))
            If IsNumeric(s) And s <> "" Then
                If CLng(s) >= 1 And CLng(s) <= 24 Then rowsFound.Add inners
            End If
        End If
        If e > Len(tbl) Then Exit Do
        p = InStr(e, tbl, "<tr", vbTextCompare)
    Loop
    If rowsFound.Count = 0 Then
        SaveDebugHtml html
        Err.Raise vbObjectError + 2, , "データの行が見つかりません。受信したHTMLを " & DebugPath() & " に保存しました"
    End If

    ReDim res(1 To rowsFound.Count, 1 To J_COLS)
    n = 0
    For Each r In rowsFound
        n = n + 1
        For c = 1 To J_COLS
            s = CleanText(r(c))
            If s = "" Then s = GetImgAlt(r(c))          ' 天気はアイコン画像
            If c = J_HOUR Then
                res(n, c) = CLng(s)
            ElseIf c = J_WDIR Or c = J_WX Then
                res(n, c) = s
            Else
                res(n, c) = ToJmaNum(s)
            End If
        Next c
    Next r
    ParseJmaDay = res
End Function

Private Function ToJmaNum(ByVal s As String) As Variant
    Dim t As String
    t = s
    If t = "--" Then ToJmaNum = 0#: Exit Function
    Do While Len(t) > 0 And (Right$(t, 1) = ")" Or Right$(t, 1) = "]")
        t = Left$(t, Len(t) - 1)
    Loop
    If t <> "" And IsNumeric(t) Then
        ToJmaNum = CDbl(t)
    Else
        ToJmaNum = s
    End If
End Function

Private Function NorthComponent(ByVal dirS As String, ByVal spd As Variant) As Variant
    Dim idx As Long, dirList As Variant
    NorthComponent = ""
    If VarType(spd) <> vbDouble Then Exit Function
    If dirS = "静穏" Then NorthComponent = 0#: Exit Function
    dirS = Replace(Replace(dirS, ")", ""), "]", "")
    dirList = Split(DIR16, ",")
    For idx = 0 To 15
        If dirList(idx) = dirS Then
            NorthComponent = Round(spd * Cos(idx * 22.5 * 4 * Atn(1) / 180), 2)
            Exit Function
        End If
    Next idx
End Function

Private Function DirIndex(ByVal dirS As String) As Long
    Dim dirList As Variant, idx As Long
    DirIndex = -1
    If dirS = "静穏" Then DirIndex = -2: Exit Function
    dirList = Split(DIR16, ",")
    For idx = 0 To 15
        If dirList(idx) = dirS Then DirIndex = idx: Exit Function
    Next idx
End Function

Private Function VecToDir(ByVal nComp As Double, ByVal eComp As Double) As String
    Dim deg As Double, idx As Long
    If Sqr(nComp * nComp + eComp * eComp) < 0.05 Then VecToDir = "静穏": Exit Function
    deg = Atan2Deg(eComp, nComp)                 ' 北=0°、時計回り
    idx = Int(deg / 22.5 + 0.5) Mod 16
    VecToDir = Split(DIR16, ",")(idx)
End Function

Private Function Atan2Deg(ByVal y As Double, ByVal x As Double) As Double
    Dim r As Double
    If x > 0 Then
        r = Atn(y / x)
    ElseIf x < 0 Then
        r = Atn(y / x) + PiVal()
    ElseIf y > 0 Then
        r = PiVal() / 2
    ElseIf y < 0 Then
        r = -PiVal() / 2
    Else
        r = 0
    End If
    r = r * 180 / PiVal()
    If r < 0 Then r = r + 360
    If r >= 360 Then r = r - 360
    Atan2Deg = r
End Function

Private Function PiVal() As Double
    PiVal = 4 * Atn(1)
End Function

Private Function DebugPath() As String
    Dim folder As String
    folder = ThisWorkbook.Path
    If folder = "" Or Left$(folder, 4) = "http" Then folder = Environ$("TEMP")
    DebugPath = folder & "\jma_debug.html"
End Function

Private Sub SaveDebugHtml(ByVal html As String)
    Static saved As Boolean
    Dim stm As Object
    If saved Then Exit Sub
    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.WriteText html
    stm.SaveToFile DebugPath(), 2
    stm.Close
    saved = True
End Sub

Private Function GetTdInners(ByVal rowHtml As String) As Collection
    Dim col As New Collection, c As Long, gt As Long, q As Long
    c = FindTdStart(rowHtml, 1)
    Do While c > 0
        gt = InStr(c, rowHtml, ">")
        q = InStr(gt, rowHtml, "</td>", vbTextCompare)
        If gt = 0 Or q = 0 Then Exit Do
        col.Add Mid$(rowHtml, gt + 1, q - gt - 1)
        c = FindTdStart(rowHtml, q)
    Loop
    Set GetTdInners = col
End Function

Private Function FindTdStart(ByVal s As String, ByVal start As Long) As Long
    Dim a As Long, b As Long
    a = InStr(start, s, "<td>", vbTextCompare)
    b = InStr(start, s, "<td ", vbTextCompare)
    If a = 0 Then
        FindTdStart = b
    ElseIf b = 0 Then
        FindTdStart = a
    Else
        FindTdStart = IIf(a < b, a, b)
    End If
End Function

Private Function GetImgAlt(ByVal s As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, "alt=""", vbTextCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 5, s, """")
    If q > p Then GetImgAlt = Mid$(s, p + 5, q - p - 5)
End Function

Private Function CleanText(ByVal s As String) As String
    Dim p As Long, q As Long
    p = InStr(s, "<")
    Do While p > 0
        q = InStr(p, s, ">")
        If q = 0 Then
            s = Left$(s, p - 1)
            Exit Do
        End If
        s = Left$(s, p - 1) & Mid$(s, q + 1)
        p = InStr(s, "<")
    Loop
    s = Replace(s, "&nbsp;", "")
    s = Replace(s, "&#160;", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    CleanText = s
End Function

Private Function HttpGetUtf8(ByVal url As String) As String
    Dim body As Variant, stm As Object
    body = TryGet("WinHttp.WinHttpRequest.5.1", url)
    If IsEmpty(body) Then body = TryGet("MSXML2.ServerXMLHTTP.6.0", url)
    If IsEmpty(body) Then body = TryGet("MSXML2.XMLHTTP", url)
    If IsEmpty(body) Then Err.Raise vbObjectError + 1, , "気象庁サイトに接続できませんでした"

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write body
    stm.Position = 0
    stm.Type = 2: stm.Charset = "utf-8"
    HttpGetUtf8 = stm.ReadText
    stm.Close
End Function

Private Function TryGet(ByVal progId As String, ByVal url As String) As Variant
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject(progId)
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status = 200 Then TryGet = http.responseBody
    Exit Function
Fail:
    TryGet = Empty
End Function

Private Function IsMissJ(ByVal v As Variant, ByVal c As Long) As Boolean
    If VarType(v) = vbDouble Then Exit Function
    Select Case c
        Case J_SUN, J_RAD, J_SNOWF, J_SNOWD
            If IsEmpty(v) Then
                IsMissJ = True
            Else
                IsMissJ = (CStr(v) <> "")
            End If
        Case Else
            IsMissJ = True
    End Select
End Function

Private Function WdirMissing(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then WdirMissing = True: Exit Function
    WdirMissing = (DirIndex(CleanDir(v)) = -1)
End Function

Private Function CleanDir(ByVal v As Variant) As String
    CleanDir = Replace(Replace(CStr(v), ")", ""), "]", "")
End Function

Private Sub InterpScalar(ser() As Variant, flg() As String, ByVal nSlot As Long, ByVal c As Long, ByVal maxGap As Long)
    Dim i As Long, j As Long, p As Long, gap As Long
    p = 0
    For i = 1 To nSlot
        If VarType(ser(i, c)) = vbDouble Then
            gap = i - p - 1
            If p > 0 And gap >= 1 And gap <= maxGap Then
                For j = p + 1 To i - 1
                    ser(j, c) = Round(ser(p, c) + (ser(i, c) - ser(p, c)) * (j - p) / (i - p), 2)
                    flg(j, c) = INTERP_MARK
                Next j
            End If
            p = i
        End If
    Next i
End Sub

Private Sub InterpWind(ser() As Variant, flg() As String, ByVal nSlot As Long, ByVal spdCol As Long, _
                       ByVal dirCol As Long, ByVal nwCol As Long, ByVal maxGap As Long)
    Dim i As Long, j As Long, p As Long, gap As Long, t As Double
    Dim nP As Double, eP As Double, nQ As Double, eQ As Double, nJ As Double, eJ As Double
    Dim spd As Double, dirS As String
    p = 0
    For i = 1 To nSlot
        If WindValid(ser, i, spdCol, dirCol) Then
            gap = i - p - 1
            If p > 0 And gap >= 1 And gap <= maxGap Then
                WindVec ser, p, spdCol, dirCol, nwCol, nP, eP
                WindVec ser, i, spdCol, dirCol, nwCol, nQ, eQ
                For j = p + 1 To i - 1
                    t = (j - p) / (i - p)
                    spd = Round(ser(p, spdCol) + (ser(i, spdCol) - ser(p, spdCol)) * t, 2)
                    nJ = nP + (nQ - nP) * t
                    eJ = eP + (eQ - eP) * t
                    dirS = VecToDir(nJ, eJ)
                    If dirS = "静穏" And spd >= 0.3 Then
                        If t < 0.5 Then dirS = CleanDir(ser(p, dirCol)) Else dirS = CleanDir(ser(i, dirCol))
                    End If
                    ser(j, spdCol) = spd: flg(j, spdCol) = INTERP_MARK
                    ser(j, dirCol) = dirS: flg(j, dirCol) = INTERP_MARK
                    If nwCol > 0 Then ser(j, nwCol) = Round(nJ, 2): flg(j, nwCol) = INTERP_MARK
                Next j
            End If
            p = i
        End If
    Next i
End Sub

Private Function WindValid(ser() As Variant, ByVal i As Long, ByVal spdCol As Long, ByVal dirCol As Long) As Boolean
    If VarType(ser(i, spdCol)) <> vbDouble Then Exit Function
    WindValid = Not WdirMissing(ser(i, dirCol))
End Function

Private Sub WindVec(ser() As Variant, ByVal i As Long, ByVal spdCol As Long, ByVal dirCol As Long, _
                    ByVal nwCol As Long, nComp As Double, eComp As Double)
    Dim idx As Long, ang As Double
    idx = DirIndex(CleanDir(ser(i, dirCol)))
    If idx < 0 Then                                  ' 静穏
        nComp = 0: eComp = 0
    Else
        ang = idx * 22.5 * PiVal() / 180
        nComp = ser(i, spdCol) * Cos(ang)
        eComp = ser(i, spdCol) * Sin(ang)
        If nwCol > 0 Then
            If VarType(ser(i, nwCol)) = vbDouble Then nComp = ser(i, nwCol)
        End If
    End If
End Sub

Private Function SeasonOf(ByVal dd As Date) As Long
    If Month(dd) >= 8 Then SeasonOf = Year(dd) Else SeasonOf = Year(dd) - 1
End Function

Private Function AnyTrue(arr() As Boolean, ByVal i1 As Long, ByVal i2 As Long) As Boolean
    Dim i As Long
    For i = i1 To i2
        If arr(i) Then AnyTrue = True: Exit Function
    Next i
End Function

Private Function DefaultSubs(ByVal nm As String) As Variant
    Select Case nm
        Case "広島": DefaultSubs = Array("呉", "67", "47766", "福山", "67", "47767", "松山", "73", "47887")
        Case "岡山": DefaultSubs = Array("津山", "66", "47756", "福山", "67", "47767", "姫路", "63", "47769")
        Case "山口": DefaultSubs = Array("萩", "81", "47754", "下関", "81", "47762", "浜田", "68", "47755")
        Case "松江": DefaultSubs = Array("米子", "69", "47744", "鳥取", "69", "47746", "境", "69", "47742")
        Case "鳥取": DefaultSubs = Array("米子", "69", "47744", "豊岡", "63", "47747", "姫路", "63", "47769")
        Case Else:   DefaultSubs = Empty
    End Select
End Function

Private Function FetchPage(ByVal page As String, ByVal precNo As String, ByVal blockNo As String, ByVal d As Date) As String
    FetchPage = HttpGetUtf8(BuildUrl(precNo, blockNo, d, page))
    Application.Wait Now + TimeSerial(0, 0, WAIT_SEC)
End Function
