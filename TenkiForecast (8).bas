Attribute VB_Name = "TenkiForecast"
'==========================================================================
' tenki.jp 1時間天気 複数地点取得＋加重平均
'   実行: マクロ「GetTenkiForecast」      … 1時間ごと
'         マクロ「GetTenkiForecast30min」 … 1時間値を2行に複製した疑似30分値（コマ1～48）
'   シート:
'     「地点設定」 … 地点名 / URL / 重み（初回実行時に自動作成。重みは自由に変更可）
'     「天気予報」 … 地点ごとの1時間予報（天気・気温・降水確率・降水量・降雪量・湿度・風向・風速）
'     「加重平均」 … 時刻ごとの各地点の気温と、数値項目の加重平均
'   参照設定不要。HTML解析は文字列処理のみ
'==========================================================================
Option Explicit

'---- 設定 --------------------------------------------------------------
Private Const SET_SHEET As String = "地点設定"
Private Const OUT_SHEET As String = "天気予報"
Private Const AVG_SHEET As String = "加重平均"
Private Const GET_DAYS As Long = 3              ' 取得日数　1=今日 / 2=今日・明日 / 3=今日・明日・明後日
Private Const APPEND_MODE As Boolean = False    ' True で「天気予報」シートに追記（予報履歴を残す）
Private Const WAIT_SEC As Long = 1              ' 地点ごとの待機秒
Private Const OUT30_SHEET As String = "天気予報30分"
Private Const AVG30_SHEET As String = "加重平均30分"
'------------------------------------------------------------------------

' ParseTable が返す配列の列
Private Const P_DATE = 1, P_HOUR = 2, P_DT = 3, P_WX = 4, P_TEMP = 5, P_PROB = 6
Private Const P_PREC = 7, P_SNOW = 8, P_HUM = 9, P_WDIR = 10, P_WSPD = 11, P_NWIND = 12
Private Const P_COLS = 12

' 16方位（北=0°から時計回りに22.5°刻み）
Private Const DIR16 As String = "北,北北東,北東,東北東,東,東南東,南東,南南東,南,南南西,南西,西南西,西,西北西,北西,北北西"


Private gStep As String
Private gMissing As String

Public Sub GetTenkiForecast()
    RunTenki False
End Sub

Public Sub GetTenkiForecast30min()
    RunTenki True
End Sub

' half=True：時刻h の予報を コマ 2h-1（h-1:00～h-1:30）と 2h（h-1:30～h:00）の2行に複製
Private Sub RunTenki(ByVal half As Boolean)
    Dim wsSet As Worksheet, wsOut As Worksheet
    Dim names() As String, urls() As String, weights() As Double, nPt As Long
    Dim k As Long, html As String, fetched As Date, outRow As Long
    Dim dat As Variant, tbls As Variant, t As Long, dayOffset As Long
    Dim errMsg As String

    ' 加重平均の集計用（slot × 項目）
    Dim sumV() As Double, sumW() As Double, slotDt() As Variant
    Dim cityTemp() As Variant
    Dim avgItems As Variant, a As Long, i As Long, slot As Long, v As Variant
    Dim nSlot As Long, perDay As Long, j As Long, nRep As Long, km As Long, hr As Long
    Dim rowVals As Variant, tStart As Date, tEnd As Date

    On Error GoTo ErrHandler
    fetched = Now
    gMissing = "": errMsg = ""

    gStep = "地点設定の読み込み"
    Set wsSet = PrepareSettingSheet()
    ReadSettings wsSet, names, urls, weights, nPt
    If nPt = 0 Then Err.Raise vbObjectError + 10, , "「地点設定」シートに地点がありません"

    gStep = "シート準備"
    Set wsOut = PrepareOutSheet(half)
    outRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1

    avgItems = Array(P_TEMP, P_PROB, P_PREC, P_SNOW, P_HUM, P_WSPD, P_NWIND)
    If half Then perDay = 48: nRep = 2 Else perDay = 24: nRep = 1
    nSlot = perDay * 3
    ReDim sumV(1 To nSlot, 0 To UBound(avgItems))
    ReDim sumW(1 To nSlot, 0 To UBound(avgItems))
    ReDim slotDt(1 To nSlot, 1 To 4)               ' 日付・時(コマ)・終了日時・開始日時
    ReDim cityTemp(1 To nSlot, 1 To nPt)

    Select Case GET_DAYS
        Case 1: tbls = Array("forecast-point-1h-today")
        Case 2: tbls = Array("forecast-point-1h-today", "forecast-point-1h-tomorrow")
        Case Else
            tbls = Array("forecast-point-1h-today", "forecast-point-1h-tomorrow", "forecast-point-1h-dayaftertomorrow")
    End Select

    Application.ScreenUpdating = False
    For k = 1 To nPt
        Application.StatusBar = "取得中: " & names(k)
        On Error Resume Next
        gStep = names(k) & " 通信"
        html = HttpGetUtf8(urls(k))
        If Err.Number <> 0 Then
            errMsg = errMsg & "・" & names(k) & "：" & Err.Description & vbCrLf
            Err.Clear
            On Error GoTo ErrHandler
            GoTo NextCity
        End If
        On Error GoTo ErrHandler

        For t = 0 To UBound(tbls)
            dayOffset = t
            gStep = names(k) & " 表の解析"
            On Error Resume Next
            dat = ParseTable(html, CStr(tbls(t)), Date + dayOffset, names(k))
            If Err.Number <> 0 Then
                errMsg = errMsg & "・" & names(k) & "：" & Err.Description & vbCrLf
                Err.Clear
                On Error GoTo ErrHandler
                GoTo NextTable
            End If
            On Error GoTo ErrHandler

            For i = 1 To UBound(dat, 1)
                hr = dat(i, P_HOUR)
                For j = 1 To nRep
                    If half Then
                        km = 2 * hr - 2 + j
                        tStart = dat(i, P_DATE) + TimeSerial(0, (km - 1) * 30, 0)
                        tEnd = dat(i, P_DATE) + TimeSerial(0, km * 30, 0)
                    Else
                        km = hr
                        tStart = dat(i, P_DT) - TimeSerial(1, 0, 0)
                        tEnd = dat(i, P_DT)
                    End If

                    ' 「天気予報」シートへ
                    wsOut.Cells(outRow, 1).Value = fetched
                    wsOut.Cells(outRow, 2).Value = names(k)
                    If half Then
                        rowVals = RowOf30(dat, i, km, tStart, tEnd)
                        wsOut.Cells(outRow, 3).Resize(1, UBound(rowVals, 2)).Value = rowVals
                    Else
                        wsOut.Cells(outRow, 3).Resize(1, P_COLS).Value = RowOf(dat, i)
                    End If
                    outRow = outRow + 1

                    ' 加重平均の集計
                    slot = dayOffset * perDay + km
                    If slot >= 1 And slot <= nSlot Then
                        slotDt(slot, 1) = dat(i, P_DATE)
                        slotDt(slot, 2) = km
                        slotDt(slot, 3) = tEnd
                        slotDt(slot, 4) = tStart
                        cityTemp(slot, k) = dat(i, P_TEMP)
                        For a = 0 To UBound(avgItems)
                            v = dat(i, avgItems(a))
                            If VarType(v) = vbDouble Then
                                sumV(slot, a) = sumV(slot, a) + v * weights(k)
                                sumW(slot, a) = sumW(slot, a) + weights(k)
                            End If
                        Next a
                    End If
                Next j
            Next i
NextTable:
        Next t
NextCity:
        If k < nPt Then Application.Wait Now + TimeSerial(0, 0, WAIT_SEC)
    Next k

    gStep = "加重平均の出力"
    WriteAverage names, weights, nPt, avgItems, sumV, sumW, slotDt, cityTemp, nSlot, half

    wsOut.Columns("A:P").AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True

    If errMsg = "" And gMissing = "" Then
        MsgBox "完了しました（" & nPt & " 地点）", vbInformation
    Else
        MsgBox "完了しました。ただし次の問題がありました：" & vbCrLf & errMsg & gMissing, vbExclamation
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "処理: " & gStep & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' 30分用の1行：日付・コマ・開始・終了・天気～北風成分
Private Function RowOf30(dat As Variant, ByVal i As Long, ByVal km As Long, _
                         ByVal tStart As Date, ByVal tEnd As Date) As Variant
    Dim r() As Variant, c As Long
    ReDim r(1 To 1, 1 To 4 + (P_COLS - P_WX + 1))
    r(1, 1) = dat(i, P_DATE): r(1, 2) = km: r(1, 3) = tStart: r(1, 4) = tEnd
    For c = P_WX To P_COLS: r(1, 4 + c - P_WX + 1) = dat(i, c): Next c
    RowOf30 = r
End Function

Private Function RowOf(dat As Variant, ByVal i As Long) As Variant
    Dim r() As Variant, c As Long
    ReDim r(1 To 1, 1 To P_COLS)
    For c = 1 To P_COLS: r(1, c) = dat(i, c): Next c
    RowOf = r
End Function

'---- 加重平均シート -------------------------------------------------------
Private Sub WriteAverage(names() As String, weights() As Double, ByVal nPt As Long, avgItems As Variant, _
                         sumV() As Double, sumW() As Double, slotDt As Variant, cityTemp As Variant, _
                         ByVal nSlot As Long, ByVal half As Boolean)
    Dim ws As Worksheet, slot As Long, r As Long, k As Long, a As Long, c As Long
    Dim itemNames As Variant, totalW As Double

    itemNames = Array("気温(℃)", "降水確率(%)", "降水量(mm/h)", "降雪量(cm/h)", "湿度(%)", "風速(m/s)", "北風成分(m/s)")

    If half Then Set ws = GetOrAddSheet(AVG30_SHEET) Else Set ws = GetOrAddSheet(AVG_SHEET)
    ws.Cells.Clear

    ' 重みの一覧
    For k = 1 To nPt: totalW = totalW + weights(k): Next k
    ws.Range("A1").Value = "重み"
    For k = 1 To nPt
        ws.Cells(1, 1 + k).Value = names(k) & "：" & weights(k) & "（" & Format(weights(k) / totalW, "0.0%") & "）"
    Next k

    ' 見出し
    r = 3
    ws.Cells(r, 1).Value = "日付"
    ws.Cells(r, 2).Value = IIf(half, "コマ", "時")
    ws.Cells(r, 3).Value = "開始時刻"
    ws.Cells(r, 4).Value = "終了時刻"
    c = 5
    For k = 1 To nPt
        ws.Cells(r, c).Value = names(k) & " 気温": c = c + 1
    Next k
    For a = 0 To UBound(itemNames)
        ws.Cells(r, c + a).Value = "加重平均 " & itemNames(a)
    Next a
    ws.Rows(r).Font.Bold = True

    ' データ
    r = r + 1
    For slot = 1 To nSlot
        If Not IsEmpty(slotDt(slot, 1)) Then
            ws.Cells(r, 1).Value = slotDt(slot, 1)
            ws.Cells(r, 2).Value = slotDt(slot, 2)
            ws.Cells(r, 3).Value = slotDt(slot, 4)
            ws.Cells(r, 4).Value = slotDt(slot, 3)
            c = 5
            For k = 1 To nPt
                ws.Cells(r, c).Value = cityTemp(slot, k): c = c + 1
            Next k
            For a = 0 To UBound(avgItems)
                If sumW(slot, a) > 0 Then
                    ws.Cells(r, c + a).Value = Round(sumV(slot, a) / sumW(slot, a), 2)
                End If
            Next a
            r = r + 1
        End If
    Next slot

    ws.Columns(1).NumberFormat = "yyyy/mm/dd"
    ws.Columns(3).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Columns(4).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit
End Sub

'---- 地点設定 -------------------------------------------------------------
Private Function PrepareSettingSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SET_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SET_SHEET
        ws.Range("A1:C1").Value = Array("地点名", "URL（tenki.jp 1時間天気）", "重み")
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("A2:C2").Value = Array("広島", "https://tenki.jp/forecast/7/37/6710/34100/1hour.html", 1)
        ws.Range("A3:C3").Value = Array("岡山", "https://tenki.jp/forecast/7/36/6610/33100/1hour.html", 1)
        ws.Range("A4:C4").Value = Array("山口", "https://tenki.jp/forecast/7/38/8120/35203/1hour.html", 1)
        ws.Range("A5:C5").Value = Array("松江", "https://tenki.jp/forecast/7/35/6810/32201/1hour.html", 1)
        ws.Range("A6:C6").Value = Array("鳥取", "https://tenki.jp/forecast/7/34/6910/31201/1hour.html", 1)
        ws.Range("E1").Value = "※重みは比率で扱います（合計が1でなくても可）。重みが0または空欄の地点は取得しません。"
        ws.Columns("A:C").AutoFit
    End If
    Set PrepareSettingSheet = ws
End Function

Private Sub ReadSettings(ws As Worksheet, names() As String, urls() As String, weights() As Double, nPt As Long)
    Dim r As Long, lastR As Long, w As Variant
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim names(1 To 50): ReDim urls(1 To 50): ReDim weights(1 To 50)
    nPt = 0
    For r = 2 To lastR
        w = ws.Cells(r, 3).Value
        If Trim$(ws.Cells(r, 1).Value) <> "" And Trim$(ws.Cells(r, 2).Value) <> "" _
           And IsNumeric(w) And Not IsEmpty(w) Then
            If CDbl(w) > 0 Then
                nPt = nPt + 1
                names(nPt) = Trim$(ws.Cells(r, 1).Value)
                urls(nPt) = Trim$(ws.Cells(r, 2).Value)
                weights(nPt) = CDbl(w)
            End If
        End If
    Next r
End Sub

'---- 表の解析（1日分を配列で返す）----------------------------------------
Private Function ParseTable(ByVal html As String, ByVal tableId As String, ByVal d As Date, _
                            ByVal cityName As String) As Variant
    Dim p As Long, q As Long, tbl As String
    Dim hours As Collection, wx As Collection, temps As Collection, prob As Collection
    Dim prec As Collection, snow As Collection, hum As Collection
    Dim wDir() As String, wSpd() As Variant, sample As String
    Dim i As Long, n As Long, hr As Long, res() As Variant

    p = InStr(1, html, "id=""" & tableId & """", vbTextCompare)
    If p = 0 Then Err.Raise vbObjectError + 2, , "表 " & tableId & " が見つかりません"
    q = InStr(p, html, "</table>", vbTextCompare)
    tbl = Mid$(html, p, q - p)

    Set hours = FindRow(tbl, "hour", "時刻")
    Set wx = FindRow(tbl, "weather", "天気")
    Set temps = FindRow(tbl, "temperature", "気温")
    Set prob = FindRow(tbl, "prob-precip", "降水確率")
    Set prec = FindRow(tbl, "precipitation", "降水量")
    Set snow = FindRow(tbl, "snow_fall|snow|snowfall", "降雪")
    Set hum = FindRow(tbl, "humidity", "湿度")
    If hours Is Nothing Then Err.Raise vbObjectError + 3, , "時刻の行が見つかりません"

    ' 風向：<tr class="wind-blow">（アイコン＋文字）、風速：<tr class="wind-speed">（数値）
    Dim dirC As Collection, spdC As Collection
    Set dirC = FindRow(tbl, "wind-blow|wind-direction", "")
    Set spdC = FindRow(tbl, "wind-speed", "")
    If Not dirC Is Nothing And Not spdC Is Nothing Then
        ReDim wDir(1 To hours.Count): ReDim wSpd(1 To hours.Count)
        For i = 1 To hours.Count
            wDir(i) = FindDirection(Pick(dirC, i))
            wSpd(i) = ToNum(Pick(spdC, i))
        Next i
    Else
        ' 構成が変わった場合の予備：wind を含む行をまとめて調べる
        ReadWind tbl, hours.Count, wDir, wSpd, sample
    End If

    NoteMissing wx, cityName, "天気": NoteMissing temps, cityName, "気温"
    NoteMissing prob, cityName, "降水確率": NoteMissing prec, cityName, "降水量"
    NoteMissing snow, cityName, "降雪量": NoteMissing hum, cityName, "湿度"
    If Not AnyNumeric(wSpd) Then
        gMissing = gMissing & "・" & cityName & "：風速が見つかりません（風の行の先頭: " & sample & "）" & vbCrLf
    End If

    For i = 1 To hours.Count
        If IsNumeric(hours(i)) Then n = n + 1
    Next i
    If n = 0 Then Err.Raise vbObjectError + 4, , "時刻の値が読み取れません"
    ReDim res(1 To n, 1 To P_COLS)

    n = 0
    For i = 1 To hours.Count
        If IsNumeric(hours(i)) Then
            n = n + 1
            hr = CLng(hours(i))
            res(n, P_DATE) = d
            res(n, P_HOUR) = hr
            res(n, P_DT) = d + TimeSerial(hr, 0, 0)      ' 24時→翌0時
            res(n, P_WX) = Pick(wx, i)
            res(n, P_TEMP) = ToNum(Pick(temps, i))
            res(n, P_PROB) = ToNum(Pick(prob, i))
            res(n, P_PREC) = ToNum(Pick(prec, i))
            res(n, P_SNOW) = ToNum(Pick(snow, i))
            res(n, P_HUM) = ToNum(Pick(hum, i))
            res(n, P_WDIR) = wDir(i)
            res(n, P_WSPD) = wSpd(i)
            res(n, P_NWIND) = NorthComponent(wDir(i), wSpd(i))
        End If
    Next i
    ParseTable = res
End Function

Private Sub NoteMissing(col As Collection, ByVal cityName As String, ByVal itemName As String)
    Dim s As String
    If col Is Nothing Then
        s = "・" & cityName & "：" & itemName & " が見つかりません" & vbCrLf
        If InStr(gMissing, s) = 0 Then gMissing = gMissing & s
    End If
End Sub

'---- 行の検索 ------------------------------------------------------------
Private Function FindRow(ByVal tbl As String, ByVal classList As String, ByVal label As String) As Collection
    Dim cls As Variant, rowHtml As String, p As Long, e As Long

    For Each cls In Split(classList, "|")
        rowHtml = GetRowByClass(tbl, CStr(cls))
        If rowHtml <> "" Then
            Set FindRow = GetTdCells(rowHtml)
            If FindRow.Count > 0 Then Exit Function
            Set FindRow = Nothing
        End If
    Next cls

    If label = "" Then Exit Function
    p = InStr(1, tbl, "<tr", vbTextCompare)
    Do While p > 0
        e = InStr(p, tbl, "</tr>", vbTextCompare)
        If e = 0 Then Exit Do
        rowHtml = Mid$(tbl, p, e - p)
        If InStr(GetThText(rowHtml), label) > 0 Then
            Set FindRow = GetTdCells(rowHtml)
            If FindRow.Count > 0 Then Exit Function
            Set FindRow = Nothing
        End If
        p = InStr(e, tbl, "<tr", vbTextCompare)
    Loop
End Function

Private Function GetRowByClass(ByVal tbl As String, ByVal cls As String) As String
    Dim p As Long, e As Long
    p = InStr(1, tbl, "<tr class=""" & cls & """", vbTextCompare)
    If p = 0 Then p = InStr(1, tbl, "<tr class=""" & cls & " ", vbTextCompare)
    If p = 0 Then Exit Function
    e = InStr(p, tbl, "</tr>", vbTextCompare)
    If e = 0 Then Exit Function
    GetRowByClass = Mid$(tbl, p, e - p)
End Function

'---- HTML部品の取り出し ---------------------------------------------------
Private Function GetThText(ByVal rowHtml As String) As String
    Dim p As Long, gt As Long, q As Long
    p = InStr(1, rowHtml, "<th>", vbTextCompare)
    If p = 0 Then p = InStr(1, rowHtml, "<th ", vbTextCompare)
    If p = 0 Then Exit Function
    gt = InStr(p, rowHtml, ">")
    q = InStr(gt, rowHtml, "</th>", vbTextCompare)
    If q = 0 Then Exit Function
    GetThText = CleanText(Mid$(rowHtml, gt + 1, q - gt - 1))
End Function

Private Function GetTdCells(ByVal rowHtml As String) As Collection
    Dim col As New Collection
    Dim c As Long, gt As Long, q As Long, inner As String, s As String
    c = FindTdStart(rowHtml, 1)
    Do While c > 0
        gt = InStr(c, rowHtml, ">")
        q = InStr(gt, rowHtml, "</td>", vbTextCompare)
        If gt = 0 Or q = 0 Then Exit Do
        inner = Mid$(rowHtml, gt + 1, q - gt - 1)
        s = CleanText(inner)
        If s = "" Then s = GetImgAlt(inner)
        col.Add s
        c = FindTdStart(rowHtml, q)
    Loop
    Set GetTdCells = col
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
    s = Replace(s, "&amp;", "&")
    s = Replace(s, ChrW(160), "")
    s = Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    CleanText = s
End Function

'---- 風向・風速 -----------------------------------------------------------
Private Sub ReadWind(ByVal tbl As String, ByVal n As Long, wDir() As String, wSpd() As Variant, sample As String)
    Dim p As Long, gt As Long, e As Long, trTag As String, rowHtml As String
    Dim inners As Collection, i As Long, d As String, v As Variant

    ReDim wDir(1 To n): ReDim wSpd(1 To n)
    For i = 1 To n: wSpd(i) = "": Next i

    p = InStr(1, tbl, "<tr", vbTextCompare)
    Do While p > 0
        gt = InStr(p, tbl, ">")
        e = InStr(p, tbl, "</tr>", vbTextCompare)
        If gt = 0 Or e = 0 Then Exit Do
        trTag = Mid$(tbl, p, gt - p + 1)
        rowHtml = Mid$(tbl, p, e - p)
        If InStr(1, trTag, "wind", vbTextCompare) > 0 Or InStr(GetThText(rowHtml), "風") > 0 Then
            Set inners = GetTdInners(rowHtml)
            If inners.Count > 0 And sample = "" Then
                sample = Left$(Replace(Replace(inners(1), vbCr, ""), vbLf, ""), 120)
            End If
            For i = 1 To IIf(inners.Count < n, inners.Count, n)
                ParseWindCell inners(i), d, v
                If wDir(i) = "" And d <> "" Then wDir(i) = d
                If VarType(wSpd(i)) = vbString And Not IsEmpty(v) Then wSpd(i) = v
            Next i
        End If
        p = InStr(e, tbl, "<tr", vbTextCompare)
    Loop
End Sub

' 1マスから風向（16方位 or 静穏）と風速（数値）を取り出す
Private Sub ParseWindCell(ByVal inner As String, dirS As String, spdV As Variant)
    Dim txt As String, alt As String, src As String

    txt = CleanText(inner)
    alt = GetAllImgAlt(inner) & GetAttr(inner, "title")
    src = alt & txt

    dirS = FindDirection(src)
    spdV = Empty
    If dirS = "静穏" Then spdV = 0#
    Dim num As String
    num = FirstNumber(txt)
    If num = "" Then num = FirstNumber(src)
    If num <> "" Then spdV = CDbl(num)
End Sub

Private Function FindDirection(ByVal s As String) As String
    Dim d As Variant, best As String
    If InStr(s, "静穏") > 0 Then FindDirection = "静穏": Exit Function
    For Each d In Split(DIR16, ",")           ' 一番長く一致する方位を採用
        If InStr(s, d) > 0 And Len(d) > Len(best) Then best = d
    Next d
    FindDirection = best
End Function

Private Function FirstNumber(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Or (ch = "." And r <> "") Then
            r = r & ch
        ElseIf r <> "" Then
            Exit For
        End If
    Next i
    If Right$(r, 1) = "." Then r = Left$(r, Len(r) - 1)
    FirstNumber = r
End Function

' 北風成分：北寄りの風が＋、南寄りの風が－（風速×cos(風向角)）
Private Function NorthComponent(ByVal dirS As String, ByVal spd As Variant) As Variant
    Dim idx As Long, dirList As Variant
    NorthComponent = ""
    If VarType(spd) = vbString Or IsEmpty(spd) Then Exit Function
    If dirS = "静穏" Then NorthComponent = 0#: Exit Function
    dirList = Split(DIR16, ",")
    For idx = 0 To 15
        If dirList(idx) = dirS Then
            NorthComponent = Round(spd * Cos(idx * 22.5 * Application.WorksheetFunction.Pi() / 180), 2)
            Exit Function
        End If
    Next idx
End Function

Private Function AnyNumeric(arr() As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If VarType(arr(i)) = vbDouble Then AnyNumeric = True: Exit Function
    Next i
End Function

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

Private Function GetAllImgAlt(ByVal s As String) As String
    Dim p As Long, q As Long, r As String
    p = InStr(1, s, "alt=""", vbTextCompare)
    Do While p > 0
        q = InStr(p + 5, s, """")
        If q = 0 Then Exit Do
        r = r & Mid$(s, p + 5, q - p - 5)
        p = InStr(q, s, "alt=""", vbTextCompare)
    Loop
    GetAllImgAlt = r
End Function

Private Function GetAttr(ByVal s As String, ByVal attrName As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, attrName & "=""", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(attrName) + 2
    q = InStr(p, s, """")
    If q > p Then GetAttr = Mid$(s, p, q - p)
End Function

'---- 値の整形 -------------------------------------------------------------

Private Function Pick(col As Collection, ByVal i As Long) As String
    If col Is Nothing Then Exit Function
    If i <= col.Count Then Pick = col(i)
End Function

Private Function ToNum(ByVal s As String) As Variant
    s = Replace(Replace(s, "%", ""), "％", "")
    If s <> "" And IsNumeric(s) Then
        ToNum = CDbl(s)
    Else
        ToNum = s
    End If
End Function

'---- 通信 ------------------------------------------------------------------
Private Function HttpGetUtf8(ByVal url As String) As String
    Dim body As Variant, stm As Object
    body = TryGet("WinHttp.WinHttpRequest.5.1", url)
    If IsEmpty(body) Then body = TryGet("MSXML2.ServerXMLHTTP.6.0", url)
    If IsEmpty(body) Then body = TryGet("MSXML2.XMLHTTP", url)
    If IsEmpty(body) Then Err.Raise vbObjectError + 1, , "接続できませんでした"

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

'---- シート ---------------------------------------------------------------
Private Function GetOrAddSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrAddSheet = ws
End Function

Private Function PrepareOutSheet(ByVal half As Boolean) As Worksheet
    Dim ws As Worksheet
    If half Then Set ws = GetOrAddSheet(OUT30_SHEET) Else Set ws = GetOrAddSheet(OUT_SHEET)
    If Not APPEND_MODE Then ws.Cells.Clear
    If ws.Range("A1").Value = "" Then
        If half Then
            ws.Range("A1:O1").Value = Array("取得日時", "地点", "日付", "コマ", "開始時刻", "終了時刻", "天気", _
                "気温(℃)", "降水確率(%)", "降水量(mm/h)", "降雪量(cm/h)", "湿度(%)", "風向", "風速(m/s)", "北風成分(m/s)")
            ws.Range("A1:O1").Font.Bold = True
            ws.Columns(1).NumberFormat = "yyyy/mm/dd hh:mm"
            ws.Columns(3).NumberFormat = "yyyy/mm/dd"
            ws.Columns(5).NumberFormat = "yyyy/mm/dd hh:mm"
            ws.Columns(6).NumberFormat = "yyyy/mm/dd hh:mm"
        Else
            ws.Range("A1:N1").Value = Array("取得日時", "地点", "日付", "時", "日時", "天気", _
                "気温(℃)", "降水確率(%)", "降水量(mm/h)", "降雪量(cm/h)", "湿度(%)", "風向", "風速(m/s)", "北風成分(m/s)")
            ws.Range("A1:N1").Font.Bold = True
            ws.Columns(1).NumberFormat = "yyyy/mm/dd hh:mm"
            ws.Columns(3).NumberFormat = "yyyy/mm/dd"
            ws.Columns(5).NumberFormat = "yyyy/mm/dd hh:mm"
        End If
    End If
    Set PrepareOutSheet = ws
End Function
