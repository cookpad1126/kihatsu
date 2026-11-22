Attribute VB_Name = "FileNameDate"
Option Explicit

'---------------------------- 設定 --------------------------------------------
Private Const TARGET_SHEET As String = ""   ' 貼り付け先シート名（空欄＝実行時のアクティブシート）
Private Const TARGET_CELL As String = "A1"  ' 貼り付け先セル
'------------------------------------------------------------------------------

' ファイル名の先頭8桁（yyyymmdd）を日付として TARGET_CELL に値で貼り付ける
Public Sub ファイル名から日付を取得()
    Dim fn As String, s As String, d As Date, ws As Worksheet

    fn = ThisWorkbook.Name                  ' 例: 20261111_abcd.xlsm
    s = Left$(fn, 8)

    If Not s Like "########" Then
        MsgBox "ファイル名の先頭が yyyymmdd 形式ではありません。" & vbCrLf & fn, vbExclamation
        Exit Sub
    End If

    d = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
    If Format$(d, "yyyymmdd") <> s Then     ' 20261332 のような存在しない日付を弾く
        MsgBox "ファイル名の日付が正しくありません：" & s, vbExclamation
        Exit Sub
    End If

    If TARGET_SHEET = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)
    End If

    With ws.Range(TARGET_CELL)
        .Value = d                          ' 数式ではなく値として入力
        .NumberFormat = "yyyy/mm/dd"
    End With
End Sub
