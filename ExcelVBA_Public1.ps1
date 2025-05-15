Option Explicit

Sub XXX()

Dim i As Long
Dim s As Long
Dim MyLastRow As Long 'Worksheets("Sharepoint")の最終行取得用
Dim t1 As Long 'PO Ownerの最終列取得用
Dim MyLastCol As Long 'Worksheets("PO Owner")の最終列取得用

'1行目の最終列を認識する
'-----------------------------------------
MyLastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
'-----------------------------------------

'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "PO Owner" Then

          t1 = s
          Exit For
         
     End If

Next s
'-----------------------------------------

'「Sharepoint」シートの列を「XXX」シートに張り付ける
'*******************************
'ProjectID
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(1).Copy
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'サービスオーダー番号
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(2).Copy
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 2).PasteSpecial Paste:=xlPasteValues

'ユーザー名
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(3).Copy
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 3).PasteSpecial Paste:=xlPasteValues

'回収
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(4).Copy
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 4).PasteSpecial Paste:=xlPasteValues

'検収書のファイル名
'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(8).Copy
'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 5).PasteSpecial Paste:=xlPasteValues

'PO Owner
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Columns(t1).Copy
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 6).PasteSpecial Paste:=xlPasteValues
'*******************************

'*******************************
'項目を入力する
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 1) = "Project ID"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 2) = "サービスオーダー番号"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 3) = "ユーザー名"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 4) = "回収"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 5) = "検収書のファイル名"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 6) = "PO Owner"
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(1, 7) = "備考"
'*******************************

'G列の対象行に「御社にて回収と伺っております。」を入力する
'*******************************
'最終行を取得する
MyLastRow = Cells(Rows.Count, 1).End(xlUp).Row

'メイン処理
For i = 2 To MyLastRow

     If Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 4) = "無" Then
     
          Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 7).Value = "貴社にて回収と伺っております。"

     End If

Next i
'*******************************

'メールアドレスをローカルパートだけにする
'*******************************
For i = 2 To MyLastRow

     If InStr(Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6), "@") <> 0 Then

          '@の前を抽出する
          Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6) = _
          Left(Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6), InStr(Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6), "@"))

          '@を削除する
          Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6) = _
          Replace(Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Cells(i, 6), "@", "")

     End If

Next i
'*******************************

'C列、D列、E列、F列、G列の幅をAutoFitする
'*******************************
'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("C:C").Select
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("C:C").EntireColumn.AutoFit

'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("D:D").Select
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("D:D").EntireColumn.AutoFit

'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("E:E").Select
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("E:E").EntireColumn.AutoFit

'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("F:F").Select
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("F:F").EntireColumn.AutoFit

'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("G:G").Select
Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Columns("G:G").EntireColumn.AutoFit
'*******************************

'A列の一番上のセルをアクティブにする
'Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Range("A1").Select

End Sub



-----------------------------
Option Explicit

Sub ClosePP()

'------------ 中断前 ------------
Dim i As Long
Dim s As Long
Dim t1 As Long 'Partner Portal処理状況
Dim t2 As Long 'オーダー番号
Dim t3 As Long '客先名
Dim t4 As Long '検収書回収有無
Dim t5 As Long '作業終了日
Dim t6 As Long '進捗
Dim t7 As Long 'PP
Dim MyLastRow As Long 'Worksheets("kintone")の最終行取得用
Dim MyLastRow_53 As Long 'Worksheets("53")の最終列取得用(1回目)
Dim MyLastRow_53_2 As Long 'Worksheets("53")の最終列取得用(2回目)
Dim MyLastCol As Long 'Worksheets("kintone")の最終列取得用
Dim T_Cell As Range
Dim A(100) As String
'------------ 中断前 ------------

'------------ 中断後 ------------
Dim MyLastRow_17 As Long 'Worksheets("17")の最終列取得用
Dim MyLastRow_10 As Long 'Worksheets("10")の最終列取得用
Dim MyLastRow_16 As Long 'Worksheets("16")の最終列取得用
Dim MyLastRow_15 As Long 'Worksheets("15")の最終列取得用
Dim MyLastRow_summary As Long 'Worksheets("summary ")の最終列取得用
'------------ 中断後 ------------

'後続作業のために「PP_Close_Ver.1.00.xlsm」の「kintone」シートをアクティベイトする
'*******************************
Application.Workbooks("PP_Close_Ver.1.02.xlsm").Worksheets("kintone").Activate
'*******************************

'1行目の最終列を認識する
'-----------------------------------------
MyLastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
'-----------------------------------------

'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "Partner Portal処理状況" Then

          t1 = s
          Exit For
         
     End If

Next s
'-----------------------------------------
'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "オーダー番号" Then

          t2 = s
          Exit For
         
     End If

Next s
'-----------------------------------------
'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "客先名" Then

          t3 = s
          Exit For
         
     End If

Next s
'-----------------------------------------
'-----------------------------------------
'For s = 1 To MyLastCol

'     If Cells(1, s) = "検収書回収有無" Then

'          t4 = s
'          Exit For
         
'     End If

'Next s
'-----------------------------------------
'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "作業終了日" Then

          t5 = s
          Exit For
         
     End If

Next s
'-----------------------------------------

'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "進捗" Then

          t6 = s
          Exit For
         
     End If

Next s
'-----------------------------------------

'PP
'-----------------------------------------
For s = 1 To MyLastCol

     If Cells(1, s) = "DPS_No" Then

          t7 = s
          Exit For
         
     End If

Next s
'-----------------------------------------

't3列(客先名)の最終行を認識する
'-----------------------------------------
MyLastRow = ActiveSheet.Cells(Rows.Count, t3).End(xlUp).Row
'-----------------------------------------

'列の高さを規定の値(12)に変更する
'-----------------------------------------
ActiveSheet.Range(Cells(1, 1), Cells(MyLastRow, 17)).Select
Selection.RowHeight = 12
'-----------------------------------------

't6列(進捗)の値が「納品完了 100%」以外であれば、その行を削除する
'-----------------------------------------
For i = 2 To MyLastRow

     If Cells(i, t6) <> "" And Cells(i, t6) <> "納品完了 100%" Then

          ActiveSheet.Rows(i).Delete
          i = i - 1

     End If

Next i
'-----------------------------------------

't5列(作業終了日)の値が空欄であれば、その行を削除する
'-----------------------------------------
For i = 2 To MyLastRow

     If Cells(i, t3) <> "" And Cells(i, t5) = "" Then

          ActiveSheet.Rows(i).Delete
          i = i - 1

     End If

Next i
'-----------------------------------------

't5列(作業終了日)が「本日-7日」未満の値であれば、その行を削除する
'-----------------------------------------
For i = 2 To MyLastRow

     If Cells(i, t5) = "" Then

          Exit For

     End If

     If DateSerial(Year(CDate(Replace(Cells(i, t5), "　", " "))), Month(CDate(Replace(Cells(i, t5), "　", " "))), Day(CDate(Replace(Cells(i, t5), "　", " ")))) _
          < DateSerial(Year(Now), Month(Now), Day(Now)) - 7 Then

          ActiveSheet.Rows(i).Delete
          i = i - 1

     End If

Next i
'-----------------------------------------

't3列(客先名)の最終行を再確認する
'-----------------------------------------
MyLastRow = ActiveSheet.Cells(Rows.Count, t3).End(xlUp).Row
'-----------------------------------------

't7列(DPS_No)が既に作業していれれば、O列に「Close処理済み/処理不要」を入力する
'-----------------------------------------
Dim B As Range
Dim MySearch As String

For i = 2 To MyLastRow

MySearch = ActiveSheet.Cells(i, t7)

Set B = Worksheets("53").Range("C:C").Find(What:=MySearch)

     If B Is Nothing Then

          MsgBox "メインまたは作業対象です"

     Else

          ActiveSheet.Cells(i, t7).Offset(0, 9) = "Close処理済み/処理不要"

     End If

Next i
'-----------------------------------------

Stop

Exit Sub

'保存時のメッセージ抑止
Application.DisplayAlerts = False

'ちらつき防止
Application.ScreenUpdating = False

'検索対象を配列に格納する
'*******************************
For i = 1 To UBound(A)

If Not IsNull(Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("kintone").Cells(i, 2)) Then
If Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("kintone").Cells(i, 2) <> "Ariba" Then

     A(i - 1) = Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("kintone").Cells(i, 2).Value

End If
End If

Next i
'*******************************

'*******************************
For i = 0 To UBound(A)

If A(i) = "" Then

     A(i) = "オーダー番号"

End If

Next i
'*******************************

MsgBox "kintoneレコードからデータを抽出します"

'*******************************
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\5)_VBAツール類\【Data】kintoneレコード.xlsx"
'*******************************

'列の高さを規定の値(19.5)に変更する
'*******************************
ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 57)).Select
Selection.RowHeight = 19.5
'*******************************

'「【Data】kintoneレコード」のAS列をオートフィルターで空白以外を抽出する
'*******************************
ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 45, "<>"
'*******************************

'「【Data】kintoneレコード」をオートフィルターで検索する
'*******************************
ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 3, A, xlFilterValues
'*******************************

'「【Data】kintoneレコード」をオートフィルターの結果をツール側に張り付ける
'*******************************
'オーダー番号
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("C:F").Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'客先名
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(4).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 2).PasteSpecial Paste:=xlPasteValues

'作業開始日
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(5).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 3).PasteSpecial Paste:=xlPasteValues

'作業終了日
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(6).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 4).PasteSpecial Paste:=xlPasteValues

'DPS番号
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AS:AT").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(46).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 5).PasteSpecial Paste:=xlPasteValues

'オーダー番号
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(46).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(47).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 6).PasteSpecial Paste:=xlPasteValues

'数量
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(42).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(43).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 7).PasteSpecial Paste:=xlPasteValues

'SKU番号
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(44).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(45).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 8).PasteSpecial Paste:=xlPasteValues

'SOW
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AM:AN").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(40).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 9).PasteSpecial Paste:=xlPasteValues

'時間帯
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(40).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(41).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 10).PasteSpecial Paste:=xlPasteValues

'進捗
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(2).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 11).PasteSpecial Paste:=xlPasteValues

'ディスパッチ状況(トラブル/懸念事項/引継ぎ事項)
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(7).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 12).PasteSpecial Paste:=xlPasteValues

'レコード番号
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(13).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(14).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 13).PasteSpecial Paste:=xlPasteValues

'YYY支店
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("P:R").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(17).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 14).PasteSpecial Paste:=xlPasteValues

'部署コード
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(17).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(18).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 15).PasteSpecial Paste:=xlPasteValues

'住所
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(18).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(19).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 16).PasteSpecial Paste:=xlPasteValues

'検収書回収有無
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("X:AA").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(25).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 17).PasteSpecial Paste:=xlPasteValues

'検収書提出日
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(25).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(26).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 18).PasteSpecial Paste:=xlPasteValues

'見積番号
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(26).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(27).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 19).PasteSpecial Paste:=xlPasteValues

'機器販売/運搬
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(27).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(28).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 20).PasteSpecial Paste:=xlPasteValues

'受注金額(運搬・設備除外)
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AU:AW").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(48).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 21).PasteSpecial Paste:=xlPasteValues

'FCN作業
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(48).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(49).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 22).PasteSpecial Paste:=xlPasteValues

'FCN金額
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(49).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(50).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 23).PasteSpecial Paste:=xlPasteValues

'デル社PM
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("BC:BD").Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(56).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 24).PasteSpecial Paste:=xlPasteValues

'属性
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(56).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(57).Copy
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 25).PasteSpecial Paste:=xlPasteValues

'ProjectID
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(11).Copy
'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(12).Copy
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Cells(1, 26).PasteSpecial Paste:=xlPasteValues
'*******************************

Application.Workbooks("【Data】kintoneレコード.xlsx").Save
Application.Workbooks("【Data】kintoneレコード.xlsx").Close

MsgBox "完了しました"

'後続作業のために「PP_Close_Ver.1.00.xlsm」の「kintone」シートをアクティベイトする
'*******************************
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("kintone").Activate
'*******************************

'DPS番号をツールに追記するための「Stop」
'*******************************
'Stop
'*******************************

MsgBox "抽出したデータをツールにペーストします"

MyLastRow_53 = Worksheets("53").Cells(Rows.Count, 1).End(xlUp).Row
MyLastRow_summary = Worksheets("summary").Cells(Rows.Count, 2).End(xlUp).Row

'Worksheets("53")のA列にオーダー番号、客先名、DPS番号、作業終了日、メニュー等をペーストする
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Activate

Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 1), Cells(MyLastRow_summary, 1)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 1)
'2024/2/21変更
'*******************************
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 1), Cells(MyLastRow_summary, 1)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("17").Cells(MyLastRow_53 + 1, 1)
'*******************************
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 1), Cells(MyLastRow_summary, 1)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 1)
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 1), Cells(MyLastRow_summary, 1)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Cells(MyLastRow_53 + 1, 1)
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 1), Cells(10, 1)).Copy _
'Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("15").Cells(MyLastRow_53 + 1, 1)

Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 2), Cells(MyLastRow_summary, 2)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 2)
'2024/2/21変更
'*******************************
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 2), Cells(MyLastRow_summary, 2)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("17").Cells(MyLastRow_53 + 1, 2)
'*******************************
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 2), Cells(MyLastRow_summary, 2)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 2)
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 2), Cells(MyLastRow_summary, 2)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Cells(MyLastRow_53 + 1, 2)
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 2), Cells(10, 2)).Copy _
'Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("15").Cells(MyLastRow_53 + 1, 2)

Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 5), Cells(MyLastRow_summary, 5)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 3)
'2024/2/21変更
'*******************************
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 5), Cells(MyLastRow_summary, 5)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("17").Cells(MyLastRow_53 + 1, 3)
'*******************************
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 5), Cells(MyLastRow_summary, 5)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 3)
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 5), Cells(MyLastRow_summary, 5)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Cells(MyLastRow_53 + 1, 3)
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 5), Cells(10, 5)).Copy _
'Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("15").Cells(MyLastRow_53 + 1, 3)

'作業終了日
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 4), Cells(MyLastRow_summary, 4)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 4)

'検収書回収有無
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 17), Cells(MyLastRow_summary, 17)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 6)

'数量
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 7), Cells(MyLastRow_summary, 7)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 7)

'SKU番号
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 8), Cells(MyLastRow_summary, 8)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 8)

'SOW
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 9), Cells(MyLastRow_summary, 9)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 9)

'時間帯
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 10), Cells(MyLastRow_summary, 10)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("53").Cells(MyLastRow_53 + 1, 10)

'Worksheets("10")のF列、G列、H列にYYY支店、部署コード、住所等をペーストする
'YYY支店
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 14), Cells(MyLastRow_summary, 14)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 6)

'部署コード
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 15), Cells(MyLastRow_summary, 15)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 7)

'住所
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 16), Cells(MyLastRow_summary, 16)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_53 + 1, 8)

'Range(Cells(2, t2), Cells(MyLastRow, t2)).Copy Destination:=Worksheets("53").Cells(MyLastRow_53 + 1, 1)
'Range(Cells(2, t3), Cells(MyLastRow, t3)).Copy Destination:=Worksheets("53").Cells(MyLastRow_53 + 1, 2)
'*******************************
'Range(Cells(2, t4), Cells(MyLastRow, t4)).Copy Destination:=Worksheets("53").Cells(MyLastRow_53 + 1, 6)
'*******************************
'Range(Cells(2, t5), Cells(MyLastRow, t5)).Copy
'Worksheets("53").Cells(MyLastRow_53 + 1, 4).PasteSpecial Paste:=xlPasteValues

'Worksheets("53")のA列の最終行を取得する
MyLastRow_53_2 = Worksheets("53").Cells(Rows.Count, 1).End(xlUp).Row

'Worksheets("53")のE列に作業完了/XXX回収可 or 作業完了/XXX回収不可を入力する
'-----------------------------------------
For i = MyLastRow_53 + 1 To MyLastRow_53_2

     If Worksheets("53").Cells(i, 6) = "有" Then

        Worksheets("53").Cells(i, 5) = "作業完了/XXX回収可"

     End If

     If Worksheets("53").Cells(i, 6) = "無" Then

        Worksheets("53").Cells(i, 5) = "作業完了/XXX回収不可"

     End If

Next i
'-----------------------------------------

'Worksheets("53")のA列の最終行を再度取得する
MyLastRow_53_2 = Worksheets("53").Cells(Rows.Count, 1).End(xlUp).Row

'各WorksheetのD列の最終行を取得する
MyLastRow_17 = Worksheets("17").Cells(Rows.Count, 4).End(xlUp).Row
MyLastRow_10 = Worksheets("10").Cells(Rows.Count, 4).End(xlUp).Row
MyLastRow_16 = Worksheets("16").Cells(Rows.Count, 4).End(xlUp).Row
MyLastRow_15 = Worksheets("15").Cells(Rows.Count, 4).End(xlUp).Row

'各Worksheetにオーダー番号、客先名、Dispatchをペーストする
'Range(Worksheets("53").Cells(MyLastRow_53 + 1, 1), Worksheets("53").Cells(MyLastRow_53_2, 3)).Copy Destination:=Worksheets("17").Cells(MyLastRow_17 + 1, 1)
'Range(Worksheets("53").Cells(MyLastRow_53 + 1, 1), Worksheets("53").Cells(MyLastRow_53_2, 3)).Copy Destination:=Worksheets("10").Cells(MyLastRow_10 + 1, 1)
'Range(Worksheets("53").Cells(MyLastRow_53 + 1, 1), Worksheets("53").Cells(MyLastRow_53_2, 3)).Copy Destination:=Worksheets("16").Cells(MyLastRow_16 + 1, 1)
Range(Worksheets("53").Cells(MyLastRow_53 + 1, 1), Worksheets("53").Cells(MyLastRow_53_2, 3)).Copy Destination:=Worksheets("15").Cells(MyLastRow_15 + 1, 1)

'各Worksheetに作業開始日をペーストする
'-----------------------------------------
'2024/2/21変更
'*******************************
'Worksheets("17")に作業開始日をペーストする
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Activate
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 3), Cells(MyLastRow_summary, 3)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("17").Cells(MyLastRow_17 + 1, 4)
'*******************************

'Worksheets("10")に作業開始日をペーストする
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 3), Cells(MyLastRow_summary, 3)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("10").Cells(MyLastRow_10 + 1, 4)

'Worksheets("16")に作業開始日をペーストする
Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("summary").Range(Cells(2, 3), Cells(MyLastRow_summary, 3)).Copy _
Destination:=Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Cells(MyLastRow_10 + 1, 4)
'-----------------------------------------

'Worksheets("10")のA列の最終行を取得する
MyLastRow_10 = Worksheets("10").Cells(Rows.Count, 5).End(xlUp).Row

'Worksheets("10")のE列にエンジニアコードを入力する
'-----------------------------------------
For i = MyLastRow_10 + 1 To MyLastRow_53_2

     Select Case Worksheets("10").Cells(i, 6).Value

     Case "関東"

        Worksheets("10").Cells(i, 5) = "111111"
       
     Case "長野"

        Worksheets("10").Cells(i, 5) = "111111"
       
     Case "中部"

        Worksheets("10").Cells(i, 5) = "111111"
       
     Case "静岡"

        Worksheets("10").Cells(i, 5) = "111111"

     Case "北陸"

        Worksheets("10").Cells(i, 5) = "222222"

     Case "関西"

        Worksheets("10").Cells(i, 5) = "222222"
     
     Case "中国"

        Worksheets("10").Cells(i, 5) = "222222"

     Case "四国"

        Worksheets("10").Cells(i, 5) = "222222"

     Case "九州"

        Worksheets("10").Cells(i, 5) = "333333"

     Case "北海道"

        Worksheets("10").Cells(i, 5) = "444444"

     Case "東北"

        Worksheets("10").Cells(i, 5) = "444444"

     Case "新潟"

        Worksheets("10").Cells(i, 5) = "444444"

     Case Else

        Worksheets("10").Cells(i, 5) = ""

     End Select

Next i
'-----------------------------------------

MsgBox "完了しました"

MsgBox "PPに投入するExcelファイルのブランクを作成します"

'15_FC.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\0)_15\15_FC.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear

For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
'ActiveSheet.Cells(i, 4).Value = "ETA_DEFERRED"
ActiveSheet.Cells(i, 4).Value = "CUSTOMER_CONTACTED"
'ActiveSheet.Cells(i, 6).Value = "15"
ActiveSheet.Cells(i, 6).Value = "C04"
ActiveSheet.Cells(i, 7).Value = "FC"

Next i
'-----------------------------------------

ActiveWorkbook.Save
ActiveWorkbook.Close

'10_16_日程Fix_エンジニアアサイン.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\1)_10_16\10_16_日程Fix_エンジニアアサイン.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 12), Cells(Rows.Count, 12)).Select
Selection.Clear

For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
ActiveSheet.Cells(i, 4).Value = "ETA_PROVIDED"
'ActiveSheet.Cells(i, 6).Value = "16"
ActiveSheet.Cells(i, 6).Value = "C01"
ActiveSheet.Cells(i, 7).Value = "日程Fix"
ActiveSheet.Cells(i, 12).Value = ""

Next i

For i = (MyLastRow_53_2 - MyLastRow_53 + 2) To (MyLastRow_53_2 - MyLastRow_53) * 2 + 1

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
ActiveSheet.Cells(i, 4).Value = "ETA_PROVIDED"
'ActiveSheet.Cells(i, 6).Value = "10"
ActiveSheet.Cells(i, 6).Value = "E01"
ActiveSheet.Cells(i, 7).Value = "エンジニアアサイン"
ActiveSheet.Cells(i, 12).Value = "111111"

Next i

'2024/1/26追加
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Range(Cells(MyLastRow_53 + 1, 3), Cells(MyLastRow_53_2, 3)).Copy _
Destination:=ActiveSheet.Cells(2, 1)
'-----------------------------------------

ActiveWorkbook.Save
ActiveWorkbook.Close

'17_53_ONSITE_COMPLETE.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\2)_17_53\17_53_ONSITE_COMPLETE.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear

'For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

'ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
'ActiveSheet.Cells(i, 4).Value = "ON_SITE"
'ActiveSheet.Cells(i, 6).Value = "17"
'ActiveSheet.Cells(i, 7).Value = ""

'Next i

'For i = (MyLastRow_53_2 - MyLastRow_53 + 2) To (MyLastRow_53_2 - MyLastRow_53) * 2 + 1
For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "CALL_CLOSURE"
ActiveSheet.Cells(i, 4).Value = "SERVICE_COMPLETED"
'ActiveSheet.Cells(i, 6).Value = "53"
ActiveSheet.Cells(i, 6).Value = "E04"
ActiveSheet.Cells(i, 7).Value = "作業完了/XXX回収可"

Next i
'-----------------------------------------

MsgBox "完了しました"

ActiveWorkbook.Save
ActiveWorkbook.Close

'保存時のメッセージ抑止解除
Application.DisplayAlerts = True

'ちらつき防止解除
Application.ScreenUpdating = True

End Sub



-----------------------------
Option Explicit

Sub Make_PP_Sheets()

Dim i As Long
Dim MyLastRow_53 As Long 'Worksheets("53")の最終列取得用(1回目)
Dim MyLastRow_53_2 As Long 'Worksheets("53")の最終列取得用(2回目)

'Worksheets("53")のB列、C列の最終行を取得する
MyLastRow_53 = Worksheets("53").Cells(Rows.Count, 2).End(xlUp).Row
MyLastRow_53_2 = Worksheets("53").Cells(Rows.Count, 3).End(xlUp).Row

'保存時のメッセージ抑止
Application.DisplayAlerts = False

'ちらつき防止
Application.ScreenUpdating = False

MsgBox "PPに投入するExcelファイルのブランクを作成します"

'15_FC.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\0)_15\15_FC.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear

For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
'ActiveSheet.Cells(i, 4).Value = "ETA_DEFERRED"
ActiveSheet.Cells(i, 4).Value = "CUSTOMER_CONTACTED"
'ActiveSheet.Cells(i, 6).Value = "15"
ActiveSheet.Cells(i, 6).Value = "C04"
ActiveSheet.Cells(i, 7).Value = "FC"

Next i
'-----------------------------------------

ActiveWorkbook.Save
ActiveWorkbook.Close

'10_16_日程Fix_エンジニアアサイン.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\1)_10_16\10_16_日程Fix_エンジニアアサイン.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 12), Cells(Rows.Count, 12)).Select
Selection.Clear

For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
ActiveSheet.Cells(i, 4).Value = "ETA_PROVIDED"
'ActiveSheet.Cells(i, 6).Value = "16"
ActiveSheet.Cells(i, 6).Value = "C01"
ActiveSheet.Cells(i, 7).Value = "日程Fix"
ActiveSheet.Cells(i, 12).Value = ""

Next i

For i = (MyLastRow_53_2 - MyLastRow_53 + 2) To (MyLastRow_53_2 - MyLastRow_53) * 2 + 1

ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
ActiveSheet.Cells(i, 4).Value = "ETA_PROVIDED"
'ActiveSheet.Cells(i, 6).Value = "10"
ActiveSheet.Cells(i, 6).Value = "E01"
ActiveSheet.Cells(i, 7).Value = "エンジニアアサイン"
ActiveSheet.Cells(i, 12).Value = "111111"

Next i

'2024/1/26追加
'Application.Workbooks("PP_Close_Ver.1.02xlsm").Worksheets("16").Range(Cells(MyLastRow_53 + 1, 3), Cells(MyLastRow_53_2, 3)).Copy _
Destination:=ActiveSheet.Cells(2, 1)
'-----------------------------------------

ActiveWorkbook.Save
ActiveWorkbook.Close

'17_53_ONSITE_COMPLETE.xlsxのブランクシートを作成する
'-----------------------------------------
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\3)_PP締め業務\0)_PP一括処理\2)_17_53\17_53_ONSITE_COMPLETE.xlsx"
ActiveSheet.Range(Cells(2, 1), Cells(Rows.Count, 1)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 2), Cells(Rows.Count, 2)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 3), Cells(Rows.Count, 3)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 4), Cells(Rows.Count, 4)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 5), Cells(Rows.Count, 5)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 6), Cells(Rows.Count, 6)).Select
Selection.Clear
ActiveSheet.Range(Cells(2, 7), Cells(Rows.Count, 7)).Select
Selection.Clear

'For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

'ActiveSheet.Cells(i, 2).Value = "SERVICE_UPDATES"
'ActiveSheet.Cells(i, 4).Value = "ON_SITE"
'ActiveSheet.Cells(i, 6).Value = "17"
'ActiveSheet.Cells(i, 7).Value = ""

'Next i

'For i = (MyLastRow_53_2 - MyLastRow_53 + 2) To (MyLastRow_53_2 - MyLastRow_53) * 2 + 1
For i = 2 To (MyLastRow_53_2 - MyLastRow_53 + 1)

ActiveSheet.Cells(i, 2).Value = "CALL_CLOSURE"
ActiveSheet.Cells(i, 4).Value = "SERVICE_COMPLETED"
'ActiveSheet.Cells(i, 6).Value = "53"
ActiveSheet.Cells(i, 6).Value = "E04"
ActiveSheet.Cells(i, 7).Value = "作業完了/XXX回収可"

Next i
'-----------------------------------------

MsgBox "完了しました"

ActiveWorkbook.Save
ActiveWorkbook.Close

'保存時のメッセージ抑止解除
Application.DisplayAlerts = True

'ちらつき防止解除
Application.ScreenUpdating = True

End Sub



-----------------------------
Option Explicit

Sub SearchKintoneData()

Dim i As Long
Dim A(100) As String

'保存時のメッセージ抑止
Application.DisplayAlerts = False

'ちらつき防止
Application.ScreenUpdating = False

'検索対象を配列に格納する
'*******************************
For i = 1 To UBound(A)

If Not IsNull(Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("kintone").Cells(i, 1)) Then

     A(i - 1) = Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("kintone").Cells(i, 1).Value

End If

Next i
'*******************************

'*******************************
For i = 0 To UBound(A)

If A(i) = "" Then

     A(i) = "オーダー番号"

End If

Next i
'*******************************

'*******************************
Workbooks.Open Filename:= _
"\\VDI01\DFS\home04\dmorikawa2\Documents\5)_VBAツール類\【Data】kintoneレコード.xlsx"
'*******************************

'列の高さを規定の値(19.5)に変更する
'*******************************
ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 57)).Select
Selection.RowHeight = 19.5
'*******************************

'「【Data】kintoneレコード」AS列をオートフィルターで空白以外を抽出する
'*******************************
'ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 45, "<>"
'*******************************

'「【Data】kintoneレコード」AS列のオートフィルターを解除する
'*******************************
ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 45
'*******************************

'「【Data】kintoneレコード」をオートフィルターで検索する
'*******************************
'ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 3, A, xlFilterValues
ActiveWorkbook.ActiveSheet.Range("A1").AutoFilter 46, A, xlFilterValues
'*******************************

'「【Data】kintoneレコード」をオートフィルターの結果をツール側に張り付ける
'*******************************
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("C:F").Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(4).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 2).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(5).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 3).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(6).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 4).PasteSpecial Paste:=xlPasteValues

Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AS:AT").Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 5).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(46).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 6).PasteSpecial Paste:=xlPasteValues

Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(13).Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 7).PasteSpecial Paste:=xlPasteValues

Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(42).Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 8).PasteSpecial Paste:=xlPasteValues

Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(44).Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 9).PasteSpecial Paste:=xlPasteValues

Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AM:AN").Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 10).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(40).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 11).PasteSpecial Paste:=xlPasteValues
'*******************************
'*******************************
Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns("AI:BB").Copy
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 12).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(36).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 13).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(37).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 14).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(38).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 15).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(39).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 16).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(40).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 17).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(41).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 18).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(42).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 19).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(43).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 20).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(44).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 21).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(45).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 22).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(46).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 23).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(47).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 24).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(48).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 25).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(49).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 26).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(50).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 27).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(51).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 28).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(52).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 29).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(53).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 30).PasteSpecial Paste:=xlPasteValues

'Application.Workbooks("【Data】kintoneレコード.xlsx").ActiveSheet.Columns(54).Copy
'Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Cells(1, 31).PasteSpecial Paste:=xlPasteValues
'*******************************

Application.Workbooks("【Data】kintoneレコード.xlsx").Save
Application.Workbooks("【Data】kintoneレコード.xlsx").Close

'後続作業のために「SearchKintoneData_Ver.1.00.xlsm」の「kintone」シートをアクティベイトする
'*******************************
Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("kintone").Activate
'*******************************

ActiveWorkbook.Save
'ActiveWorkbook.Close

'保存時のメッセージ抑止解除
Application.DisplayAlerts = True

'ちらつき防止解除
Application.ScreenUpdating = True

End Sub



-----------------------------
Option Explicit

Sub OrderSupportShape()

Dim i As Long
Dim MyLastRow As Long

'項目のために１行追加する
If Cells(1, 1) <> "アイテムの説明" Then

     Rows("1:1").Select
     Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

End If

'項目を入力する
Cells(1, 1) = "アイテムの説明"
Cells(1, 2) = "アイテム番号"
Cells(1, 3) = "数量"

'最終行を取得する
MyLastRow = Cells(Rows.Count, 1).End(xlUp).Row

'メイン処理
For i = 1 To MyLastRow

     If Left(Cells(i, 1), 7) = "アイテム番号:" Then
     
          Cells(i, 1).Offset(-1, 1).Value = Trim(Mid(Cells(i, 1), 8))
          Rows(i).Delete
         
     End If

Next i

'A列、B列、C列の幅をAutoFitする
Columns("A:A").Select
Columns("A:A").EntireColumn.AutoFit

Columns("B:B").Select
Columns("B:B").EntireColumn.AutoFit

Columns("C:C").Select
Columns("C:C").EntireColumn.AutoFit

'A列の一番上のセルをアクティブにする
Worksheets(1).Range("A1").Select

End Sub



-----------------------------
Option Explicit

Sub ArrayTally()

On Error GoTo MyError

Dim MyFlag  As Boolean
Dim MyTarget As String
Dim i As Long
Dim k As Long
Dim MyLastRow As Long
Dim ItemNum As Range
Dim buf As String
Dim col As String
Dim MyFilePath As String
Dim MyArray() As String

'B列の最終行の値を取得する
MyLastRow = Cells(Rows.Count, 2).End(xlUp).Row

If MyLastRow = 1 Then

     MsgBox "B列にアイテム番号のデータがありません"
     
     Exit Sub

End If

'ユーザ定義関数FilePathPickerを呼び出し、戻り値を文字列型変数MyFilePathに格納する
MyFilePath = FilePathPicker()

'MyFilePathを（サイレントに）開く
Open MyFilePath For Input As #1

'繰り返し処理開始
'---------------------------------------------
Do Until EOF(1)

L1:

'配列の値をクリアする
Erase MyArray

'フラグをFalseに戻す
MyFlag = False

Line Input #1, buf

MyArray = Split(buf, ",")

'配列MyArrayの要素のうち、数値の入っている要素番号の値をLong型変数kに格納する
For i = 0 To UBound(MyArray)

     If IsNumeric(MyArray(i)) Then

          k = i
         
     End If
     
'Stop

'想定の情報をカウント対象としているかのチェック
     If MyArray(i) Like "?ASSY" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If

     If MyArray(i) Like "?SRV" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
    If MyArray(i) Like "?DPK" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?KIT" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?Kit" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?INFO" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?PWR SPLY" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?LOM" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?CORD" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?Cord" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?DIMM" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?TRPM" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?220V TO 110V CONVERSION" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If
     
     If MyArray(i) Like "?PRC" Then
   
          'フラグをTrueにする
          MyFlag = True
     
     End If

Next i

'Stop

'メイン処理
If InStr(MyArray(0), ":") <> 0 Then
If IsNumeric(MyArray(k)) Then
If MyArray(0) <> "" Then

col = InStr(MyArray(0), ":")

MyTarget = RTrim(Mid(MyArray(0), 1, col - 1))

     If Len(MyTarget) = 9 Then

          MyTarget = Right(MyTarget, 8)

     End If

Set ItemNum = Range(Cells(2, 2), Cells(MyLastRow, 2)).Find(what:=MyTarget)

'検索後の処理
'---------------------------------------------
If ItemNum Is Nothing Then

     '検索対象のセルが見つからない場合、繰り返し処理の先頭に戻る
     GoTo L1

Else

'検索対象のセルが見った場合の処理
ItemNum.Offset(0, 1) = ItemNum.Offset(0, 1) + MyArray(k)

     If MyFlag = False Then
     
          'セルを黄色く網掛する
          ItemNum.Offset(0, 1).Select
         
          With Selection.Interior
         
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
               .Color = 65535
               .TintAndShade = 0
               .PatternTintAndShade = 0
               
          End With
     
     End If

End If
'---------------------------------------------

End If
End If
End If

Loop

'MyFilePathを（サイレントに）閉じる
Close #1

Exit Sub

MyError:

MsgBox "エラーが発生しましたので、処理を止めます"

Stop

End Sub

Function FilePathPicker() As String

Dim MyDialog As FileDialog

Set MyDialog = Application.FileDialog(msoFileDialogFilePicker)

With MyDialog

     .InitialFileName = ""
     .Filters.Clear
     .Filters.Add "csv", "*.csv", 1
     
End With

If MyDialog.Show Then

     FilePathPicker = MyDialog.SelectedItems(1)

Else

     MsgBox "キャンセルされました"
     
     Stop
     
End If

Set MyDialog = Nothing

End Function



-----------------------------
Option Explicit

Sub Test1()

'A列アイテムの説明のセル背景色で並び替える
'-----------------------------------------
With Sheets(1).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnCellColor, Order:=xlDescending
        '.SortFields.Add Key:=Range("A1"), SortOn:=xlSortonfontcolor, Order:=xlAscending
        '.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
End With

With Sheets(1).Sort
        .SetRange Range("A2:C10000")
        .Header = xlNo
        .MatchCase = False
        .Apply
End With
'-----------------------------------------

End Sub




-----------------------------
Option Explicit

Sub ShapeFormats()

Dim Z As String
Dim ZZ As String
Dim MyLastRow As Long
Dim i As Long
Dim MySheetName As String
Dim MyFileName As String

'A列の最終行を認識する
'-----------------------------------------
MyLastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
'-----------------------------------------

'列の幅を各々の規定の値に変更する
'-----------------------------------------
'ActiveSheet.Columns("A:A").Select
'Selection.ColumnWidth = 17.25

'ActiveSheet.Columns("B:B").Select
'Selection.ColumnWidth = 20.25

'ActiveSheet.Columns("C:C").Select
'Selection.ColumnWidth = 24.5

'ActiveSheet.Columns("D:D").Select
'Selection.ColumnWidth = 16.13

'ActiveSheet.Columns("E:E").Select
'Selection.ColumnWidth = 8.38

'ActiveSheet.Columns("F:F").Select
'Selection.ColumnWidth = 10.25

'ActiveSheet.Columns("G:G").Select
'Selection.ColumnWidth = 8.38

'ActiveSheet.Columns("H:H").Select
'Selection.ColumnWidth = 13.38

'ActiveSheet.Columns("I:I").Select
'Selection.ColumnWidth = 13.38

'ActiveSheet.Columns("J:J").Select
'Selection.ColumnWidth = 9.38

'ActiveSheet.Columns("K:K").Select
'Selection.ColumnWidth = 11.38
'-----------------------------------------

'列の高さを規定の値(12)に変更する
'-----------------------------------------
'MyLastRow = Cells(Rows.Count, 10).End(xlUp).Row
'ActiveSheet.Range(Cells(1, 1), Cells(MyLastRow, 11)).Select
ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 11)).Select
Selection.RowHeight = 12
'-----------------------------------------

'A列タイトル、J列受注日で昇順に並び替える
'-----------------------------------------
With Sheets(1).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range("J1"), SortOn:=xlSortOnValues, Order:=xlAscending
        '.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
End With

With Sheets(1).Sort
        .SetRange Range("A2:CO10000")
        .Header = xlNo
        .MatchCase = False
        .Apply
End With
'-----------------------------------------

'シート名を当日の日付に変更する
'-----------------------------------------
If Month(Now) < 10 Then

Z = "0"

End If

If Day(Now) < 10 Then

ZZ = "0"

End If

MySheetName = Year(Date) & Z & Month(Date) & ZZ & Day(Date)
ActiveWorkbook.ActiveSheet.Name = MySheetName
'-----------------------------------------

'値が入っている範囲のセルの書式を指定する
'-----------------------------------------
ActiveSheet.Range(Cells(1, 1), Cells(MyLastRow, 93)).Select

With Selection.Font

        .Name = "游ゴシック"
        .FontStyle = "標準"
        .Size = 10

End With
   
With Selection
   
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False

End With

With Selection

        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter

End With
'-----------------------------------------

'罫線を引く
'-----------------------------------------
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'-----------------------------------------

'セルの書式を設定する
'-----------------------------------------
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'-----------------------------------------

'最上段の下部に二重罫線を引く
'-----------------------------------------
    ActiveSheet.Range(Cells(1, 1), Cells(1, 93)).Select
   
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
'-----------------------------------------

'A列の値に「次回用」があれば、その行を削除する
'A列の値に「次回用」があれば、メッセージを表示する
'-----------------------------------------
'For i = 1 To MyLastRow

'     If Cells(i, 1) = "次回用" Then

'          'ActiveSheet.Row(i).Delete
'          MsgBox "次回用が紛れているので、手動で削除してください"

'     End If

'Next i
'-----------------------------------------

'アクティブシートのA1セルをアクティブにする
ActiveSheet.Cells(1, 1).Select

'ファイルを保存するためのウィンドウを起動する
'MyFileName = Application.GetSaveAsFilename(InitialFileName:="EDT" + MySheetName + "YYY", FileFilter:="Excelファイル,*.xlsx*")
'ActiveWorkbook.SaveAs MyFileName

End Sub



-----------------------------
Option Explicit

'Sub getCSV()
'
'On Error GoTo MyError
'
'Dim MyFlag  As Boolean
'Dim MyTarget As String
'Dim i As Long
'Dim k As Long
'Dim MyLastRow As Long
'Dim ItemNum As Range
'Dim buf As String
'Dim col As String
'Dim MyFilePath As String
'Dim MyArray() As String
'
''ユーザ定義関数FilePathPickerを呼び出し、戻り値を文字列型変数MyFilePathに格納する
'MyFilePath = FilePathPicker()
'
''MyFilePathを（サイレントに）開く
'Open MyFilePath For Input As #1
'
'Stop
'
''繰り返し処理開始
''---------------------------------------------
'Do Until EOF(1)
'
'Line Input #1, buf
'
'Stop
'
'MyArray = Split(buf, ",")
'
'ThisWorkbook.Sheets("YYYYMMDD").Cells(1, 1) = MyArray(0)
'
'Stop
'
'Loop
''---------------------------------------------
'
''MyFilePathを（サイレントに）閉じる
'Close #1
'
'Exit Sub
'
'MyError:
'
'MsgBox "エラーが発生しましたので、処理を止めます"
'
'Stop
'
'End Sub

Function FilePathPicker() As String

Dim MyDialog As FileDialog

Set MyDialog = Application.FileDialog(msoFileDialogFilePicker)

With MyDialog

     .InitialFileName = ""
     .Filters.Clear
     .Filters.Add "csv", "*.csv", 1
     
End With
Stop
If MyDialog.Show Then

     FilePathPicker = MyDialog.SelectedItems(1)

Else

     MsgBox "キャンセルされました"
     
     Stop
     
End If

Set MyDialog = Nothing

End Function




-----------------------------
Option Explicit

Sub ShapeFormats()

Dim Z As String
Dim ZZ As String
Dim MyLastRow As Long
Dim i As Long
Dim MySheetName As String
Dim MyFileName As String

'最終行を認識する
'-----------------------------------------
MyLastRow = ActiveSheet.Cells(Rows.Count, 10).End(xlUp).Row
'-----------------------------------------

'列の幅を各々の規定の値に変更する
'-----------------------------------------
ActiveSheet.Columns("A:A").Select
Selection.ColumnWidth = 17.25

ActiveSheet.Columns("B:B").Select
Selection.ColumnWidth = 20.25

ActiveSheet.Columns("C:C").Select
Selection.ColumnWidth = 24.5

ActiveSheet.Columns("D:D").Select
Selection.ColumnWidth = 16.13

ActiveSheet.Columns("E:E").Select
Selection.ColumnWidth = 8.38

ActiveSheet.Columns("F:F").Select
Selection.ColumnWidth = 10.25

ActiveSheet.Columns("G:G").Select
Selection.ColumnWidth = 8.38

ActiveSheet.Columns("H:H").Select
Selection.ColumnWidth = 13.38

ActiveSheet.Columns("I:I").Select
Selection.ColumnWidth = 13.38

ActiveSheet.Columns("J:J").Select
Selection.ColumnWidth = 9.38

ActiveSheet.Columns("K:K").Select
Selection.ColumnWidth = 11.38
'-----------------------------------------

'列の高さを規定の値(12)に変更する
'-----------------------------------------
'MyLastRow = Cells(Rows.Count, 10).End(xlUp).Row
'ActiveSheet.Range(Cells(1, 1), Cells(MyLastRow, 11)).Select
ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 11)).Select
Selection.RowHeight = 12
'-----------------------------------------

'J列受注日で昇順に並び替える
'-----------------------------------------
With Sheets(1).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("J1"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending
        '.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
End With

With Sheets(1).Sort
        .SetRange Range("A2:K10000")
        .Header = xlNo
        .MatchCase = False
        .Apply
End With
'-----------------------------------------

'シート名を当日の日付に変更する
'-----------------------------------------
If Month(Now) < 10 Then

Z = "0"

End If

If Day(Now) < 10 Then

ZZ = "0"

End If

MySheetName = Year(Date) & Z & Month(Date) & ZZ & Day(Date)
ActiveWorkbook.ActiveSheet.Name = MySheetName
'-----------------------------------------

'値が入っている範囲のセルの書式を指定する
'-----------------------------------------
ActiveSheet.Range(Cells(1, 1), Cells(MyLastRow, 11)).Select

With Selection.Font

        .Name = "游ゴシック"
        .FontStyle = "標準"
        .Size = 10

End With
   
With Selection
   
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False

End With

With Selection

        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter

End With
'-----------------------------------------

'罫線を引く
'-----------------------------------------
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'-----------------------------------------

'セルの書式を設定する
'-----------------------------------------
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'-----------------------------------------

'最上段の下部に二重罫線を引く
'-----------------------------------------
    ActiveSheet.Range(Cells(1, 1), Cells(1, 11)).Select
   
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
'-----------------------------------------

'A列の値に「次回用」があれば、その行を削除する
'A列の値に「次回用」があれば、メッセージを表示する
'-----------------------------------------
For i = 1 To MyLastRow

     If Cells(i, 1) = "次回用" Then

          'ActiveSheet.Row(i).Delete
          MsgBox "次回用が紛れているので、手動で削除してください"

     End If

Next i
'-----------------------------------------

'アクティブシートのA1セルをアクティブにする
ActiveSheet.Cells(1, 1).Select

'ファイルを保存するためのウィンドウを起動する
'MyFileName = Application.GetSaveAsFilename(InitialFileName:="EDT" + MySheetName + "YYY", FileFilter:="Excelファイル,*.xlsx*")
'ActiveWorkbook.SaveAs MyFileName

End Sub



-----------------------------
Option Explicit

'Sub getCSV()
'
'On Error GoTo MyError
'
'Dim MyFlag  As Boolean
'Dim MyTarget As String
'Dim i As Long
'Dim k As Long
'Dim MyLastRow As Long
'Dim ItemNum As Range
'Dim buf As String
'Dim col As String
'Dim MyFilePath As String
'Dim MyArray() As String
'
''ユーザ定義関数FilePathPickerを呼び出し、戻り値を文字列型変数MyFilePathに格納する
'MyFilePath = FilePathPicker()
'
''MyFilePathを（サイレントに）開く
'Open MyFilePath For Input As #1
'
'Stop
'
''繰り返し処理開始
''---------------------------------------------
'Do Until EOF(1)
'
'Line Input #1, buf
'
'Stop
'
'MyArray = Split(buf, ",")
'
'ThisWorkbook.Sheets("YYYYMMDD").Cells(1, 1) = MyArray(0)
'
'Stop
'
'Loop
''---------------------------------------------
'
''MyFilePathを（サイレントに）閉じる
'Close #1
'
'Exit Sub
'
'MyError:
'
'MsgBox "エラーが発生しましたので、処理を止めます"
'
'Stop
'
'End Sub

Function FilePathPicker() As String

Dim MyDialog As FileDialog

Set MyDialog = Application.FileDialog(msoFileDialogFilePicker)

With MyDialog

     .InitialFileName = ""
     .Filters.Clear
     .Filters.Add "csv", "*.csv", 1
     
End With
Stop
If MyDialog.Show Then

     FilePathPicker = MyDialog.SelectedItems(1)

Else

     MsgBox "キャンセルされました"
     
     Stop
     
End If

Set MyDialog = Nothing

End Function
