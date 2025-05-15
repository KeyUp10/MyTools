Option Explicit

Private Sub Workbook_Open()

     Dim MyLastRow As Long
     
     '「XXX」シートをアクティベイトする
     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Activate
     '*******************************
     '最終行を取得する
     MyLastRow = Worksheets("XXX").Cells(Rows.Count, 2).End(xlUp).Row

     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Range(Cells(1, 1), Cells(MyLastRow, 10)).Select
     
     Selection.ClearContents
     
     With Selection.Interior
   
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
     End With
     
     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("XXX").Range("A1").Select
     '*******************************
     
     '「Sharepoint」シートをアクティベイトする
     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Activate
     '*******************************
     '最終行を取得する
     MyLastRow = Worksheets("Sharepoint").Cells(Rows.Count, 2).End(xlUp).Row

     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Range(Cells(1, 1), Cells(MyLastRow, 10)).Select
     
     Selection.ClearContents
     
     With Selection.Interior
   
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
     End With
     
     Application.Workbooks("XXX_Ver.1.01.xlsm").Worksheets("Sharepoint").Range("A1").Select
     '*******************************
     
     Application.DisplayAlerts = False
     
     ThisWorkbook.Save
     
     Application.DisplayAlerts = True
     
End Sub
-----------------------------
Option Explicit

Private Sub Workbook_Open()

     Dim MyLastRow As Long
     
     '「PP」シートをアクティベイトする
     Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Activate
     
     '最終行を取得する
     MyLastRow = Worksheets("PP").Cells(Rows.Count, 1).End(xlUp).Row

     Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("PP").Range(Cells(1, 1), Cells(MyLastRow, 10)).Select
     
     Selection.ClearContents
     
     With Selection.Interior
   
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
     End With

     '「kintone」シートをアクティベイトする
     Application.Workbooks("SearchKintoneData_Ver.1.00.xlsm").Worksheets("kintone").Activate
     
     Application.DisplayAlerts = False
     
     ThisWorkbook.Save
     
     Application.DisplayAlerts = True
     
End Sub
-----------------------------
Option Explicit

Private Sub Workbook_Open()

     Dim MyLastRow As Long
     
     '最終行を取得する
     MyLastRow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row

     Worksheets(1).Range(Cells(1, 1), Cells(MyLastRow, 3)).Select
     
     Selection.ClearContents
     
    With Selection.Interior
   
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
    End With
     
     Worksheets(1).Range("A1").Select
     
     Worksheets(1).Name = "納品物"
     
     Application.DisplayAlerts = False
     
     ThisWorkbook.Save
     
     Application.DisplayAlerts = True
     
End Sub
-----------------------------
Private Sub Workbook_Open()

     Dim MyLastRow As Long
     
     'ワークシートの名前を変更する
     Worksheets(1).Name = "YYYYMMDD"
     
     '最終行を取得する
     MyLastRow = Cells(Rows.Count, 1).End(xlUp).Row

     'Worksheets(1).Range(Cells(1, 1), Cells(MyLastRow, 100)).Select
     ActiveSheet.Cells.Select
     
     With Selection
     
        'セルの値を削除する
        .ClearContents
       
        '罫線を削除する
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
     
     End With
     
     With Selection.Interior
   
        'セルの背景色を削除する
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
     End With

     Worksheets(1).Range("A1").Select
     
     Application.DisplayAlerts = False
     
     ThisWorkbook.Save
     
     Application.DisplayAlerts = True
     
End Sub
-----------------------------
Private Sub Workbook_Open()

     Dim MyLastRow As Long
     
     'ワークシートの名前を変更する
     Worksheets(1).Name = "YYYYMMDD"
     
     '最終行を取得する
     MyLastRow = Cells(Rows.Count, 1).End(xlUp).Row

     'Worksheets(1).Range(Cells(1, 1), Cells(MyLastRow, 100)).Select
     ActiveSheet.Cells.Select
     
     With Selection
     
        'セルの値を削除する
        .ClearContents
       
        '罫線を削除する
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
     
     End With
     
     With Selection.Interior
   
        'セルの背景色を削除する
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
       
     End With

     Worksheets(1).Range("A1").Select
     
     Application.DisplayAlerts = False
     
     ThisWorkbook.Save
     
     Application.DisplayAlerts = True
     
End Sub
