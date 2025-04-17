Attribute VB_Name = "Mod03_ViewFormat"
'***************************************************************************************************
'                 引数に指定したセルから、書式設定の内容を返します。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ アドインから呼び出すプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：選択したセル範囲に対応した書式設定の内容を返します。
'---------------------------------------------------------------------------------------------------
'* 機能説明：スピルにも対応するユーザー定義関数です。
'***************************************************************************************************
Function CheckFormat(selectCell As Range)
    '返り値
    Dim results() As String

    '選択行数を取得
    Dim selectRowCount As Long: selectRowCount = selectCell.Rows.Count

    '選択列数を取得
    Dim selectColumnCount As Long: selectColumnCount = selectCell.Columns.Count
    
    '配列数を決定
    ReDim results(1 To selectRowCount, 1 To selectColumnCount)
    
    '格納
    Dim i As Long, j As Long
    For i = 1 To selectRowCount
        For j = 1 To selectColumnCount
            results(i, j) = selectCell.Cells(i, j).NumberFormatLocal
        Next
    Next

    '返却
    CheckFormat = results
End Function
