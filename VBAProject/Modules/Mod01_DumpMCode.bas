Attribute VB_Name = "Mod01_DumpMCode"
'***************************************************************************************************
'                 選択したBookにあるPowerQueryのソースコードを指定フォルダに出力します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ セルの名前一覧 ■■■
'***************************************************************************************************
Public Const RangeName_BeforePathName As String = "前回の出力先パス"



'***************************************************************************************************
'                           ■■■ アドインから呼び出すプロシージャ名 ■■■
'***************************************************************************************************
'* 機能　　：一連の処理をまとめて、PowerQueryのソースコードを指定フォルダに出力します
'---------------------------------------------------------------------------------------------------
'* 注意事項：現時点では、グループ名の取得に非対応です
'***************************************************************************************************
Sub StartDumpPowerQuerySource()
    '開いているBook名を配列に格納
    Dim OpeningBooksList: OpeningBooksList = OpeningBooks

    'リストをタスクダイアログに渡します
    tdl01_DumpForm.TaskDialogShow OpeningBooksList
End Sub



'***************************************************************************************************
'                                   ■■■ 各種情報収集 ■■■
'***************************************************************************************************
'* 機能　　：現在開いているBook名を確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：現在開いているBook名の1次元配列
'***************************************************************************************************
Private Function OpeningBooks()
    '必要な変数を用意
    Dim wb As Workbook
    Dim bookNames() As String
    Dim i As Long

    '開いているBook数で、配列を拡張
    ReDim bookNames(1 To Application.Workbooks.Count)

    '順次、入れていきます
    i = 1
    For Each wb In Application.Workbooks
        bookNames(i) = wb.Name
        i = i + 1
    Next

    '返却
    OpeningBooks = bookNames
End Function

'***************************************************************************************************
'* 機能　　：指定Bookから、PowerQueryコードとクエリ名を取得します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：下記のような2次元配列で返されます
'            列
'               1:クエリ名
'               2:コメント
'               3:PowerQueryコード
'            行
'               作られているクエリ数
'
'* 引数　　：TargetBookName   取得したいBook名
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ここでは、VBAネイティブによる、PowerQueryのMコード出力を行います
'***************************************************************************************************
Public Function GetPowerQueryCode(ByVal targetBookName As String)
    '必要な変数を用意
    Dim wb As Workbook: Set wb = Workbooks(targetBookName)
    Dim queryCount As Long: queryCount = wb.Queries.Count
    Dim resultArray
    Dim i As Long
    Dim pq As WorkbookQuery

    'PowerQueryが設定されていない場合はここで、Stop
    If queryCount = 0 Then
        MsgBox "このBookにはPower Queryが定義されていません。", vbCritical, "Not found"
        Exit Function
    Else
        'クエリ数分、拡張
        ReDim resultArray(1 To queryCount, 1 To 3)
    End If

    'クエリごとに配列へ格納
    i = 1
    For Each pq In wb.Queries
        resultArray(i, PowerQueryInfos.QueryName) = pq.Name      'クエリ名
        resultArray(i, PowerQueryInfos.Comment) = pq.Description 'コメント
        resultArray(i, PowerQueryInfos.M_Code) = pq.Formula      'Mコード
        
        'カウントUP
        i = i + 1
    Next pq

    '返却
    GetPowerQueryCode = resultArray

End Function
