VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tdl01_DumpForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "tdl01_DumpForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "tdl01_DumpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'               選択したBookにあるPowerQueryのソースコードを指定フォルダに出力します
'               Excel標準のユーザーフォームは使用しません
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                             ■■■ 動作に必要な変数定義 ■■■
'***************************************************************************************************
' 機能：様々な設定値等をタスクダイアログのイベント等からでもアクセスできるように定義します。
'===================================================================================================
Private WithEvents TaskDialogForDumpForm    As cTaskDialog      'タスクダイアログ本体
Attribute TaskDialogForDumpForm.VB_VarHelpID = -1

Private OpeningBooksList                                        '現在開いているBookのリスト
Private BeforeMARQUEE                       As Boolean          '以前にMARQUEEをしたか？

'ボタンID
Private Enum ButtonAchievementID
    出力 = 101
    フォルダを選択
    閉じる
End Enum

'PowerQueryの情報
Private Enum PowerQuery
    QueryName = 1
    Comment
    M_Code
End Enum

'***************************************************************************************************
'                                   ■■■ 表示構成 ■■■
'***************************************************************************************************
'* 機能　　：PowerQueryCodeを出力する際の設定ウィンドウを表示させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：BookNameList   開いているExcelBookリスト
'***************************************************************************************************
Sub TaskDialogShow(BookNameList)
    '必要な変数を設定
    Dim i As Long
    OpeningBooksList = BookNameList
    BeforeMARQUEE = True


    'オブジェクトの生成
    Set TaskDialogForDumpForm = New cTaskDialog

    '設定を施す
    With TaskDialogForDumpForm
        '初期化
        .Init

        'Excelをハンドラにする
        .ParenthWnd = Application.hWnd
    
        'ウィンドウタイトル
        .Title = "Mコードエクスポート"

        '見出し
        .MainInstruction = "PowerQueryのソースコードを出力します"

        '内容
        .Content = "現時点では、グループ出力には対応していません。" & vbCrLf & vbCrLf & "出力フォルダ："

        'TaskDialogで使えるコントロールを設定します
        .Flags = TDF_INPUT_BOX Or TDF_COMBO_BOX Or TDF_SHOW_PROGRESS_BAR Or TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_EXPAND_FOOTER_AREA

        'ボタンの設定(閉じる以外は、閉じれないようにする)
        .AddCustomButton ButtonAchievementID.出力, "出力": .SetButtonHold ButtonAchievementID.出力
        .AddCustomButton ButtonAchievementID.フォルダを選択, "フォルダを選択": .SetButtonHold ButtonAchievementID.フォルダを選択
        .AddCustomButton ButtonAchievementID.閉じる, "閉じる"

        '展開メッセージに関する設定
        .CollapsedControlText = "手順を表示"        'これから展開するメッセージ
        .ExpandedControlText = "手順を非表示"       'これから収納するメッセージ
        .ExpandedInfo = "上のコンボボックスでワークブックを選択、出力フォルダを入力し、OKボタンを押下してください"  '展開時に出すメッセージ

        'テキストボックスの設定
        .InputAlign = TDIBA_Content     'メッセージ部分に配置

        'コンボボックスを設定
        .ComboAlign = TDIBA_Footer      'フッターと同じ位置に配置
        .ComboStyle = cbtDropdownList   'ドロップダウンリストにする
        '　先程渡された引数から、リストを作成
        For i = 1 To UBound(BookNameList)
            .ComboAddItem CStr(BookNameList(i))
        Next
        '　初期選択位置を設定
        .ComboSetInitialItem 0

        'フッター部分に進捗用数値を出す用
        .Footer = "Ready..." & vbCrLf & vbCrLf

        '前回のパスを読み込む
        .InputText = Sh99_Setting.Range(RangeName_BeforePathName).Value


        '表示させます
        .ShowDialog

    End With
End Sub



'***************************************************************************************************
'                                   ■■■ 各種イベント ■■■
'***************************************************************************************************
'* 機能　　：ボタン押下時の各種イベントです。
'***************************************************************************************************
Private Sub TaskDialogForDumpForm_ButtonClick(ByVal ButtonID As Long)
    '各ボタン押下に応じて、値を変える
    Select Case ButtonID
        'フォルダ選択ダイアログを表示して、パスを入力します
        Case ButtonAchievementID.フォルダを選択
            'オブジェクト：フォルダ選択ダイアログ　を定義
            Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFolderPicker)
            With fd
                'ダイアログ設定
                .Title = "Mコードの出力先を選んでください"
                .AllowMultiSelect = False   '複数選択拒否
                
                '表示
                If .Show = -1 Then
                    TaskDialogForDumpForm.InputText = .SelectedItems(1)
                End If
            End With

        Case ButtonAchievementID.出力
            '入力チェック(新規フォルダ、既存フォルダはOKとする)
            Dim ResultCode As Long: ResultCode = BatchCreationFolder(TaskDialogForDumpForm.InputText)
            If ResultCode = 0 Or ResultCode = 183 Then
                '配列位置判定
                Dim TargetIndex As Long, i As Long
                If TaskDialogForDumpForm.ResultComboIndex = -1 Then
                    TargetIndex = 1
                Else
                    TargetIndex = TaskDialogForDumpForm.ResultComboIndex + 1
                End If
            
                'PowerQuery情報を取得
                Dim Infos_PowerQuery: Infos_PowerQuery = GetPowerQueryCode(OpeningBooksList(TargetIndex))
                
                '進捗更新
                TaskDialogForDumpForm.Footer = "0/" & UBound(Infos_PowerQuery)

                '情報がない場合(配列なし)はここで、終了
                If Not (IsArray(Infos_PowerQuery)) Then Exit Sub

                'ファイル出力
                Dim AddComment As String
                For i = 1 To UBound(Infos_PowerQuery)
                    'コメントがある場合は、それも加える
                    If Infos_PowerQuery(i, PowerQuery.Comment) <> "" Then
                        '改行コードを統一化
                        Infos_PowerQuery(i, PowerQuery.Comment) = Replace(Infos_PowerQuery(i, PowerQuery.Comment), vbLf, vbCrLf)
                    
                        'コメントのフォーマットに沿って、追加
                        AddComment = "//***************************************************************************************************" & vbCrLf & _
                                     "//" & vbTab & Replace(Infos_PowerQuery(i, PowerQuery.Comment), vbCrLf, vbCrLf & "//" & vbTab) & vbCrLf & _
                                     "//***************************************************************************************************" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
                    Else
                        'コメントなし
                        AddComment = ""
                    End If

                    '保存処理
                    SaveFile AddComment & Infos_PowerQuery(i, PowerQuery.M_Code), TaskDialogForDumpForm.ResultInput & "\" & Infos_PowerQuery(i, PowerQuery.QueryName)
                    
                    '進捗更新
                    TaskDialogForDumpForm.Footer = i & "/" & UBound(Infos_PowerQuery)
                    TaskDialog_UpdateProgressBar i, UBound(Infos_PowerQuery)
                Next
            
                'パスを記憶させる
                Sh99_Setting.Range(RangeName_BeforePathName).Value = TaskDialogForDumpForm.ResultInput

                '終了メッセージ
                MsgBox "Mコードのエクスポートを完了しました。" & vbCrLf & "OKを押下すると、エクスプローラーが開きます。", vbInformation, "エクスポート完了"
    
                ' エクスプローラーで開く
                Shell "explorer.exe """ & TaskDialogForDumpForm.ResultInput & """", vbNormalFocus
    
                '閉じる
                TaskDialogForDumpForm.CloseDialog
                
            Else
                MsgBox "パスが不正あるいは、空欄です。" & vbCrLf & "なお、ドライブ直下への保存はできません。", vbCritical, "ErrorCode：" & ResultCode
            End If
        Case Else
            'None
    
    End Select
End Sub



'***************************************************************************************************
'                                   ■■■ 各種情報収集 ■■■
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
'***************************************************************************************************
Private Function GetPowerQueryCode(ByVal targetBookName As String)
    '必要な変数を用意
    Dim wb As Workbook: Set wb = Workbooks(targetBookName)
    Dim queryCount As Long: queryCount = wb.Queries.Count
    Dim resultArray
    Dim i As Long
    Dim pq As WorkbookQuery

    '拡張子
    Const File拡張子 As String = ".pqm"

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
        resultArray(i, PowerQuery.QueryName) = pq.Name & File拡張子 'クエリ名
        resultArray(i, PowerQuery.Comment) = pq.Description         'コメント
        resultArray(i, PowerQuery.M_Code) = pq.Formula              'Mコード
        
        'カウントUP
        i = i + 1
    Next pq

    '返却
    GetPowerQueryCode = resultArray

End Function



'***************************************************************************************************
'                          ■■■ ファイル保存プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：指定した引数で、ファイル保存します。
'---------------------------------------------------------------------------------------------------
'* 引数　　：WriteText      書き込む内容を渡します。
'            SaveFilePass   入力した絶対パスにファイルを保存します。
'            OverWrite      上書きしたくない場合は、Falseで
'---------------------------------------------------------------------------------------------------
'* 注意事項：保存の文字コードは、「UTF-8(BOMなし)」のみです。
'***************************************************************************************************
Private Sub SaveFile(ByVal writeText As String, ByVal SaveFilePass As String, Optional OverWrite As Boolean = True)
    '上書きしない場合、後続処理しない
    If Not OverWrite And Dir(SaveFilePass, vbNormal) <> "" Then Exit Sub

    Dim tmp() As Byte 'BOM付きを外すための一時格納用
    With CreateObject("ADODB.Stream")
        '書き込み形式の設定
        .Charset = "UTF-8" 'UTF-8
        .Type = 2 'テキストモード
        .Open '上記の設定で、ストリームを開く

        'レスポンス結果を書き込む(末尾に改行コードあり)
        .writeText writeText, 1
        
        'BOM付きを外す処理
        .Position = 0 'ストリームの位置を0にセット
        .Type = 1 'データの種類をバイナリデータに変更
        .Position = 3 'ストリームの位置を3にセットして、BOMデータを飛ばす
        tmp = .Read 'ストリームの内容を一時格納用変数に保存。先程セットした始点3から最後まで
        .Close '一旦ストリームを閉じる（リセット）

        .Open 'ストリームを開く
        .Write tmp 'ストリームに一時格納したデータを流し込む
        .SaveToFile SaveFilePass, 2 'ファイルに上書き保存
        .Close
    End With

End Sub



'***************************************************************************************************
'                                   ■■■ ヘルパー機能 ■■■
'***************************************************************************************************
'* 機能　　：進捗バーの更新用に使います
'---------------------------------------------------------------------------------------------------
'* 引数　　：NowValue   現在の進捗値
'            MaxValue   最大値(ゲージMaxにする値)
'            StateMode  プログレスバーのステータス
'---------------------------------------------------------------------------------------------------
'* 機能説明：現在の値と、最大値、ステータスの3つの引数だけで、いい感じにプログレスバーを操作できます
'***************************************************************************************************
Private Sub TaskDialog_UpdateProgressBar(NowValue As Long, Optional MaxValue As Long = 100, Optional StateMode As eProgressBarStates = ePBST_NORMAL)
    With TaskDialogForDumpForm
        '進捗値が負なら、MARQUEEモードにします
        If NowValue < 0 Then
            If Not (BeforeMARQUEE) Then
                BeforeMARQUEE = True

                .ProgressSetType 1
                .ProgressStartMarquee
            End If
        Else
            '進捗のステータスを設定
            .ProgressSetState StateMode

            'MARQUEEモードをOFF
            If BeforeMARQUEE Then
                BeforeMARQUEE = False
            
                .ProgressSetType 0
                .ProgressStopMarquee
            End If
        
            '0~100までに収まるようにします
            If NowValue < MaxValue Then
                '99%までは、徐々に伸びるアニメーションを有効
                .ProgressSetValue Int(NowValue / MaxValue * 100)
            Else
                '最大値の場合は、徐々に伸びるアニメーションを無効にし、きっちりとゲージMaxにする
                '→https://dobon.net/vb/dotnet/control/pbdisableanimation.html
                
                '最大値を1つ増やしてから、元に戻す
                .ProgressSetRange 0, 101
                .ProgressSetValue 101
                .ProgressSetValue 100
                .ProgressSetRange 0, 100
            End If
        End If

    End With
End Sub
