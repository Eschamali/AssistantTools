Attribute VB_Name = "Mod05_DumpTreeMCode"
'***************************************************************************************************
'     クリップボードの仕様を利用して、ツリー情報を含むPowerQueryのソースコードを出力します。
'                 Mod01_DumpMCode と違い、フォルダ情報を含んだ出力となります。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ グローバル定義 ■■■
'***************************************************************************************************
Dim クエリ名に対するグループパス情報    As Scripting.Dictionary
Dim グループパスに対する説明文          As Scripting.Dictionary



'***************************************************************************************************
'         ■■■ クリップボードから、"Microsoft Mashup Format"形式にある内容を取得 ■■■
'***************************************************************************************************
'* 機能    ：PowerQueryの記録フォーマット、MashupFormat形式 をクリップボードから取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：MashupFormat形式 のxml文字列(UTF-8)
'---------------------------------------------------------------------------------------------------
'* 詳細情報：PowerQueryの定義データを直接取得することが出来ないため、クリップボードを経由しての取得となります。
'            「クエリと接続」パネルを開いて、選択→コピー→このプロシージャを実行　で、得れます。
'            Shiftキーを押しながら、上部と下部を選択することで、実質全部の構造を取得できます。
'            現状、クリップボードを経由しないと、ツリー情報の取得は出来なさそうです…
'
'* 注意事項：・クリップボードに、"Microsoft Mashup Format"がないと、vbnullstring が返ります。
'          ：・当然ですが事前に、WindowsAPI宣言が必要です。
'***************************************************************************************************
Private Function クリップボードからMashupFormat形式のデータを抽出する() As String
    '取り出したいクリップボードのフォーマットを指定する
    Const クリップボードから抽出したいフォーマット名 As String = "Microsoft Mashup Format"
    Dim formatID As Long: formatID = RegisterClipboardFormat(クリップボードから抽出したいフォーマット名)
    
    'フォーマットチェック
    If formatID = 0 Then MsgBox クリップボードから抽出したいフォーマット名 & " のフォーマット登録に失敗しました。", vbCritical, "クリップボードでは扱えないフォーマット形式": Exit Function

    'クリップボードOpenチェック
    If OpenClipboard(0) = 0 Then MsgBox "別のアプリケーションでロックされてる可能性があります。" & vbCrLf & "心当たりのあるアプリ側で対処して下さい。", vbCritical, "クリップボードを開けません": Exit Function

    '指定したフォーマット形式が、クリップボードにあるかチェック
    Dim hMem As LongPtr: hMem = GetClipboardData(formatID)
    If hMem = 0 Then
        MsgBox クリップボードから抽出したいフォーマット名 & " 形式のクリップボードが検出できませんでした。" & vbCrLf & vbCrLf & "別のコンテンツをコピーした可能性があります。再度、クエリと接続からコピー操作を行って下さい。", vbExclamation, クリップボードから抽出したいフォーマット名 & " が見つかりません"
        CloseClipboard
        Exit Function
    End If

    'クリップボードのデータのポインタを得て、クリップボードをロックする
    Dim lpMem As LongPtr: lpMem = GlobalLock(hMem)
    
    '格納サイズを取得
    Dim dataSize As Long: dataSize = GlobalSize(hMem)

    'ロックまたはサイズ取得失敗判定
    If lpMem = 0 Or dataSize = 0 Then
        MsgBox "別のアプリケーションでロックされてる可能性があります。" & vbCrLf & "心当たりのあるアプリ側で対処して下さい。", vbCritical, "クリップボードのサイズ検出に失敗": Exit Function
        CloseClipboard
        Exit Function
    End If

    ' バイナリ全体を一括コピー
    Dim byteData() As Byte: ReDim byteData(dataSize - 1)
    Call RtlMoveMemoryArray(byteData(0), lpMem, dataSize)

    '後始末
    Call GlobalUnlock(hMem) 'メモリのロック解除
    Call CloseClipboard     'クリップボードを閉じる


    'バイト配列から、テキストを取得
    Dim rawStr As String: rawStr = BytesToString(byteData)

    ' バイナリの先頭にゴミが含まれている可能性があるため <?xml から探す
    Dim xmlStart As Long: xmlStart = InStr(rawStr, "<?xml")
    If xmlStart > 0 Then
        クリップボードからMashupFormat形式のデータを抽出する = Mid$(rawStr, xmlStart, Len(rawStr) - xmlStart)
    Else
        クリップボードからMashupFormat形式のデータを抽出する = rawStr ' xml始まりがないときは一旦、全部出す
    End If
End Function

'***************************************************************************************************
'* 機能　　：Byte() → String に変換するヘルパー関数です
'---------------------------------------------------------------------------------------------------
'* 引数　　：bytes()    バイト配列
'            encoding   エンコード名称(デフォルト：UTF-8)
'***************************************************************************************************
Private Function BytesToString(bytes() As Byte, Optional encoding As String = "UTF-8") As String
    With CreateObject("ADODB.Stream")
        '読み込みモードを設定
        .Type = 1 'バイナリ
        
        '開く
        .Open
        
        'Streamへ書き込む
        .Write bytes
        
        '先頭スタートにする
        .Position = 0
        
        'モードを変更
        .Type = 2 'テキスト
        
        '文字コード名を設定
        .Charset = encoding
        
        'エンコードした内容を返す
        BytesToString = .ReadText
        
        '後始末
        .Close
    End With
End Function



'***************************************************************************************************
'         ■■■ "Microsoft Mashup Format"形式のPowerQueryの定義データを解析 ■■■
'***************************************************************************************************
'* 機能    ：事前に抽出した「PowerQueryの定義XMLデータ」を基に、下記2種類の「Scripting.Dictionary」を作成します
'               ・クエリ名に対するグループパス情報
'               ・グループパスに対する説明文
'---------------------------------------------------------------------------------------------------
'* 引数    ：XML_PowerQuery     事前に抽出した「PowerQueryの定義XMLデータ」
'* 返り値  ：バージョン情報     Client,Version,MinVersion,Culture,SafeCombine
'---------------------------------------------------------------------------------------------------
'* 注意事項：・2種類の返り値を返す都合上、グローバル変数による格納を行います
'            ・引数が不正の場合、vbnullstring が返ります
'            ・下記の参照設定が必要です
'               - Microsoft XML v6.0
'               - Microsoft Scripting Runtime
'***************************************************************************************************
Function ParseMashupQueries(XML_PowerQuery As String) As String
    '空文字引数なら、ここで終わり
    If XML_PowerQuery = "" Then Exit Function


    '------------------------------------------初期化------------------------------------------
    Set クエリ名に対するグループパス情報 = New Scripting.Dictionary
    Set グループパスに対する説明文 = New Scripting.Dictionary


    '------------------------------------------XMLを読み込む準備------------------------------------------
    Dim xmlDoc As New MSXML2.DOMDocument60
    xmlDoc.async = False                                                                            '読み込みが終わるまで待機する
    xmlDoc.validateOnParse = False                                                                  'DTD（Document Type Definition）による妥当性チェックをしない。「<!DOCTYPE 〇〇」というのがないので OFF
    xmlDoc.LoadXML XML_PowerQuery                                                                   '読み込む
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/DataMashup'"   '名前空間の指定
    
    
    '------------------------ 1.グループパスに対するDescriptionを網羅的に登録 ------------------------
    '欲しい情報名を指定
    Const クエリグループ名_要素名       As String = "QueryGroup"
    Const クエリグループ名_属性名       As String = "Name"
    Const クエリグループ名の説明_要素名 As String = "Description"

    '必要な変数を用意
    Dim groupNodes As MSXML2.IXMLDOMNodeList: Set groupNodes = xmlDoc.SelectNodes("//d:" & クエリグループ名_要素名) '"QueryGroup"という要素名を一覧化
    Dim parentNode As MSXML2.IXMLDOMNode: Dim groupNameAttr As MSXML2.IXMLDOMAttribute
    Dim groupPath As String, selfDescText As String

    '探索開始
    Dim groupNode As MSXML2.IXMLDOMNode
    For Each groupNode In groupNodes
        '現在のノード情報を取得
        Set parentNode = groupNode.parentNode
        
        'このグループ自身を含めてパスを作る準備
        groupPath = "/" & groupNode.Attributes.getNamedItem(クエリグループ名_属性名).Text
    
        '先頭の要素まで、遡って探索します
        Do While Not parentNode Is Nothing
            '"QueryGroup"要素ゾーンに入ったら、パスを連結させます
            If parentNode.nodeName = クエリグループ名_要素名 Then groupPath = "/" & parentNode.Attributes.getNamedItem(クエリグループ名_属性名).Text & groupPath

            '現在位置のノードを登録
            Set parentNode = parentNode.parentNode
        Loop

        '現在位置の"Description"を取得
        selfDescText = groupNode.SelectSingleNode("d:" & クエリグループ名の説明_要素名).Text


        ' 登録
        グループパスに対する説明文.Add groupPath, selfDescText
        Debug.Print "グループパスに対する説明文：" & groupPath & " → " & selfDescText
    Next

    
    
    '---------------------------------------2.クエリ名に対する"QueryGroup"パスを網羅的に登録 ---------------------------------------
    '欲しい情報名を指定
    Const クエリ名_要素名       As String = "Query"
    Const クエリ名_属性名       As String = "Name"

    Dim queryNodes As MSXML2.IXMLDOMNodeList: Set queryNodes = xmlDoc.SelectNodes("//d:" & クエリ名_要素名)      '"Query"という要素名を一覧化
    Dim queryNode As MSXML2.IXMLDOMNode, queryName As String
    For Each queryNode In queryNodes
        '現在位置の、"Name"属性値を取得
        queryName = queryNode.Attributes.getNamedItem(クエリ名_属性名).Text
        
        'グループパス取得準備
        'グループパス（上位のQueryGroup）をたどる
        Set parentNode = queryNode.parentNode
        
        groupPath = ""
        
        Do While Not parentNode Is Nothing
            If parentNode.nodeName = クエリグループ名_要素名 Then
                Set groupNameAttr = parentNode.Attributes.getNamedItem("Name")
                If Not groupNameAttr Is Nothing Then
                    groupPath = "/" & groupNameAttr.Text & groupPath
                End If
            End If
            Set parentNode = parentNode.parentNode
        Loop
        
        クエリ名に対するグループパス情報.Add queryName, groupPath
        Debug.Print "クエリ名に対するグループパス情報：" & queryName & " → " & groupPath
    Next



    '後でバージョン情報やメタ情報を、カンマ区切りで、返す予定
    ParseMashupQueries = "EXCEL,2.145.7712.4,2.21.0.0,ja-JP,false"

End Function



'***************************************************************************************************
'* 機能　　：指定した引数で、ファイル保存します。
'---------------------------------------------------------------------------------------------------
'* 引数　　：WriteText      書き込む内容を渡します。
'            SaveFilePass   入力した絶対パスにファイルを保存します。
'            OverWrite      上書きしたくない場合は、Falseで
'***************************************************************************************************
Private Sub SaveFile(writeText As String, ByVal SaveFilePass As String, Optional OverWrite As Boolean = True)
    '上書きしない場合、後続処理しない
    If Not OverWrite And Dir(SaveFilePass, vbNormal) <> "" Then Exit Sub

    Dim tmp() As Byte 'BOM付きを外すための一時格納用
    With CreateObject("ADODB.Stream")
        '書き込み形式の設定
        .Charset = "UTF-8" 'UTF-8
        .Type = 2 'テキストモード
        .Open '上記の設定で、ストリームを開く

        'レスポンス結果を書き込む(末尾に改行コードなし)
        .writeText writeText, 0
        
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
