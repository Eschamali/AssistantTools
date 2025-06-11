Attribute VB_Name = "Mod05_DumpTreeMCode"
'***************************************************************************************************
'     クリップボードの仕様を利用して、ツリー情報を含むPowerQueryのソースコードを出力します。
'                 Mod01_DumpMCode と違い、フォルダ情報を含んだ出力となります。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'* 機能    ：PowerQueryの記録フォーマット、MashupFormat形式 をクリップボードから取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：MashupFormat形式 のxml文字列(UTF-8)
'***************************************************************************************************
Function クリップボードからMashupFormat形式のデータを抽出する() As String
    '取り出したいクリップボードのフォーマットを指定する
    Const クリップボードから抽出したいフォーマット名 As String = "Microsoft Mashup Format"
    Dim formatID As Long: formatID = RegisterClipboardFormat(クリップボードから抽出したいフォーマット名)
    
    'フォーマットチェック
    If formatID = 0 Then MsgBox "フォーマット登録に失敗しました", vbCritical, "クリップボードにはないフォーマット形式": Exit Function

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
