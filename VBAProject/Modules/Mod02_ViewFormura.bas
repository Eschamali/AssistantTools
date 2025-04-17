Attribute VB_Name = "Mod02_ViewFormura"
'***************************************************************************************************
'                 アクティブセルにある計算式を、コメントで表示させる機能です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ アドインから呼び出すプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：アクティブセルにある計算式を、コメントで表示させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：・既にコメントがある場合は、削除します。
'            ・絵文字等の「Unicode サロゲート ペア」を含む数式は、サポートしません
'            ・BYROW関数等で使える「イータ縮小ラムダ関数」の色付けには非対応です
'***************************************************************************************************
Sub FormuraIntoMemo()
Attribute FormuraIntoMemo.VB_Description = "数式のあるセルをプレビューします。"
Attribute FormuraIntoMemo.VB_ProcData.VB_Invoke_Func = "m\n14"
    'コメントが既にある場合は、一旦削除だけします。
    If Not ActiveCell.Comment Is Nothing Then ActiveCell.Comment.Delete: Exit Sub


    '-------------------------------------------------0.設定値-------------------------------------------------
'    Call StartOrLapTimer("数式をコメント表示")


    '記述するフォントサイズを指定
    Const MemoFontSize = 22

    '色設定(必要に応じて、追加してください)
    Const Color_FunctionName As Long = vbBlue   '関数名：青
    Const Color_String As Long = rgbBrown       '文字列：茶
    Const Color_Comma As Long = vbRed           'カンマ：赤
    Const Color_Ampersand As Long = vbMagenta   '＆記号：マゼンダ
    Const Color_NumberSign As Long = vbCyan     '＃記号：シアン


    '-------------------------------------------------1.変数用意-------------------------------------------------


    '調査に必要な変数
    Dim A_Formula As String             'セル内の文字列格納用
    Dim charNum As Integer              'セルに含むことができる合計文字数(32767)に準拠してあえて、Integer
    Dim CountDoubleQuotation As Integer '今の文字位置が、文字列かそうじゃないか判定

    '判定用(前述で用意した「色設定」分)
    Dim Flag_FunctionName() As Boolean  '1文字1文字が、関数名に当てはまるかフラグ
    Dim Flag_String() As Boolean        '1文字1文字が、文字列に当てはまるかフラグ
    Dim Flag_Comma() As Boolean         '1文字1文字が、カンマに当てはまるかフラグ
    Dim Flag_Ampersand() As Boolean     '1文字1文字が、＆記号に当てはまるかフラグ
    Dim Flag_NumberSign() As Boolean    '1文字1文字が、＃記号に当てはまるかフラグ


    '-------------------------------------------------2.配列準備-------------------------------------------------


    With ActiveCell
        '計算式を格納
        A_Formula = .Formula2
        
        '数式判定(1文字目は必ず、「=」)
        If Not Left(A_Formula, 1) = "=" Then
            MsgBox "選択したセルは数式ではありません。", vbExclamation, "数式のないセル"
            Exit Sub
        End If

        '文字数を取得し、それを配列数として定義する(前述で用意した「色設定」分)
        charNum = Len(A_Formula)
        ReDim Flag_FunctionName(1 To charNum)
        ReDim Flag_String(1 To charNum)
        ReDim Flag_Comma(1 To charNum)
        ReDim Flag_Ampersand(1 To charNum)
        ReDim Flag_NumberSign(1 To charNum)


        '-------------------------------------------------3.装飾箇所調査処理-------------------------------------------------
'        Call StartOrLapTimer("準備完了")


        '文字数分、所定の処理を行う
        Dim i As Integer, j As Integer
        Dim nowChar As String
        For i = 1 To charNum
            nowChar = Mid(A_Formula, i, 1)
            
            '文字列ゾーン切り替えcheck
            If nowChar = """" Then
                '各文字にて、ダブルクォーテーションがあったらカウントします
                CountDoubleQuotation = CountDoubleQuotation + 1

                'ダブルクォーテーション自体は文字列としておく
                Flag_String(i) = True
            
            'ダブルクォーテーションの出現数が偶数時、所定の処理を行う
            ElseIf CountDoubleQuotation Mod 2 = 0 Then
                '数式エリアで、特定の記号を検知したら、所定の処理を行う
                Select Case nowChar
                    '-------------------------------------------------3-1.関数名、調査処理-------------------------------------------------
                    Case "("
                        '関数名の最終文字位置から、1文字ずつ戻して始点(関数名の最初文字位置)を探す
                        For j = i - 1 To 1 Step -1
                            '1文字ずつ、抜き出す
                            nowChar = Mid(A_Formula, j, 1)
                            
                            '始点位置文字列による検出法を用います
                            Select Case nowChar
                                '始点位置が検出されたら、調査終了とする
                                '※ここに始点位置であろうパターンをカンマ区切りで1文字ずつ、列挙してください…
                                Case "(", ",", "=", vbLf, " ", "+", "-", "*", "/", "^"
                                    Exit For

                                Case Else
                                    '関数名なので、フラグ付けする
                                    Flag_FunctionName(j) = True


'                            '関数名であるパターン目印での検出法を用います
'                                Case "A" To "Z", "a" To "z", ".", "_"
'                                    '関数名なので、フラグ付けする
'                                    Flag_FunctionName(j) = True
'
'                                Case Else
'                                    '大文字小文字アルファベット、ピリオド、アンダーバー以外の文字列が検出されたら、調査終了
'                                    Exit For

                            End Select
                        Next
                         
                    '-------------------------------------------------3-2.単一文字、調査処理-------------------------------------------------
                    '「,」があったら、フラグ付けする
                    Case ","
                        Flag_Comma(i) = True
                        
                    '「&」があったら、フラグ付けする
                    Case "&"
                        Flag_Ampersand(i) = True

                    '「#」があったら、フラグ付けする
                    Case "#"
                        Flag_NumberSign(i) = True

                    Case Else
                        '何もしない

                End Select
                
            '奇数なら、文字列エリアとする。
            Else
                '文字列フラグ
                Flag_String(i) = True
            End If
        
        Next


        '-------------------------------------------------4.メモ準備-------------------------------------------------
'        Call StartOrLapTimer("文字の調査完了")


        '色の変化による画面更新を無効化する
        Application.ScreenUpdating = False
        
        '現在のアクティブセルにメモを挿入
        .AddComment
        
        With .Comment
            '常時表示
            .Visible = True
        
            'コメント記述
            .Text Text:=A_Formula
            
            'フォントサイズ変更
            .Shape.TextFrame.Characters.Font.SIZE = MemoFontSize
          
            '幅と高さを記述文字に合わせて変更
            .Shape.TextFrame.AutoSize = True


            '-------------------------------------------------5.3で調査したフラグを元に色付け-------------------------------------------------
'            Call StartOrLapTimer("メモの挿入完了")


            Dim StringCount As Integer  '文字数カウント用
            For i = 1 To charNum
                '初期化
                StringCount = 0
            
                '用意したフラグ分、if else文を用意
                '　関数名
                If Flag_FunctionName(i) Then
                    '複数文字数を考慮して、文字数をカウント
                    For j = i To charNum
                        'Falseが出るまでカウント
                        If Flag_FunctionName(j) Then StringCount = StringCount + 1 Else Exit For
                    Next

                    '色付け
                    .Shape.TextFrame.Characters(Start:=i, Length:=StringCount).Font.Color = Color_FunctionName
                    
                    'カウントした分、ずらす
                    i = i + StringCount


                '　文字列
                ElseIf Flag_String(i) Then
                    '複数文字数を考慮して、文字数をカウント
                    For j = i To charNum
                        'Falseが出るまでカウント
                        If Flag_String(j) Then StringCount = StringCount + 1 Else Exit For
                    Next

                    '色付け
                    .Shape.TextFrame.Characters(Start:=i, Length:=StringCount).Font.Color = Color_String

                    'カウントした分、ずらす
                    i = i - 1 + StringCount


                '　カンマ
                ElseIf Flag_Comma(i) Then
                    '1文字なので、そのまま色付けへ
                    .Shape.TextFrame.Characters(Start:=i, Length:=1).Font.Color = Color_Comma


                '　＆記号
                ElseIf Flag_Ampersand(i) Then
                    '1文字なので、そのまま色付けへ
                    .Shape.TextFrame.Characters(Start:=i, Length:=1).Font.Color = Color_Ampersand


                '　＃記号
                ElseIf Flag_NumberSign(i) Then
                    '1文字なので、そのまま色付けへ
                    .Shape.TextFrame.Characters(Start:=i, Length:=1).Font.Color = Color_NumberSign
                
                End If
            Next
            
        End With
    
    End With

'    Call StartOrLapTimer("メモの色付け完了")

    '画面更新をON
    Application.ScreenUpdating = True
'    Call EndTimer("画面更新して、処理終了")
End Sub
