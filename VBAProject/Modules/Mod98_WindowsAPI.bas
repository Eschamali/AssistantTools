Attribute VB_Name = "Mod98_WindowsAPI"
'***************************************************************************************************
'                                   WindowsAPI宣言モジュール
'---------------------------------------------------------------------------------------------------
'           ただし、複雑な処理はここには記述しません。基本的に宣言と簡易仕様関数とすること
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ SHCreateDirectoryEx API宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "Shell32" _
    Alias "SHCreateDirectoryExA" _
    (ByVal hWnd As LongPtr, _
     ByVal pszPath As String, _
     ByVal psa As LongPtr) As Long



'***************************************************************************************************
'                           ■■■ クリップボードにアクセスする API宣言 ■■■
'***************************************************************************************************
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long



'***************************************************************************************************
'           ■■■ 代替の CopyMemory：ループではなくバイト配列に直接書き込み API宣言 ■■■
'***************************************************************************************************
Public Declare PtrSafe Sub RtlMoveMemoryArray Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByVal Source As LongPtr, _
    ByVal Length As Long)



'***************************************************************************************************
'                               ■■■ 文字コード一式宣言 ■■■
'***************************************************************************************************
Public CharConv As New CharacterCodeConversion



'***************************************************************************************************
'                       ■■■ WindowsAPIを使えるように、ヘルパー実装 ■■■
'***************************************************************************************************
'* 機能：SHCreateDirectoryEx　を利用して多階層フォルダを一気に作成します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：0      正常にフォルダが作成された場合
'            161    パスの記述が不正
'            183    フォルダが既に存在している場合
'
'* 引数　　：fullPassDir      作成するフォルダまでのフルパス
'---------------------------------------------------------------------------------------------------
'注意事項：その他のエラーコードは、https://learn.microsoft.com/ja-jp/windows/win32/debug/system-error-codes--0-499-　で
'          とはいえ、上記2つ以外は、基本エラーと考えてよいです。
'***************************************************************************************************
Function BatchCreationFolder(fullPassDir As String) As Long
    
    BatchCreationFolder = SHCreateDirectoryEx(0, fullPassDir, 0)

End Function
