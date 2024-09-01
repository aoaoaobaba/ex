Attribute VB_Name = "modCheck"
Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' ショートパスが使えるか確認
Sub CheckShortPath()
    Dim longPath As String
    Dim shortPath As String
    Dim bufferSize As Long
    
    ' TODO: ★確認する長いパスを指定
    longPath = "\\ServerName\ShareFolder\DeepPath\長いパスのテスト.xlsx"
    
    ' バッファサイズを指定
    bufferSize = 260
    shortPath = String(bufferSize, vbNullChar)
    
    ' ショートパスを取得
    bufferSize = GetShortPathName(longPath, shortPath, bufferSize)
    
    If bufferSize > 0 Then
        shortPath = Left(shortPath, bufferSize)
        MsgBox "ショートパスが利用可能です: " & shortPath
    Else
        MsgBox "ショートパスを取得できませんでした。"
    End If
End Sub


