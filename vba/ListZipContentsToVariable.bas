Sub ListZipContentsToVariable()
    Dim zipFilePath As String
    Dim shellCommand As String
    Dim wsh As Object
    Dim execObj As Object
    Dim outputLine As String
    Dim zipContents As String

    ' zipファイルのパス
    zipFilePath = "C:\Path\To\archive.zip"
    
    ' 7zipコマンド
    shellCommand = "C:\Path\To\7z.exe l """ & zipFilePath & """"
    
    ' WScript.Shellオブジェクトを作成
    Set wsh = CreateObject("WScript.Shell")
    
    ' コマンドを実行し、出力を取得
    Set execObj = wsh.Exec(shellCommand)
    
    ' コマンドの出力を読み取る
    Do While Not execObj.StdOut.AtEndOfStream
        outputLine = execObj.StdOut.ReadLine
        zipContents = zipContents & outputLine & vbCrLf
    Loop
    
    ' 出力結果をメッセージボックスで表示
    MsgBox zipContents
    
    ' 後で使うために、変数に格納された内容を利用することも可能
    ' 例: デバッグプリント
    Debug.Print zipContents
End Sub
