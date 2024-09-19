# 指定されたディレクトリ内のすべての .jar ファイルから特定のテキストファイルを抽出して表示
$jarDir = "C:\path\to\your\jar\files"  # .jar ファイルが置かれているディレクトリ
$targetExtension = ".txt"  # 取得したいファイルの拡張子（例: .txt）

# System.IO.Compression.FileSystem アセンブリを読み込む（PowerShell 5.0以降では標準で利用可能）
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# .jar ファイルごとに処理
Get-ChildItem -Path $jarDir -Filter *.jar | ForEach-Object {
    $jarFile = $_.FullName

    # .jar (ZIP) ファイルを開いて内容を確認
    [System.IO.Compression.ZipFile]::OpenRead($jarFile).Entries | ForEach-Object {
        $entry = $_
        
        # ファイルがテキストファイルであるかどうかを確認
        if ($entry.FullName -like "*$targetExtension") {
            Write-Host "Extracting from: $jarFile -> $entry.FullName"
            
            # ファイル内容を文字列として読み取る
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $content = $reader.ReadToEnd()
            $reader.Close()
            
            # テキスト内容を出力
            Write-Host $content
        }
    }
}
