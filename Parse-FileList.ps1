function ConvertTo-AbsolutePath {
    param (
        [string]$basePath,
        [string]$relativePath
    )

    # 基準パスの末尾にスラッシュがあれば削除
    $basePath = $basePath.TrimEnd('/')

    # 相対パスがスラッシュで始まる場合、それは既に絶対パスなので、そのまま返す
    if ($relativePath -match '^/') {
        return $relativePath
    }

    # パスを配列に分割
    $baseParts = $basePath -split '/'
    $relativeParts = $relativePath -split '/'

    # 結合されたパス部分を保持する配列を初期化
    $pathParts = $baseParts

    # 相対パスの各部分を処理
    foreach ($part in $relativeParts) {
        if ($part -eq '..') {
            # '..' が出現した場合、基準パスの最後の部分を削除
            if ($pathParts.Count -gt 0) {
                $pathParts = $pathParts[0..($pathParts.Count - 2)]
            }
        } elseif ($part -ne '.' -and $part -ne '') {
            # '.' や空でない部分をパスに追加
            $pathParts += $part
        }
    }

    # 基準パスと解決された相対パスを結合
    $absolutePath = $pathParts -join '/'

    return $absolutePath
}

function Parse-FileList {
    param (
        [string]$inputFile,
        [string]$outputFile
    )

    # 正規表現パターン（シンボリックリンク対応）
    $pattern = '\s*\d+\s+\d+\s+(\S+)\s+(\d+)\s+(\S+)\s+(\S+)\s+(\d+)\s+(\S+)\s+(\d+)\s+(\S+)\s+(.+?)(?:\s*->\s*(.+))?$'

    # CSVファイルにデータを書き込むための配列
    $csvData = @()

    # 入力ファイルを読み込み、正規表現で分割
    $lines = Get-Content $inputFile
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line -match $pattern) {
            $permissions = $matches[1]
            $links = $matches[2]
            $owner = $matches[3]
            $group_name = $matches[4]
            $size = $matches[5]
            $month = $matches[6]
            $day = $matches[7]
            $time_or_year = $matches[8]
            $name = $matches[9]
            $target = if ($matches[10]) { $matches[10] } else { "" }

            # ファイルタイプの判別
            $type = switch ($permissions[0]) {
                '-' { 'File' }
                'd' { 'Directory' }
                'l' { 'SymbolicLink' }
                'c' { 'CharacterDevice' }
                'b' { 'BlockDevice' }
                'p' { 'FIFO' }
                's' { 'Socket' }
                default { 'Unknown' }
            }

            # シンボリックリンクのターゲットが相対パスの場合、絶対パスに変換
            if ($target -and ($permissions[0] -eq 'l')) {
                $basePath = Split-Path -Path $name -Parent
                $absoluteTarget = ConvertTo-AbsolutePath -basePath $basePath -relativePath $target
                $target = $absoluteTarget
            }

            $csvData += [PSCustomObject]@{
                Type = $type
                Permissions = $permissions
                Links = $links
                Owner = $owner
                Group_Name = $group_name
                Size = $size
                Month = $month
                Day = $day
                Time_Or_Year = $time_or_year
                Name = $name
                Target = $target
                Line_Number = $i + 1
                Original_Line = $line
            }
        }
    }

    # タブ区切りでファイルに書き込む
    $csvData | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "`t"

    Write-Host "Data has been successfully parsed and saved to $outputFile"
}

# 関数の使用例
# Parse-FileList -inputFile "C:\path\to\ls_output.txt" -outputFile "C:\path\to\parsed_ls_output.tsv"
