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

function IsExcluded {
    param (
        [string]$name,
        [string]$target
    )

    # yymmdd, yyyymmddの正規表現パターン
    $datePattern = '\d{2}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])|(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])'

    if ($name -match $datePattern) {
        return "name: $($matches[0])"
    } elseif ($target -match $datePattern) {
        return "target: $($matches[0])"
    } else {
        return $null
    }
}
function Parse-FileList {
    param (
        [string]$inputFile,
        [string]$outputFile,
        [string]$excludedFile
    )

    # 正規表現パターン（シンボリックリンク対応）
	$pattern = '\s*\d+\s+\d+\s+(\S+)\s+(\d+)\s+(\S+)\s+(\S+)\s+(\d+)\s+(\S+\s+\d+\s+\S+)\s+(.+?)(?:\s*->\s*(.+))?$'

    # CSVファイルにデータを書き込むための配列
    $csvData = @()
    $excludedData = @()

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
            $date = $matches[6]
            $name = $matches[7]
            $target = if ($matches[8]) { $matches[8] } else { "" }

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

            # ディレクトリとファイル名を分ける
            if ($type -eq 'Directory') {
                $directory = $name
                $fileName = ""
            } else {
                $directory = Split-Path -Path $name -Parent
                $fileName = Split-Path -Path $name -Leaf
            }

            # シンボリックリンクのターゲットが相対パスの場合、絶対パスに変換
            if ($target -and ($permissions[0] -eq 'l')) {
                $basePath = Split-Path -Path $name -Parent
                $absoluteTarget = ConvertTo-AbsolutePath -basePath $basePath -relativePath $target
                $target = $absoluteTarget
            }

            # 除外条件のチェック
            $excludeReason = IsExcluded -name $name -target $target
            if ($excludeReason) {
                $excludedData += [PSCustomObject]@{
                    Type = $type
                    Permissions = $permissions
                    Links = $links
                    Owner = $owner
                    Group_Name = $group_name
                    Size = $size
                    Date = $date
                    Directory = $directory
                    FileName = $fileName
                    Target = $target
                    Line_Number = $i + 1
                    Original_Line = $line
                    ExcludeReason = $excludeReason
                }
            } else {
                $csvData += [PSCustomObject]@{
                    Type = $type
                    Permissions = $permissions
                    Links = $links
                    Owner = $owner
                    Group_Name = $group_name
                    Size = $size
                    Date = $date
                    Directory = $directory
                    FileName = $fileName
                    Target = $target
                    Line_Number = $i + 1
                    Original_Line = $line
                }
            }
        }
    }

    # タブ区切りでファイルに書き込む
    $csvData      | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "`t"
    $excludedData | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "`t" -Append
    $excludedData | Export-Csv -Path $excludedFile -NoTypeInformation -Delimiter "`t"

    Write-Host "Data has been successfully parsed and saved to $outputFile and $excludedFile"
}

# 関数の使用例
# Parse-LSOutput -inputFile "C:\path\to\ls_output.txt" -outputFile "C:\path\to\parsed_ls_output.tsv" -excludedFile "C:\path\to\excluded_output.tsv"
