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
    Get-Content $inputFile | ForEach-Object {
        if ($_ -match $pattern) {
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
            }
        }
    }

    # タブ区切りでファイルに書き込む
    $csvData | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "`t"

    Write-Host "Data has been successfully parsed and saved to $outputFile"
}

# 関数の使用例
# Parse-FileList -inputFile "C:\path\to\ls_output.txt" -outputFile "C:\path\to\parsed_ls_output.tsv"
