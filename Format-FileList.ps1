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
        }
        elseif ($part -ne '.' -and $part -ne '') {
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

    # yymmdd or yyyymmdd の正規表現パターン
    $datePattern = '\d{2}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])|(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])'

    if ($name -match $datePattern) {
        return "name: $($matches[0])"
    }
    elseif ($target -match $datePattern) {
        return "target: $($matches[0])"
    }
    else {
        return $null
    }
}

function Format-LSOutput {
    param (
        [string]$inputFile,
        [string]$outputFile,
        [string]$excludedFile,
        [string]$errorFile
    )

    # 正規表現パターン
    $pattern = '\s*\d+\s+\d+\s+(\S+)\s+(\d+)\s+(\S+)\s+(\S+)\s+(\d+)\s+(\S+\s+\d+\s+\S+)\s+(.+?)(?:\s*->\s*(.+))?$'

    # ヘッダー行の定義（大文字）
    $csvHeader = "TYPE`tPERMISSIONS`tLINKS`tOWNER`tGROUP_NAME`tSIZE`tDATE`tDIRECTORY`tFILE_NAME`tTARGET`tLINE_NUMBER`tORIGINAL_LINE"
    $excludedHeader = "TYPE`tPERMISSIONS`tLINKS`tOWNER`tGROUP_NAME`tSIZE`tDATE`tDIRECTORY`tFILE_NAME`tTARGET`tLINE_NUMBER`tORIGINAL_LINE`tEXCLUDE_REASON"
    $errorHeader = "LINE_NUMBER`tORIGINAL_LINE"

    # ヘッダー行を書き込む
    Out-File -FilePath $outputFile -InputObject $csvHeader -Encoding UTF8
    Out-File -FilePath $excludedFile -InputObject $excludedHeader -Encoding UTF8
    Out-File -FilePath $errorFile -InputObject $errorHeader -Encoding UTF8

    $i = 0
    $reader = [System.IO.File]::OpenText($inputFile)
    $outputWriter = [System.IO.StreamWriter]::new($outputFile, $true, [System.Text.Encoding]::UTF8)
    $excludedWriter = [System.IO.StreamWriter]::new($excludedFile, $true, [System.Text.Encoding]::UTF8)
    $errorWriter = [System.IO.StreamWriter]::new($errorFile, $true, [System.Text.Encoding]::UTF8)

    try {
        while ($null -ne ($line = $reader.ReadLine())) {
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
                    'l' { 'Symbolic Link' }
                    'c' { 'Character Device' }
                    'b' { 'Block Device' }
                    'p' { 'FIFO' }
                    's' { 'Socket' }
                    default { 'Unknown' }
                }

                # ディレクトリとファイル名を分ける
                if ($type -eq 'Directory') {
                    $directory = $name
                    $fileName = ""
                }
                else {
                    $directory = Split-Path -Path $name -Parent
                    $fileName = Split-Path -Path $name -Leaf
                }

                # シンボリックリンクのターゲットが相対パスの場合、絶対パスに変換
                if ($target -and ($permissions[0] -eq 'l')) {
                    $basePath = $directory
                    $absoluteTarget = ConvertTo-AbsolutePath -basePath $basePath -relativePath $target
                    $target = $absoluteTarget
                }

                # 出力データ作成
                $dataObject = [PSCustomObject]@{
                    Type          = $type
                    Permissions   = $permissions
                    Links         = $links
                    Owner         = $owner
                    Group_Name    = $group_name
                    Size          = $size
                    Date          = $date
                    Directory     = $directory
                    FileName      = $fileName
                    Target        = $target
                    Line_Number   = $i + 1
                    Original_Line = $line
                }

                # CSV形式に変換して出力
                $dataObject | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" -UseCulture | Select-Object -Skip 1 | ForEach-Object {
                    $outputWriter.WriteLine($_)
                }

                # 除外データを $excludedFile に出力
                $excludeReason = IsExcluded -name $name -target $target
                if ($excludeReason) {
                    # 除外理由を追加
                    $dataObject | Add-Member -MemberType NoteProperty -Name ExcludeReason -Value $excludeReason
                    # CSV形式に変換して出力
                    $dataObject | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" -UseCulture | Select-Object -Skip 1 | ForEach-Object {
                        $excludedWriter.WriteLine($_)
                    }
                }
            }
            else {
                $errorObject = [PSCustomObject]@{
                    Line_Number   = $i + 1
                    Original_Line = $line
                }
                $errorObject | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" -UseCulture | Select-Object -Skip 1 | ForEach-Object {
                    $errorWriter.WriteLine($_)
                }
            }
            $i++
        }
    }
    catch {
        # エラーが発生した場合の処理
        Write-Host "An error occurred:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    finally {
        $reader.Close()
        $outputWriter.Close()
        $excludedWriter.Close()
        $errorWriter.Close()
    }

    Write-Host "Data has been successfully parsed and saved to $outputFile, $excludedFile, and $errorFile"
}

# 関数の使用例
# Parse-LSOutput `
#     -inputFile "C:\path\to\ls_output.txt" `
#     -outputFile "C:\path\to\parsed_ls_output.tsv" `
#     -excludedFile "C:\path\to\excluded_output.tsv" `
#     -errorFile "C:\path\to\error_output.tsv"
