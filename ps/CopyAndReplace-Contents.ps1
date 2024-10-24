# ファイル名設定
$inputFile = "input.txt"
$outputFile1 = "output1.txt"
$outputFile2 = "output2.txt"

# 内容を読み込み、コメント行を除外
$content = Get-Content $inputFile | Where-Object { $_ -notmatch '^#' }

# 置換関数
function ReplaceContent($targetContent, $from, $to) {
    $targetContent -replace "^$from", $to
}

# 出力ファイル1に出力
ReplaceContent $content 'AAA' 'BBB' | Set-Content $outputFile1
ReplaceContent $content 'AAA' 'CCC' | Add-Content $outputFile1
ReplaceContent $content 'AAA' 'DDD' | Add-Content $outputFile1

# output1.txt の内容を output2.txt にコピー
Get-Content $outputFile1 | Set-Content $outputFile2

Write-Output "処理が完了しました。ファイルは $outputFile1 と $outputFile2 に保存されました。"
