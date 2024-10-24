@echo off

::Here-String
powershell -Command @"
    $text = 'Hello world, Hello universe!'
    $text = $text -replace '^Hello', 'Hi'
    Write-Output $text
"@
