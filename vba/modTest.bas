Attribute VB_Name = "modTest"
Option Explicit

' テスト結果表示用のサブプロシージャ
Private Sub printTestResult(testName As String, condition As Boolean)
    If condition Then
        Debug.Print testName & ": 成功"
    Else
        Debug.Print testName & ": 失敗"
    End If
End Sub

' 1. フォルダ作成 - 短いパス
Sub Test_CreateFolder_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderShort"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_ShortPath", fileHandler.folderExists(testFolderPath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 2. フォルダ作成 - 長いパス
Sub Test_CreateFolder_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    ' ここで長いパスを作成（260桁超）
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\FolderK\FolderL\FolderM\FolderN\FolderO\FolderP\FolderQ\FolderR\FolderS\FolderT\FolderU\FolderV\FolderW\FolderX\FolderY\FolderZ"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_LongPath", fileHandler.folderExists(testFolderPath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 3. フォルダ作成 - 既存のフォルダ
Sub Test_CreateFolder_Existing()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderExisting"
    
    fileHandler.createFolders testFolderPath
    On Error Resume Next
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_Existing", Err.Number = 0
    
    ' 後始末
    On Error GoTo 0
    fileHandler.deleteFolder testFolderPath
End Sub

' 4. フォルダ作成 - 不正なパス
Sub Test_CreateFolder_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = "?:\InvalidPath"
    
    On Error Resume Next
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_InvalidPath", Err.Number = 100 And InStr(Err.Description, "フォルダの作成に失敗しました。") > 0
    On Error GoTo 0
End Sub

' 5. ファイル存在チェック - 短いパス
Sub Test_FileExists_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\TestFileShort.txt"
    
    ' テストファイル作成
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    printTestResult "Test_FileExists_ShortPath", fileHandler.fileExists(testFilePath)
    
    ' 後始末
    Kill testFilePath
End Sub

' 6. ファイル存在チェック - 長いパス
Sub Test_FileExists_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    ' 長いパスを260桁超に設定
    testFolderPath = ThisWorkbook.path & "\LongPathFolder\SubFolder1\SubFolder2\SubFolder3\SubFolder4\SubFolder5\SubFolder6\SubFolder7\SubFolder8\SubFolder9\SubFolder10\SubFolder11\SubFolder12\SubFolder13\SubFolder14\SubFolder15\SubFolder16\SubFolder17\SubFolder18\SubFolder19\SubFolder20\SubFolder21\SubFolder22\SubFolder23\SubFolder24\SubFolder25"
    testFilePath = testFolderPath & "\TestFileLong.txt"
    
    ' フォルダとテストファイル作成
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    printTestResult "Test_FileExists_LongPath", fileHandler.fileExists(testFilePath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 7. ファイル存在チェック - 存在しないファイル
Sub Test_FileExists_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFile.txt"
    
    printTestResult "Test_FileExists_NonExistent", Not fileHandler.fileExists(testFilePath)
End Sub

' 8. フォルダ存在チェック - 短いパス
Sub Test_FolderExists_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderExistsShort"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_FolderExists_ShortPath", fileHandler.folderExists(testFolderPath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 9. フォルダ存在チェック - 長いパス
Sub Test_FolderExists_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\FolderK\FolderL\FolderM\FolderN\FolderO\FolderP\FolderQ\FolderR\FolderS\FolderT\FolderU\FolderV\FolderW\FolderX\FolderY\FolderZ"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_FolderExists_LongPath", fileHandler.folderExists(testFolderPath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 10. フォルダ存在チェック - 存在しないフォルダ
Sub Test_FolderExists_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolder"
    
    printTestResult "Test_FolderExists_NonExistent", Not fileHandler.folderExists(testFolderPath)
End Sub

' 11. ファイルコピー - 短いパス
Sub Test_CopyFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileShort.txt"
    destFilePath = ThisWorkbook.path & "\DestFileShort.txt"
    
    ' テストファイル作成
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルのコピー
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_ShortPath", fileHandler.fileExists(destFilePath)
    
    ' 後始末
    Kill sourceFilePath
    Kill destFilePath
End Sub

' 12. ファイルコピー - 長いパス
Sub Test_CopyFile_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFolderPath As String
    Dim destFolderPath As String
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    destFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\DestFolder"
    sourceFilePath = sourceFolderPath & "\SourceFileLong.txt"
    destFilePath = destFolderPath & "\DestFileLong.txt"
    
    ' フォルダとテストファイル作成
    fileHandler.createFolders sourceFolderPath
    fileHandler.createFolders destFolderPath
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルのコピー
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_LongPath", fileHandler.fileExists(destFilePath)
    
    ' 後始末
    fileHandler.deleteFolder sourceFolderPath
    fileHandler.deleteFolder destFolderPath
End Sub

' 13. ファイルコピー - コピー先フォルダが存在しない
Sub Test_CopyFile_DestFolderNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFolderPath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileNoDestFolder.txt"
    destFolderPath = ThisWorkbook.path & "\NoDestFolder"
    destFilePath = destFolderPath & "\DestFileNoDestFolder.txt"
    
    ' テストファイル作成
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルのコピー（フォルダがない場合の作成を有効に）
    fileHandler.copyFile sourceFilePath, destFilePath, True
    printTestResult "Test_CopyFile_DestFolderNotExist", fileHandler.fileExists(destFilePath)
    
    ' 後始末
    Kill sourceFilePath
    fileHandler.deleteFolder destFolderPath
End Sub

' 14. ファイルコピー - コピー元ファイルが存在しない
Sub Test_CopyFile_SourceNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\NonExistentSourceFile.txt"
    destFilePath = ThisWorkbook.path & "\DestFileNonExistent.txt"
    
    On Error Resume Next
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_SourceNotExist", Err.Number = 100 And InStr(Err.Description, "コピー元ファイルが存在しません。") > 0
    On Error GoTo 0
End Sub

' 15. ファイル移動 - 短いパス
Sub Test_MoveFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileMoveShort.txt"
    destFilePath = ThisWorkbook.path & "\DestFileMoveShort.txt"
    
    ' テストファイル作成
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルの移動
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_ShortPath", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' 後始末
    Kill destFilePath
End Sub

' 16. ファイル移動 - 長いパス
Sub Test_MoveFile_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFolderPath As String
    Dim destFolderPath As String
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    destFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\DestFolder"
    sourceFilePath = sourceFolderPath & "\SourceFileMoveLong.txt"
    destFilePath = destFolderPath & "\DestFileMoveLong.txt"
    
    ' フォルダとテストファイル作成
    fileHandler.createFolders sourceFolderPath
    fileHandler.createFolders destFolderPath
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルの移動
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_LongPath", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' 後始末
    fileHandler.deleteFolder sourceFolderPath
    fileHandler.deleteFolder destFolderPath
End Sub

' 17. ファイル移動 - 移動先フォルダが存在しない
Sub Test_MoveFile_DestFolderNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFolderPath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileNoDestFolderMove.txt"
    destFolderPath = ThisWorkbook.path & "\NoDestFolderMove"
    destFilePath = destFolderPath & "\DestFileNoDestFolderMove.txt"
    
    ' テストファイル作成
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルの移動（フォルダがない場合の作成を有効に）
    fileHandler.moveFile sourceFilePath, destFilePath, True
    printTestResult "Test_MoveFile_DestFolderNotExist", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' 後始末
    fileHandler.deleteFolder destFolderPath
End Sub

' 18. ファイル移動 - 移動元ファイルが存在しない
Sub Test_MoveFile_SourceNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\NonExistentSourceFileMove.txt"
    destFilePath = ThisWorkbook.path & "\DestFileNonExistentMove.txt"
    
    On Error Resume Next
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_SourceNotExist", Err.Number = 100 And InStr(Err.Description, "移動元ファイルが存在しません。") > 0
    On Error GoTo 0
End Sub

' 19. ファイル削除 - 短いパス
Sub Test_DeleteFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\TestDeleteFileShort.txt"
    
    ' テストファイル作成
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルの削除
    fileHandler.deleteFile testFilePath
    printTestResult "Test_DeleteFile_ShortPath", Not fileHandler.fileExists(testFilePath)
End Sub

' 20. ファイル削除 - 長いパス
Sub Test_DeleteFile_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF"
    testFilePath = testFolderPath & "\TestDeleteFileLong.txt"
    
    ' フォルダとテストファイル作成
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' ファイルの削除
    fileHandler.deleteFile testFilePath
    printTestResult "Test_DeleteFile_LongPath", Not fileHandler.fileExists(testFilePath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 21. ファイル削除 - 存在しないファイル（エラーを出す）
Sub Test_DeleteFile_NonExistentWithError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFileToDelete.txt"
    
    On Error Resume Next
    fileHandler.deleteFile testFilePath, True
    printTestResult "Test_DeleteFile_NonExistentWithError", Err.Number = 100 And InStr(Err.Description, "削除対象のファイルが存在しません。") > 0
    On Error GoTo 0
End Sub

' 22. ファイル削除 - 存在しないファイル（エラーを出さない）
Sub Test_DeleteFile_NonExistentWithoutError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFileToDeleteNoError.txt"
    
    On Error Resume Next
    fileHandler.deleteFile testFilePath, False
    printTestResult "Test_DeleteFile_NonExistentWithoutError", Err.Number = 0
    On Error GoTo 0
End Sub

' 23. フォルダ削除 - 空のフォルダ
Sub Test_DeleteFolder_Empty()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\EmptyFolderToDelete"
    
    ' フォルダ作成
    fileHandler.createFolders testFolderPath
    
    ' フォルダの削除
    fileHandler.deleteFolder testFolderPath
    printTestResult "Test_DeleteFolder_Empty", Not fileHandler.folderExists(testFolderPath)
End Sub

' 24. フォルダ削除 - ファイルが含まれるフォルダ
Sub Test_DeleteFolder_WithFiles()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderWithFilesToDelete"
    testFilePath = testFolderPath & "\TestFileInFolder.txt"
    
    ' フォルダとファイル作成
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' フォルダの削除
    fileHandler.deleteFolder testFolderPath
    printTestResult "Test_DeleteFolder_WithFiles", Not fileHandler.folderExists(testFolderPath)
End Sub

' 25. フォルダ削除 - 存在しないフォルダ（エラーを出す）
Sub Test_DeleteFolder_NonExistentWithError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToDelete"
    
    On Error Resume Next
    fileHandler.deleteFolder testFolderPath, True
    printTestResult "Test_DeleteFolder_NonExistentWithError", Err.Number = 100 And InStr(Err.Description, "削除対象のフォルダが存在しません。") > 0
    On Error GoTo 0
End Sub

' 26. フォルダ削除 - 存在しないフォルダ（エラーを出さない）
Sub Test_DeleteFolder_NonExistentWithoutError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToDeleteNoError"
    
    On Error Resume Next
    fileHandler.deleteFolder testFolderPath, False
    printTestResult "Test_DeleteFolder_NonExistentWithoutError", Err.Number = 0
    On Error GoTo 0
End Sub

' 27. フォルダクリア - ファイルをすべて削除
Sub Test_ClearFolder_WithFiles()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderToClear"
    testFilePath = testFolderPath & "\TestFileToClear.txt"
    
    ' フォルダとファイル作成
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' フォルダのクリア
    fileHandler.clearFolder testFolderPath
    printTestResult "Test_ClearFolder_WithFiles", Not fileHandler.containsFiles(testFolderPath)
    
    ' 後始末
    fileHandler.deleteFolder testFolderPath
End Sub

' 28. フォルダクリア - 存在しないフォルダ
Sub Test_ClearFolder_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToClear"
    
    On Error Resume Next
    fileHandler.clearFolder testFolderPath
    printTestResult "Test_ClearFolder_NonExistent", Err.Number = 100 And InStr(Err.Description, "フォルダ内のクリア操作に失敗しました。") > 0
    On Error GoTo 0
End Sub

' 29. パス結合 - 短いパス同士の結合
Sub Test_PathCombine_ShortPaths()
    Dim fileHandler As New LongPathFileSystemObject
    Dim basePath As String
    Dim additionalPath As String
    Dim expectedPath As String
    basePath = ThisWorkbook.path & "\BaseFolder"
    additionalPath = "SubFolder"
    expectedPath = basePath & "\" & additionalPath
    
    printTestResult "Test_PathCombine_ShortPaths", fileHandler.pathCombine(basePath, additionalPath) = expectedPath
End Sub

' 30. パス結合 - 短いパスと長いパスの結合
Sub Test_PathCombine_ShortAndLongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim basePath As String
    Dim additionalPath As String
    Dim expectedPath As String
    basePath = ThisWorkbook.path & "\BaseFolder"
    additionalPath = "FolderA\FolderB\FolderC\FolderD\FolderE\SubFolder"
    expectedPath = basePath & "\" & additionalPath
    
    printTestResult "Test_PathCombine_ShortAndLongPath", fileHandler.pathCombine(basePath, additionalPath) = expectedPath
End Sub

' 31. パス結合 - 長いパス同士の結合
Sub Test_PathCombine_LongPaths()
    Dim fileHandler As New LongPathFileSystemObject
    Dim basePath As String
    Dim additionalPath As String
    Dim expectedPath As String
    basePath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE"
    additionalPath = "FolderF\FolderG\FolderH\FolderI\SubFolder"
    expectedPath = basePath & "\" & additionalPath
    
    printTestResult "Test_PathCombine_LongPaths", fileHandler.pathCombine(basePath, additionalPath) = expectedPath
End Sub

' 32. パス結合 - 区切り文字が指定されている場合
Sub Test_PathCombine_WithSeparator()
    Dim fileHandler As New LongPathFileSystemObject
    Dim basePath As String
    Dim additionalPath As String
    Dim expectedPath As String
    basePath = ThisWorkbook.path & "/BaseFolder"
    additionalPath = "SubFolder"
    expectedPath = basePath & "/" & additionalPath
    
    printTestResult "Test_PathCombine_WithSeparator", fileHandler.pathCombine(basePath, additionalPath, "/") = expectedPath
End Sub

' 33. エラーハンドリング - 不正なファイル/フォルダ操作
Sub Test_ErrorHandling_InvalidOperation()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidPath As String
    invalidPath = "?:\InvalidOperation"
    
    On Error Resume Next
    fileHandler.createFolders invalidPath
    printTestResult "Test_ErrorHandling_InvalidOperation", Err.Number = 100
    On Error GoTo 0
End Sub

' 34. 長いパスを短いパスに変換 - 有効な長いパス
Sub Test_ConvertLongToShortPath_ValidLongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim longPath As String
    Dim shortPath As String
    
    ' テスト用の長いパス
    longPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    
    ' フォルダ作成
    fileHandler.createFolders longPath
    
    ' 長いパスを短いパスに変換
    shortPath = fileHandler.convertLongToShortPath(longPath)
    
    ' ショートパスが空でないことを確認
    printTestResult "Test_ConvertLongToShortPath_ValidLongPath", Len(shortPath) > 0
    
    ' 後始末
    fileHandler.deleteFolder longPath
End Sub

' 35. 長いパスを短いパスに変換 - 無効なパス
Sub Test_ConvertLongToShortPath_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidPath As String
    Dim shortPath As String
    
    ' 存在しないパス
    invalidPath = ThisWorkbook.path & "\NonExistentFolder"
    
    ' 長いパスを短いパスに変換
    shortPath = fileHandler.convertLongToShortPath(invalidPath)
    
    ' ショートパスが空であることを確認
    printTestResult "Test_ConvertLongToShortPath_InvalidPath", shortPath = ""
End Sub

' 36. 短いパスを長いパスに変換 - 有効な短いパス
Sub Test_ConvertShortToLongPath_ValidShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim longPath As String
    Dim shortPath As String
    Dim convertedLongPath As String
    
    ' テスト用の長いパス
    longPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    
    ' フォルダ作成
    fileHandler.createFolders longPath
    
    ' 長いパスを短いパスに変換
    shortPath = fileHandler.convertLongToShortPath(longPath)
    
    ' 短いパスを長いパスに変換
    convertedLongPath = fileHandler.convertShortToLongPath(shortPath)
    
    ' 変換された長いパスが元の長いパスと一致することを確認
    printTestResult "Test_ConvertShortToLongPath_ValidShortPath", LCase(convertedLongPath) = LCase(longPath)
    
    ' 後始末
    fileHandler.deleteFolder longPath
End Sub

' 37. 短いパスを長いパスに変換 - 無効な短いパス
Sub Test_ConvertShortToLongPath_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidShortPath As String
    Dim longPath As String
    
    ' 存在しない短いパス
    invalidShortPath = ThisWorkbook.path & "\NonExistentShortPath"
    
    ' 短いパスを長いパスに変換
    longPath = fileHandler.convertShortToLongPath(invalidShortPath)
    
    ' 長いパスが空であることを確認
    printTestResult "Test_ConvertShortToLongPath_InvalidPath", longPath = ""
End Sub

' 全テストを実行するマスターテスト関数
Sub RunAllTests()
    Call Test_CreateFolder_ShortPath
    Call Test_CreateFolder_LongPath
    Call Test_CreateFolder_Existing
    Call Test_CreateFolder_InvalidPath
    Call Test_FileExists_ShortPath
    Call Test_FileExists_LongPath
    Call Test_FileExists_NonExistent
    Call Test_FolderExists_ShortPath
    Call Test_FolderExists_LongPath
    Call Test_FolderExists_NonExistent
    Call Test_CopyFile_ShortPath
    Call Test_CopyFile_LongPath
    Call Test_CopyFile_DestFolderNotExist
    Call Test_CopyFile_SourceNotExist
    Call Test_MoveFile_ShortPath
    Call Test_MoveFile_LongPath
    Call Test_MoveFile_DestFolderNotExist
    Call Test_MoveFile_SourceNotExist
    Call Test_DeleteFile_ShortPath
    Call Test_DeleteFile_LongPath
    Call Test_DeleteFile_NonExistentWithError
    Call Test_DeleteFile_NonExistentWithoutError
    Call Test_DeleteFolder_Empty
    Call Test_DeleteFolder_WithFiles
    Call Test_DeleteFolder_NonExistentWithError
    Call Test_DeleteFolder_NonExistentWithoutError
    Call Test_ClearFolder_WithFiles
    Call Test_ClearFolder_NonExistent
    Call Test_PathCombine_ShortPaths
    Call Test_PathCombine_ShortAndLongPath
    Call Test_PathCombine_LongPaths
    Call Test_PathCombine_WithSeparator
    Call Test_ErrorHandling_InvalidOperation
    Call Test_ConvertLongToShortPath_ValidLongPath
    Call Test_ConvertLongToShortPath_InvalidPath
    Call Test_ConvertShortToLongPath_ValidShortPath
    Call Test_ConvertShortToLongPath_InvalidPath
End Sub

