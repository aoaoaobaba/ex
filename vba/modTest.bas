Attribute VB_Name = "modTest"
Option Explicit

' �e�X�g���ʕ\���p�̃T�u�v���V�[�W��
Private Sub printTestResult(testName As String, condition As Boolean)
    If condition Then
        Debug.Print testName & ": ����"
    Else
        Debug.Print testName & ": ���s"
    End If
End Sub

' 1. �t�H���_�쐬 - �Z���p�X
Sub Test_CreateFolder_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderShort"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_ShortPath", fileHandler.folderExists(testFolderPath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 2. �t�H���_�쐬 - �����p�X
Sub Test_CreateFolder_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    ' �����Œ����p�X���쐬�i260�����j
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\FolderK\FolderL\FolderM\FolderN\FolderO\FolderP\FolderQ\FolderR\FolderS\FolderT\FolderU\FolderV\FolderW\FolderX\FolderY\FolderZ"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_LongPath", fileHandler.folderExists(testFolderPath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 3. �t�H���_�쐬 - �����̃t�H���_
Sub Test_CreateFolder_Existing()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderExisting"
    
    fileHandler.createFolders testFolderPath
    On Error Resume Next
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_Existing", Err.Number = 0
    
    ' ��n��
    On Error GoTo 0
    fileHandler.deleteFolder testFolderPath
End Sub

' 4. �t�H���_�쐬 - �s���ȃp�X
Sub Test_CreateFolder_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = "?:\InvalidPath"
    
    On Error Resume Next
    fileHandler.createFolders testFolderPath
    printTestResult "Test_CreateFolder_InvalidPath", Err.Number = 100 And InStr(Err.Description, "�t�H���_�̍쐬�Ɏ��s���܂����B") > 0
    On Error GoTo 0
End Sub

' 5. �t�@�C�����݃`�F�b�N - �Z���p�X
Sub Test_FileExists_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\TestFileShort.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    printTestResult "Test_FileExists_ShortPath", fileHandler.fileExists(testFilePath)
    
    ' ��n��
    Kill testFilePath
End Sub

' 6. �t�@�C�����݃`�F�b�N - �����p�X
Sub Test_FileExists_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    ' �����p�X��260�����ɐݒ�
    testFolderPath = ThisWorkbook.path & "\LongPathFolder\SubFolder1\SubFolder2\SubFolder3\SubFolder4\SubFolder5\SubFolder6\SubFolder7\SubFolder8\SubFolder9\SubFolder10\SubFolder11\SubFolder12\SubFolder13\SubFolder14\SubFolder15\SubFolder16\SubFolder17\SubFolder18\SubFolder19\SubFolder20\SubFolder21\SubFolder22\SubFolder23\SubFolder24\SubFolder25"
    testFilePath = testFolderPath & "\TestFileLong.txt"
    
    ' �t�H���_�ƃe�X�g�t�@�C���쐬
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    printTestResult "Test_FileExists_LongPath", fileHandler.fileExists(testFilePath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 7. �t�@�C�����݃`�F�b�N - ���݂��Ȃ��t�@�C��
Sub Test_FileExists_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFile.txt"
    
    printTestResult "Test_FileExists_NonExistent", Not fileHandler.fileExists(testFilePath)
End Sub

' 8. �t�H���_���݃`�F�b�N - �Z���p�X
Sub Test_FolderExists_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\TestFolderExistsShort"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_FolderExists_ShortPath", fileHandler.folderExists(testFolderPath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 9. �t�H���_���݃`�F�b�N - �����p�X
Sub Test_FolderExists_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ\FolderK\FolderL\FolderM\FolderN\FolderO\FolderP\FolderQ\FolderR\FolderS\FolderT\FolderU\FolderV\FolderW\FolderX\FolderY\FolderZ"
    
    fileHandler.createFolders testFolderPath
    printTestResult "Test_FolderExists_LongPath", fileHandler.folderExists(testFolderPath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 10. �t�H���_���݃`�F�b�N - ���݂��Ȃ��t�H���_
Sub Test_FolderExists_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolder"
    
    printTestResult "Test_FolderExists_NonExistent", Not fileHandler.folderExists(testFolderPath)
End Sub

' 11. �t�@�C���R�s�[ - �Z���p�X
Sub Test_CopyFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileShort.txt"
    destFilePath = ThisWorkbook.path & "\DestFileShort.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̃R�s�[
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_ShortPath", fileHandler.fileExists(destFilePath)
    
    ' ��n��
    Kill sourceFilePath
    Kill destFilePath
End Sub

' 12. �t�@�C���R�s�[ - �����p�X
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
    
    ' �t�H���_�ƃe�X�g�t�@�C���쐬
    fileHandler.createFolders sourceFolderPath
    fileHandler.createFolders destFolderPath
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̃R�s�[
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_LongPath", fileHandler.fileExists(destFilePath)
    
    ' ��n��
    fileHandler.deleteFolder sourceFolderPath
    fileHandler.deleteFolder destFolderPath
End Sub

' 13. �t�@�C���R�s�[ - �R�s�[��t�H���_�����݂��Ȃ�
Sub Test_CopyFile_DestFolderNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFolderPath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileNoDestFolder.txt"
    destFolderPath = ThisWorkbook.path & "\NoDestFolder"
    destFilePath = destFolderPath & "\DestFileNoDestFolder.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̃R�s�[�i�t�H���_���Ȃ��ꍇ�̍쐬��L���Ɂj
    fileHandler.copyFile sourceFilePath, destFilePath, True
    printTestResult "Test_CopyFile_DestFolderNotExist", fileHandler.fileExists(destFilePath)
    
    ' ��n��
    Kill sourceFilePath
    fileHandler.deleteFolder destFolderPath
End Sub

' 14. �t�@�C���R�s�[ - �R�s�[���t�@�C�������݂��Ȃ�
Sub Test_CopyFile_SourceNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\NonExistentSourceFile.txt"
    destFilePath = ThisWorkbook.path & "\DestFileNonExistent.txt"
    
    On Error Resume Next
    fileHandler.copyFile sourceFilePath, destFilePath
    printTestResult "Test_CopyFile_SourceNotExist", Err.Number = 100 And InStr(Err.Description, "�R�s�[���t�@�C�������݂��܂���B") > 0
    On Error GoTo 0
End Sub

' 15. �t�@�C���ړ� - �Z���p�X
Sub Test_MoveFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileMoveShort.txt"
    destFilePath = ThisWorkbook.path & "\DestFileMoveShort.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̈ړ�
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_ShortPath", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' ��n��
    Kill destFilePath
End Sub

' 16. �t�@�C���ړ� - �����p�X
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
    
    ' �t�H���_�ƃe�X�g�t�@�C���쐬
    fileHandler.createFolders sourceFolderPath
    fileHandler.createFolders destFolderPath
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̈ړ�
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_LongPath", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' ��n��
    fileHandler.deleteFolder sourceFolderPath
    fileHandler.deleteFolder destFolderPath
End Sub

' 17. �t�@�C���ړ� - �ړ���t�H���_�����݂��Ȃ�
Sub Test_MoveFile_DestFolderNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFolderPath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\SourceFileNoDestFolderMove.txt"
    destFolderPath = ThisWorkbook.path & "\NoDestFolderMove"
    destFilePath = destFolderPath & "\DestFileNoDestFolderMove.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open sourceFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̈ړ��i�t�H���_���Ȃ��ꍇ�̍쐬��L���Ɂj
    fileHandler.moveFile sourceFilePath, destFilePath, True
    printTestResult "Test_MoveFile_DestFolderNotExist", fileHandler.fileExists(destFilePath) And Not fileHandler.fileExists(sourceFilePath)
    
    ' ��n��
    fileHandler.deleteFolder destFolderPath
End Sub

' 18. �t�@�C���ړ� - �ړ����t�@�C�������݂��Ȃ�
Sub Test_MoveFile_SourceNotExist()
    Dim fileHandler As New LongPathFileSystemObject
    Dim sourceFilePath As String
    Dim destFilePath As String
    sourceFilePath = ThisWorkbook.path & "\NonExistentSourceFileMove.txt"
    destFilePath = ThisWorkbook.path & "\DestFileNonExistentMove.txt"
    
    On Error Resume Next
    fileHandler.moveFile sourceFilePath, destFilePath
    printTestResult "Test_MoveFile_SourceNotExist", Err.Number = 100 And InStr(Err.Description, "�ړ����t�@�C�������݂��܂���B") > 0
    On Error GoTo 0
End Sub

' 19. �t�@�C���폜 - �Z���p�X
Sub Test_DeleteFile_ShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\TestDeleteFileShort.txt"
    
    ' �e�X�g�t�@�C���쐬
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̍폜
    fileHandler.deleteFile testFilePath
    printTestResult "Test_DeleteFile_ShortPath", Not fileHandler.fileExists(testFilePath)
End Sub

' 20. �t�@�C���폜 - �����p�X
Sub Test_DeleteFile_LongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF"
    testFilePath = testFolderPath & "\TestDeleteFileLong.txt"
    
    ' �t�H���_�ƃe�X�g�t�@�C���쐬
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�@�C���̍폜
    fileHandler.deleteFile testFilePath
    printTestResult "Test_DeleteFile_LongPath", Not fileHandler.fileExists(testFilePath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 21. �t�@�C���폜 - ���݂��Ȃ��t�@�C���i�G���[���o���j
Sub Test_DeleteFile_NonExistentWithError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFileToDelete.txt"
    
    On Error Resume Next
    fileHandler.deleteFile testFilePath, True
    printTestResult "Test_DeleteFile_NonExistentWithError", Err.Number = 100 And InStr(Err.Description, "�폜�Ώۂ̃t�@�C�������݂��܂���B") > 0
    On Error GoTo 0
End Sub

' 22. �t�@�C���폜 - ���݂��Ȃ��t�@�C���i�G���[���o���Ȃ��j
Sub Test_DeleteFile_NonExistentWithoutError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFilePath As String
    testFilePath = ThisWorkbook.path & "\NonExistentFileToDeleteNoError.txt"
    
    On Error Resume Next
    fileHandler.deleteFile testFilePath, False
    printTestResult "Test_DeleteFile_NonExistentWithoutError", Err.Number = 0
    On Error GoTo 0
End Sub

' 23. �t�H���_�폜 - ��̃t�H���_
Sub Test_DeleteFolder_Empty()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\EmptyFolderToDelete"
    
    ' �t�H���_�쐬
    fileHandler.createFolders testFolderPath
    
    ' �t�H���_�̍폜
    fileHandler.deleteFolder testFolderPath
    printTestResult "Test_DeleteFolder_Empty", Not fileHandler.folderExists(testFolderPath)
End Sub

' 24. �t�H���_�폜 - �t�@�C�����܂܂��t�H���_
Sub Test_DeleteFolder_WithFiles()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderWithFilesToDelete"
    testFilePath = testFolderPath & "\TestFileInFolder.txt"
    
    ' �t�H���_�ƃt�@�C���쐬
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�H���_�̍폜
    fileHandler.deleteFolder testFolderPath
    printTestResult "Test_DeleteFolder_WithFiles", Not fileHandler.folderExists(testFolderPath)
End Sub

' 25. �t�H���_�폜 - ���݂��Ȃ��t�H���_�i�G���[���o���j
Sub Test_DeleteFolder_NonExistentWithError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToDelete"
    
    On Error Resume Next
    fileHandler.deleteFolder testFolderPath, True
    printTestResult "Test_DeleteFolder_NonExistentWithError", Err.Number = 100 And InStr(Err.Description, "�폜�Ώۂ̃t�H���_�����݂��܂���B") > 0
    On Error GoTo 0
End Sub

' 26. �t�H���_�폜 - ���݂��Ȃ��t�H���_�i�G���[���o���Ȃ��j
Sub Test_DeleteFolder_NonExistentWithoutError()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToDeleteNoError"
    
    On Error Resume Next
    fileHandler.deleteFolder testFolderPath, False
    printTestResult "Test_DeleteFolder_NonExistentWithoutError", Err.Number = 0
    On Error GoTo 0
End Sub

' 27. �t�H���_�N���A - �t�@�C�������ׂč폜
Sub Test_ClearFolder_WithFiles()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    Dim testFilePath As String
    testFolderPath = ThisWorkbook.path & "\FolderToClear"
    testFilePath = testFolderPath & "\TestFileToClear.txt"
    
    ' �t�H���_�ƃt�@�C���쐬
    fileHandler.createFolders testFolderPath
    Open testFilePath For Output As #1
    Print #1, "Test"
    Close #1
    
    ' �t�H���_�̃N���A
    fileHandler.clearFolder testFolderPath
    printTestResult "Test_ClearFolder_WithFiles", Not fileHandler.containsFiles(testFolderPath)
    
    ' ��n��
    fileHandler.deleteFolder testFolderPath
End Sub

' 28. �t�H���_�N���A - ���݂��Ȃ��t�H���_
Sub Test_ClearFolder_NonExistent()
    Dim fileHandler As New LongPathFileSystemObject
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.path & "\NonExistentFolderToClear"
    
    On Error Resume Next
    fileHandler.clearFolder testFolderPath
    printTestResult "Test_ClearFolder_NonExistent", Err.Number = 100 And InStr(Err.Description, "�t�H���_���̃N���A����Ɏ��s���܂����B") > 0
    On Error GoTo 0
End Sub

' 29. �p�X���� - �Z���p�X���m�̌���
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

' 30. �p�X���� - �Z���p�X�ƒ����p�X�̌���
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

' 31. �p�X���� - �����p�X���m�̌���
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

' 32. �p�X���� - ��؂蕶�����w�肳��Ă���ꍇ
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

' 33. �G���[�n���h�����O - �s���ȃt�@�C��/�t�H���_����
Sub Test_ErrorHandling_InvalidOperation()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidPath As String
    invalidPath = "?:\InvalidOperation"
    
    On Error Resume Next
    fileHandler.createFolders invalidPath
    printTestResult "Test_ErrorHandling_InvalidOperation", Err.Number = 100
    On Error GoTo 0
End Sub

' 34. �����p�X��Z���p�X�ɕϊ� - �L���Ȓ����p�X
Sub Test_ConvertLongToShortPath_ValidLongPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim longPath As String
    Dim shortPath As String
    
    ' �e�X�g�p�̒����p�X
    longPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    
    ' �t�H���_�쐬
    fileHandler.createFolders longPath
    
    ' �����p�X��Z���p�X�ɕϊ�
    shortPath = fileHandler.convertLongToShortPath(longPath)
    
    ' �V���[�g�p�X����łȂ����Ƃ��m�F
    printTestResult "Test_ConvertLongToShortPath_ValidLongPath", Len(shortPath) > 0
    
    ' ��n��
    fileHandler.deleteFolder longPath
End Sub

' 35. �����p�X��Z���p�X�ɕϊ� - �����ȃp�X
Sub Test_ConvertLongToShortPath_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidPath As String
    Dim shortPath As String
    
    ' ���݂��Ȃ��p�X
    invalidPath = ThisWorkbook.path & "\NonExistentFolder"
    
    ' �����p�X��Z���p�X�ɕϊ�
    shortPath = fileHandler.convertLongToShortPath(invalidPath)
    
    ' �V���[�g�p�X����ł��邱�Ƃ��m�F
    printTestResult "Test_ConvertLongToShortPath_InvalidPath", shortPath = ""
End Sub

' 36. �Z���p�X�𒷂��p�X�ɕϊ� - �L���ȒZ���p�X
Sub Test_ConvertShortToLongPath_ValidShortPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim longPath As String
    Dim shortPath As String
    Dim convertedLongPath As String
    
    ' �e�X�g�p�̒����p�X
    longPath = ThisWorkbook.path & "\FolderA\FolderB\FolderC\FolderD\FolderE\FolderF\FolderG\FolderH\FolderI\FolderJ"
    
    ' �t�H���_�쐬
    fileHandler.createFolders longPath
    
    ' �����p�X��Z���p�X�ɕϊ�
    shortPath = fileHandler.convertLongToShortPath(longPath)
    
    ' �Z���p�X�𒷂��p�X�ɕϊ�
    convertedLongPath = fileHandler.convertShortToLongPath(shortPath)
    
    ' �ϊ����ꂽ�����p�X�����̒����p�X�ƈ�v���邱�Ƃ��m�F
    printTestResult "Test_ConvertShortToLongPath_ValidShortPath", LCase(convertedLongPath) = LCase(longPath)
    
    ' ��n��
    fileHandler.deleteFolder longPath
End Sub

' 37. �Z���p�X�𒷂��p�X�ɕϊ� - �����ȒZ���p�X
Sub Test_ConvertShortToLongPath_InvalidPath()
    Dim fileHandler As New LongPathFileSystemObject
    Dim invalidShortPath As String
    Dim longPath As String
    
    ' ���݂��Ȃ��Z���p�X
    invalidShortPath = ThisWorkbook.path & "\NonExistentShortPath"
    
    ' �Z���p�X�𒷂��p�X�ɕϊ�
    longPath = fileHandler.convertShortToLongPath(invalidShortPath)
    
    ' �����p�X����ł��邱�Ƃ��m�F
    printTestResult "Test_ConvertShortToLongPath_InvalidPath", longPath = ""
End Sub

' �S�e�X�g�����s����}�X�^�[�e�X�g�֐�
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

