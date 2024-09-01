VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongPathFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' LongPathFileSystemObject �N���X���W���[��

Option Explicit

' Windows API �錾
Private Declare PtrSafe Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function FindFirstFileExW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal fInfoLevelId As Long, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal fSearchOp As Long, ByVal lpSearchFilter As LongPtr, ByVal dwAdditionalFlags As Long) As LongPtr
Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
Private Declare PtrSafe Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, ByVal bFailIfExists As Long) As Long
Private Declare PtrSafe Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function FormatMessageW Lib "kernel32" (ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal Arguments As LongPtr) As Long
Private Declare PtrSafe Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As LongPtr, ByVal lpszShortPath As LongPtr, ByVal cchBuffer As Long) As Long
Private Declare PtrSafe Function GetLongPathNameW Lib "kernel32" (ByVal lpszShortPath As LongPtr, ByVal lpszLongPath As LongPtr, ByVal cchBuffer As Long) As Long

' �\���̂̐錾
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

' �萔�̐錾
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const INVALID_HANDLE_VALUE As LongPtr = -1
Private Const FindExInfoStandard As Long = 0
Private Const FindExSearchNameMatch As Long = 0
Private Const FIND_FIRST_EX_LARGE_FETCH As Long = 2
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

' �����p�X�̃v���t�B�b�N�X
Private Const LONG_PATH_PREFIX As String = "\\?\"
Private Const DEFAULT_SEPARATOR As String = "\"

' �G���[�n���h�����O�̂��߂̃J�X�^���G���[�i���o�[
Private Const ERR_FILE_OPERATION As Long = vbObjectError + 100

'==============================================================================
' Function: getFullPath
' ����: �w�肳�ꂽ�p�X�ɒ����p�X�v���t�B�b�N�X��ǉ����ĕԂ��܂��B
' �p�����[�^:
'   - ByVal path As String: ��������t�@�C���܂��̓t�H���_�̃p�X�B
' �߂�l:
'   - String: �����p�X�v���t�B�b�N�X��ǉ������p�X�B
'==============================================================================
Public Function getFullPath(ByVal path As String) As String
    getFullPath = LONG_PATH_PREFIX & path
End Function

'==============================================================================
' Function: convertLongToShortPath
' ����: �����p�X���V���[�g�p�X�ɕϊ����܂��B
' �p�����[�^:
'   - ByVal longPath As String: �ϊ����钷���p�X�B
' �߂�l:
'   - String: �ϊ����ꂽ�V���[�g�p�X�B�ϊ��Ɏ��s�����ꍇ�̓G���[���b�Z�[�W���������܂��B
'==============================================================================
Public Function convertLongToShortPath(ByVal longPath As String) As String
    Dim shortPath As String
    Dim bufferSize As Long
    Dim result As Long
    
    ' �o�b�t�@�T�C�Y��ݒ�
    bufferSize = 32767 ' �ő�o�b�t�@�T�C�Y
    shortPath = String(bufferSize, vbNullChar)
    
    ' �����O�p�X���V���[�g�p�X�ɕϊ�
    result = GetShortPathNameW(StrPtr(longPath), StrPtr(shortPath), bufferSize)
    
    If result > 0 And result < bufferSize Then
        convertLongToShortPath = Left(shortPath, result)
    Else
        Call handleError("convertLongToShortPath", "�����O�p�X����V���[�g�p�X�ւ̕ϊ��Ɏ��s���܂����B", longPath)
    End If
End Function

'==============================================================================
' Function: convertShortToLongPath
' ����: �V���[�g�p�X�𒷂��p�X�ɕϊ����܂��B
' �p�����[�^:
'   - ByVal shortPath As String: �ϊ�����V���[�g�p�X�B
' �߂�l:
'   - String: �ϊ����ꂽ�����p�X�B�ϊ��Ɏ��s�����ꍇ�̓G���[���b�Z�[�W���������܂��B
'==============================================================================
Public Function convertShortToLongPath(ByVal shortPath As String) As String
    Dim longPath As String
    Dim bufferSize As Long
    Dim result As Long
    
    ' �o�b�t�@�T�C�Y��ݒ�
    bufferSize = 32767 ' �ő�o�b�t�@�T�C�Y
    longPath = String(bufferSize, vbNullChar)
    
    ' �V���[�g�p�X�������O�p�X�ɕϊ�
    result = GetLongPathNameW(StrPtr(shortPath), StrPtr(longPath), bufferSize)
    
    If result > 0 And result < bufferSize Then
        convertShortToLongPath = Left(longPath, result)
    Else
        Call handleError("convertShortToLongPath", "�V���[�g�p�X���烍���O�p�X�ւ̕ϊ��Ɏ��s���܂����B", shortPath)
    End If
End Function

'==============================================================================
' Function: pathCombine
' ����: �w�肳�ꂽ2�̃p�X��A�����A�K�؂ȋ�؂蕶����ǉ����܂��B
' �p�����[�^:
'   - ByVal basePath As String: ��{�ƂȂ�p�X�B
'   - ByVal additionalPath As String: �ǉ�����p�X�B
'   - Optional ByVal separator As String = "\" : �p�X�̋�؂蕶���i�f�t�H���g�� "\"�j�B
' �߂�l:
'   - String: �A�����ꂽ���S�ȃp�X�B
'==============================================================================
Public Function pathCombine(ByVal basePath As String, ByVal additionalPath As String, Optional ByVal separator As String = DEFAULT_SEPARATOR) As String
    ' basePath �̖����� separator �����݂��Ȃ��ꍇ�A�ǉ�����
    If Right(basePath, 1) <> separator Then
        basePath = basePath & separator
    End If
    
    ' ���S�ȃp�X��Ԃ�
    pathCombine = basePath & additionalPath
End Function

'==============================================================================
' Subroutine: createFolders
' ����: �w�肳�ꂽ�p�X�Ƀt�H���_���ċA�I�ɍ쐬���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: �쐬����t�H���_�̃p�X�B
'==============================================================================
Public Sub createFolders(ByVal folderPath As String)
    ' �t�H���_�����݂��Ȃ��ꍇ�̂ݍ쐬
    If Not folderExists(folderPath) Then
        Dim parentPath As String

        ' �e�t�H���_�̃p�X���擾
        parentPath = Left(folderPath, InStrRev(folderPath, DEFAULT_SEPARATOR) - 1)
        
        ' �e�t�H���_�����݂��Ȃ��ꍇ�A�ċA�I�ɍ쐬
        If Not folderExists(parentPath) Then
            Call createFolders(parentPath)
        End If
        
        ' �t�H���_�̍쐬
        If CreateDirectoryW(StrPtr(getFullPath(folderPath)), 0) = 0 Then
            Call handleError("createFolders", "�t�H���_�̍쐬�Ɏ��s���܂����B", folderPath)
        End If
    End If
End Sub

'==============================================================================
' Function: folderExists
' ����: �w�肳�ꂽ�p�X�Ƀt�H���_�����݂��邩�ǂ������`�F�b�N���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: �`�F�b�N����t�H���_�̃p�X�B
' �߂�l:
'   - Boolean: �t�H���_�����݂���ꍇ�� True�A���݂��Ȃ��ꍇ�� False�B
'==============================================================================
Public Function folderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    ' �t�H���_�̑������擾
    attributes = GetFileAttributesW(StrPtr(getFullPath(folderPath)))
    
    ' �����������łȂ����A�f�B���N�g���ł��邩���`�F�b�N
    If attributes = INVALID_FILE_ATTRIBUTES Then
        folderExists = False
    Else
        folderExists = (attributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
    End If
End Function

'==============================================================================
' Function: fileExists
' ����: �w�肳�ꂽ�p�X�Ƀt�@�C�������݂��邩�ǂ������`�F�b�N���܂��B
' �p�����[�^:
'   - ByVal filePath As String: �`�F�b�N����t�@�C���̃p�X�B
' �߂�l:
'   - Boolean: �t�@�C�������݂���ꍇ�� True�A���݂��Ȃ��ꍇ�� False�B
'==============================================================================
Public Function fileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    ' �t�@�C���̑������擾
    attributes = GetFileAttributesW(StrPtr(getFullPath(filePath)))
    
    ' �����������łȂ����A�f�B���N�g���ł͂Ȃ������`�F�b�N
    If attributes = INVALID_FILE_ATTRIBUTES Then
        fileExists = False
    Else
        fileExists = (attributes And FILE_ATTRIBUTE_DIRECTORY) = 0
    End If
End Function

'==============================================================================
' Function: containsFiles
' ����: �w�肳�ꂽ�t�H���_����т��̃T�u�t�H���_���ɓ���̃p�^�[��������
'       �t�@�C���܂��̓T�u�t�H���_���܂܂�Ă��邩�ǂ������ċA�I�Ƀ`�F�b�N���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: �`�F�b�N����t�H���_�̃p�X�B
'   - Optional ByVal excludePattern As String: �`�F�b�N�ΏۊO�Ƃ���t�@�C����t�H���_�̃p�^�[���B
'                                               �ȗ��\�ŁA�ȗ����͂��ׂẴt�@�C���ƃt�H���_���`�F�b�N���܂��B
' �߂�l:
'   - Boolean: �t�H���_�܂��̓T�u�t�H���_�Ƀt�@�C���܂��̓T�u�t�H���_���܂܂�Ă���ꍇ�� True�A
'              �܂܂�Ă��Ȃ��ꍇ�� False�B
'==============================================================================
Public Function containsFiles(ByVal folderPath As String, Optional ByVal excludePattern As String = "") As Boolean
    Dim fileList As Collection

    ' �t�@�C���̃��X�g���擾
    Set fileList = getAllFiles(folderPath, excludePattern)
    
    ' �t�@�C�����X�g����łȂ������`�F�b�N
    containsFiles = (fileList.Count > 0)
End Function

'==============================================================================
' Function: getAllFiles
' ����: �w�肳�ꂽ�t�H���_������т��̃T�u�t�H���_���̂��ׂẴt�@�C�����擾���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: ��������t�H���_�̃p�X�B
'   - Optional ByVal excludePattern As String: ���O����t�@�C����t�H���_�̃p�^�[���i�ȗ��\�j�B
' �߂�l:
'   - Collection: �����������ׂẴt�@�C���p�X���܂ރR���N�V�����B
'==============================================================================
Public Function getAllFiles(ByVal folderPath As String, Optional ByVal excludePattern As String = "") As Collection
    Dim fileList As New Collection

    ' �ċA�I�Ƀt�@�C�������X�g��
    Call listFilesRecursive(folderPath, fileList, excludePattern, False)
    
    Set getAllFiles = fileList
End Function

'==============================================================================
' Subroutine: copyFile
' ����: �w�肳�ꂽ�t�@�C�����R�s�[������R�s�[��ɃR�s�[���܂��B
' �p�����[�^:
'   - ByVal sourceFilePath As String: �R�s�[���̃t�@�C���p�X�B
'   - ByVal destinationFilePath As String: �R�s�[��̃t�@�C���p�X�B
'   - Optional ByVal ensureDestinationFolder As Boolean: �R�s�[��̃t�H���_�����݂��Ȃ��ꍇ�͍쐬���邩�ǂ����i�f�t�H���g�� False�j�B
'==============================================================================
Public Sub copyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal ensureDestinationFolder As Boolean = False)
    ' �R�s�[���t�@�C���̑��݃`�F�b�N
    If Not fileExists(sourceFilePath) Then
        Call handleError("copyFile", "�R�s�[���t�@�C�������݂��܂���B", sourceFilePath)
        Exit Sub
    End If

    ' �R�s�[��̃t�H���_�����݂��Ȃ��ꍇ�ɍ쐬����
    If ensureDestinationFolder Then
        Call ensureFolderExists(destinationFilePath)
    End If
    
    ' �t�@�C���̃R�s�[�����s
    If CopyFileW(StrPtr(getFullPath(sourceFilePath)), StrPtr(getFullPath(destinationFilePath)), 0) = 0 Then
        Call handleError("copyFile", "�t�@�C���̃R�s�[�Ɏ��s���܂����B", sourceFilePath & " -> " & destinationFilePath)
    End If
End Sub

'==============================================================================
' Subroutine: moveFile
' ����: �w�肳�ꂽ�t�@�C�����ړ�������ړ���Ɉړ����܂��B
' �p�����[�^:
'   - ByVal sourceFilePath As String: �ړ����̃t�@�C���p�X�B
'   - ByVal destinationFilePath As String: �ړ���̃t�@�C���p�X�B
'   - Optional ByVal ensureDestinationFolder As Boolean: �ړ���̃t�H���_�����݂��Ȃ��ꍇ�͍쐬���邩�ǂ����i�f�t�H���g�� False�j�B
'==============================================================================
Public Sub moveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal ensureDestinationFolder As Boolean = False)
    ' �ړ����t�@�C���̑��݃`�F�b�N
    If Not fileExists(sourceFilePath) Then
        Call handleError("moveFile", "�ړ����t�@�C�������݂��܂���B", sourceFilePath)
        Exit Sub
    End If

    ' �ړ���̃t�H���_�����݂��Ȃ��ꍇ�ɍ쐬����
    If ensureDestinationFolder Then
        Call ensureFolderExists(destinationFilePath)
    End If

    ' �t�@�C���̈ړ������s
    If MoveFileW(StrPtr(getFullPath(sourceFilePath)), StrPtr(getFullPath(destinationFilePath))) = 0 Then
        Call handleError("moveFile", "�t�@�C���̈ړ��Ɏ��s���܂����B", sourceFilePath & " -> " & destinationFilePath)
    End If
End Sub

'==============================================================================
' Subroutine: deleteFile
' ����: �w�肳�ꂽ�t�@�C�����폜���܂��B
' �p�����[�^:
'   - ByVal filePath As String: �폜����t�@�C���̃p�X�B
'   - Optional ByVal raiseErrorIfNotExists As Boolean: �t�@�C�������݂��Ȃ��ꍇ�ɃG���[�𔭐������邩�ǂ����i�f�t�H���g�� True�j�B
'==============================================================================
Public Sub deleteFile(ByVal filePath As String, Optional ByVal raiseErrorIfNotExists As Boolean = True)
    ' �t�@�C���̑��݃`�F�b�N
    If Not fileExists(filePath) Then
        If raiseErrorIfNotExists Then
            Call handleError("deleteFile", "�폜�Ώۂ̃t�@�C�������݂��܂���B", filePath)
        End If
        Exit Sub
    End If
    
    ' �t�@�C���̍폜�����s
    If DeleteFileW(StrPtr(getFullPath(filePath))) = 0 Then
        Call handleError("deleteFile", "�t�@�C���̍폜�Ɏ��s���܂����B", filePath)
    End If
End Sub

'==============================================================================
' Subroutine: clearFolder
' ����: �w�肳�ꂽ�t�H���_���̂��ׂẴt�@�C���ƃT�u�t�H���_���N���A���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: �N���A����t�H���_�̃p�X�B
'   - Optional ByVal excludePattern As String: ���O����t�@�C����t�H���_�̃p�^�[���i�ȗ��\�j�B
'==============================================================================
Public Sub clearFolder(ByVal folderPath As String, Optional ByVal excludePattern As String = "")
    Dim fileList As Object
    Dim key As Variant
    
    ' Dictionary ��������
    Set fileList = CreateObject("Scripting.Dictionary")
    
    ' �ċA�I�Ƀt�H���_���̂��ׂẴt�@�C���ƃt�H���_���擾�i�폜�p�j
    Call listFilesRecursive(folderPath, fileList, excludePattern, True)
    
    ' ���X�g�̗v�f�����ɏ����i�ł���������폜�����悤�Ɂj
    For Each key In fileList.Keys
        If fileList(key) = True Then ' True�̏ꍇ�̓f�B���N�g��
            ' �t�H���_�̏ꍇ�ARemoveDirectoryW���g�p���č폜
            If RemoveDirectoryW(StrPtr(getFullPath(key))) = 0 Then
                Call handleError("clearFolder", "�T�u�t�H���_�̍폜�Ɏ��s���܂����B", key)
            End If
        Else
            ' �t�@�C���̏ꍇ�ADeleteFileW���g�p���č폜
            If DeleteFileW(StrPtr(getFullPath(key))) = 0 Then
                Call handleError("clearFolder", "�t�@�C���̍폜�Ɏ��s���܂����B", key)
            End If
        End If
    Next key
End Sub

'==============================================================================
' Subroutine: deleteFolder
' ����: �w�肳�ꂽ�t�H���_�Ƃ��̒��g�����ׂč폜���܂��B
' �p�����[�^:
'   - ByVal folderPath As String: �폜����t�H���_�̃p�X�B
'   - Optional ByVal raiseErrorIfNotExists As Boolean: �t�H���_�����݂��Ȃ��ꍇ�ɃG���[�𔭐������邩�ǂ����i�f�t�H���g�� True�j�B
'==============================================================================
Public Sub deleteFolder(ByVal folderPath As String, Optional ByVal raiseErrorIfNotExists As Boolean = True)
    ' �t�H���_�̑��݃`�F�b�N
    If Not folderExists(folderPath) Then
        If raiseErrorIfNotExists Then
            Call handleError("deleteFolder", "�폜�Ώۂ̃t�H���_�����݂��܂���B", folderPath)
        End If
        Exit Sub
    End If

    ' �t�H���_�����N���A����
    Call clearFolder(folderPath)
    
    ' �t�H���_���̂��폜����
    If RemoveDirectoryW(StrPtr(getFullPath(folderPath))) = 0 Then
        Call handleError("deleteFolder", "�t�H���_�̍폜�Ɏ��s���܂����B", folderPath)
    End If
End Sub

'==================== Private �֐� ====================

'==============================================================================
' Function: findFiles
' ����: �w�肳�ꂽ�t�H���_�̃t�@�C�����������������A�����n���h����Ԃ��܂��B
'       �G���[�`�F�b�N���s���A���������s�����ꍇ�̓G���[�𔭐������܂��B
' �p�����[�^:
'   - ByVal folderPath As String: ��������t�H���_�̃p�X�B
'   - ByRef findData As WIN32_FIND_DATA: �t�@�C�������i�[����\���́i�Q�Ɠn���j�B
' �߂�l:
'   - LongPtr: �����n���h���i���s�����ꍇ�̓G���[�𔭐�������j�B
'==============================================================================
Private Function findFiles(ByVal folderPath As String, ByRef findData As WIN32_FIND_DATA) As LongPtr
    ' FindFirstFileExW ���g�p���ăt�@�C���������J�n
    findFiles = FindFirstFileExW(StrPtr(getFullPath(folderPath & "\*")), _
                                 FindExInfoStandard, findData, FindExSearchNameMatch, _
                                 0, FIND_FIRST_EX_LARGE_FETCH)
    
    ' �����n���h���������ȏꍇ�A�G���[�𔭐�������
    If findFiles = INVALID_HANDLE_VALUE Then
        Call handleError("findFiles", "�t�@�C�������̏������Ɏ��s���܂����B", folderPath)
    End If
End Function

'==============================================================================
' Subroutine: ensureFolderExists
' ����: �w�肳�ꂽ�t�@�C���p�X����t�H���_�𒊏o���A���̃t�H���_�����݂��Ȃ��ꍇ�͍쐬���܂��B
' �p�����[�^:
'   - ByVal filePath As String: �`�F�b�N����t�@�C���̃p�X�B
'==============================================================================
Private Sub ensureFolderExists(ByVal filePath As String)
    Dim folderPath As String

    ' �t�@�C���p�X����t�H���_�p�X�𒊏o
    folderPath = Left(filePath, InStrRev(filePath, DEFAULT_SEPARATOR) - 1)
    
    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Not folderExists(folderPath) Then
        Call createFolders(folderPath)
    End If
End Sub

'==============================================================================
' Subroutine: listFilesRecursive
' ����: �w�肳�ꂽ�t�H���_������т��̃T�u�t�H���_���̂��ׂẴt�@�C�����ċA�I�Ƀ��X�g�����܂��B
' �p�����[�^:
'   - ByVal folderPath As String: ��������t�H���_�̃p�X�B
'   - ByRef fileList As Object: �t�@�C�������i�[����Dictionary�I�u�W�F�N�g�i�Q�Ɠn���j�B
'   - ByVal excludePattern As String: ���O����t�@�C����t�H���_�̃p�^�[���i�ȗ��\�j�B
'   - ByVal includeFolders As Boolean: �t�H���_�����X�g�Ɋ܂߂邩�ǂ����i�f�t�H���g�� False�j�B
'==============================================================================
Private Sub listFilesRecursive(ByVal folderPath As String, ByRef fileList As Object, ByVal excludePattern As String, ByVal includeFolders As Boolean)
    Dim findData As WIN32_FIND_DATA
    Dim hFind As LongPtr
    Dim fileName As String
    Dim fullPath As String
    Dim isDirectory As Boolean
    
    ' �w�肳�ꂽ�t�H���_�̃t�@�C��������
    hFind = findFiles(folderPath, findData)
    
    ' �������ꂽ�t�@�C����t�H���_�����[�v�ŏ���
    Do
        ' �t�@�C�������擾
        fileName = Left$(findData.cFileName, InStr(findData.cFileName, Chr(0)) - 1)
        fullPath = pathCombine(folderPath, fileName)
        
        ' �J�����g�f�B���N�g���Ɛe�f�B���N�g�������O
        If fileName <> "." And fileName <> ".." Then
            ' ���O�p�^�[���Ɉ�v���Ȃ��ꍇ
            If excludePattern = "" Or InStr(fileName, excludePattern) = 0 Then
                isDirectory = (findData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
                
                ' �T�u�t�H���_�̏ꍇ�A��ɍċA�Ăяo��
                If isDirectory Then
                    Call listFilesRecursive(fullPath, fileList, excludePattern, includeFolders)
                    
                    ' �t�H���_�����X�g�Ɋ܂߂�ꍇ�A���X�g�ɒǉ��i�t�H���_����ɍ폜����Ȃ��悤�Ɂj
                    If includeFolders Then
                        fileList.Add fullPath, isDirectory
                    End If
                Else
                    ' �t�@�C���̏ꍇ�A���ڃ��X�g�ɒǉ�
                    fileList.Add fullPath, isDirectory
                End If
            End If
        End If
    Loop While FindNextFileW(hFind, findData)
    
    ' �t�@�C�������n���h�������
    Call FindClose(hFind)
End Sub

'==============================================================================
' Subroutine: handleError
' ����: �G���[�n���h�����O�̓�����s�����߂̃T�u�v���V�[�W���BAPI�Ăяo���̃G���[�����܂ށB
' �p�����[�^:
'   - ByVal methodName As String: �G���[�������������\�b�h���B
'   - ByVal errorMessage As String: �G���[���b�Z�[�W�B
'   - ByVal additionalInfo As String: �ǉ��̏��i��: �֘A����t�@�C���p�X�j�B
'==============================================================================
Private Sub handleError(ByVal methodName As String, ByVal errorMessage As String, ByVal additionalInfo As String)
    Dim lastError As Long
    Dim errorBuffer As String * 256
    Dim messageLength As Long
    
    ' �Ō�̃G���[�R�[�h���擾
    lastError = GetLastError()
    
    ' �G���[���b�Z�[�W���t�H�[�}�b�g
    messageLength = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lastError, 0, StrPtr(errorBuffer), Len(errorBuffer), 0)
    
    If messageLength > 0 Then
        errorMessage = errorMessage & vbCrLf & "�ڍ�: " & additionalInfo & vbCrLf & "�V�X�e���G���[���b�Z�[�W: " & Left$(errorBuffer, messageLength)
    Else
        errorMessage = errorMessage & vbCrLf & "�ڍ�: " & additionalInfo & vbCrLf & "�G���[�R�[�h: " & lastError
    End If
    
    ' �G���[�𔭐�������
    Err.Raise ERR_FILE_OPERATION, methodName, errorMessage
End Sub