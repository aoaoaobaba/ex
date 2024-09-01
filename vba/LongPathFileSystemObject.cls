VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongPathFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' LongPathFileSystemObject クラスモジュール

Option Explicit

' Windows API 宣言
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

' 構造体の宣言
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

' 定数の宣言
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const INVALID_HANDLE_VALUE As LongPtr = -1
Private Const FindExInfoStandard As Long = 0
Private Const FindExSearchNameMatch As Long = 0
Private Const FIND_FIRST_EX_LARGE_FETCH As Long = 2
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

' 長いパスのプレフィックス
Private Const LONG_PATH_PREFIX As String = "\\?\"
Private Const DEFAULT_SEPARATOR As String = "\"

' エラーハンドリングのためのカスタムエラーナンバー
Private Const ERR_FILE_OPERATION As Long = vbObjectError + 100

'==============================================================================
' Function: getFullPath
' 説明: 指定されたパスに長いパスプレフィックスを追加して返します。
' パラメータ:
'   - ByVal path As String: 処理するファイルまたはフォルダのパス。
' 戻り値:
'   - String: 長いパスプレフィックスを追加したパス。
'==============================================================================
Public Function getFullPath(ByVal path As String) As String
    getFullPath = LONG_PATH_PREFIX & path
End Function

'==============================================================================
' Function: convertLongToShortPath
' 説明: 長いパスをショートパスに変換します。
' パラメータ:
'   - ByVal longPath As String: 変換する長いパス。
' 戻り値:
'   - String: 変換されたショートパス。変換に失敗した場合はエラーメッセージが発生します。
'==============================================================================
Public Function convertLongToShortPath(ByVal longPath As String) As String
    Dim shortPath As String
    Dim bufferSize As Long
    Dim result As Long
    
    ' バッファサイズを設定
    bufferSize = 32767 ' 最大バッファサイズ
    shortPath = String(bufferSize, vbNullChar)
    
    ' ロングパスをショートパスに変換
    result = GetShortPathNameW(StrPtr(longPath), StrPtr(shortPath), bufferSize)
    
    If result > 0 And result < bufferSize Then
        convertLongToShortPath = Left(shortPath, result)
    Else
        Call handleError("convertLongToShortPath", "ロングパスからショートパスへの変換に失敗しました。", longPath)
    End If
End Function

'==============================================================================
' Function: convertShortToLongPath
' 説明: ショートパスを長いパスに変換します。
' パラメータ:
'   - ByVal shortPath As String: 変換するショートパス。
' 戻り値:
'   - String: 変換された長いパス。変換に失敗した場合はエラーメッセージが発生します。
'==============================================================================
Public Function convertShortToLongPath(ByVal shortPath As String) As String
    Dim longPath As String
    Dim bufferSize As Long
    Dim result As Long
    
    ' バッファサイズを設定
    bufferSize = 32767 ' 最大バッファサイズ
    longPath = String(bufferSize, vbNullChar)
    
    ' ショートパスをロングパスに変換
    result = GetLongPathNameW(StrPtr(shortPath), StrPtr(longPath), bufferSize)
    
    If result > 0 And result < bufferSize Then
        convertShortToLongPath = Left(longPath, result)
    Else
        Call handleError("convertShortToLongPath", "ショートパスからロングパスへの変換に失敗しました。", shortPath)
    End If
End Function

'==============================================================================
' Function: pathCombine
' 説明: 指定された2つのパスを連結し、適切な区切り文字を追加します。
' パラメータ:
'   - ByVal basePath As String: 基本となるパス。
'   - ByVal additionalPath As String: 追加するパス。
'   - Optional ByVal separator As String = "\" : パスの区切り文字（デフォルトは "\"）。
' 戻り値:
'   - String: 連結された完全なパス。
'==============================================================================
Public Function pathCombine(ByVal basePath As String, ByVal additionalPath As String, Optional ByVal separator As String = DEFAULT_SEPARATOR) As String
    ' basePath の末尾に separator が存在しない場合、追加する
    If Right(basePath, 1) <> separator Then
        basePath = basePath & separator
    End If
    
    ' 完全なパスを返す
    pathCombine = basePath & additionalPath
End Function

'==============================================================================
' Subroutine: createFolders
' 説明: 指定されたパスにフォルダを再帰的に作成します。
' パラメータ:
'   - ByVal folderPath As String: 作成するフォルダのパス。
'==============================================================================
Public Sub createFolders(ByVal folderPath As String)
    ' フォルダが存在しない場合のみ作成
    If Not folderExists(folderPath) Then
        Dim parentPath As String

        ' 親フォルダのパスを取得
        parentPath = Left(folderPath, InStrRev(folderPath, DEFAULT_SEPARATOR) - 1)
        
        ' 親フォルダが存在しない場合、再帰的に作成
        If Not folderExists(parentPath) Then
            Call createFolders(parentPath)
        End If
        
        ' フォルダの作成
        If CreateDirectoryW(StrPtr(getFullPath(folderPath)), 0) = 0 Then
            Call handleError("createFolders", "フォルダの作成に失敗しました。", folderPath)
        End If
    End If
End Sub

'==============================================================================
' Function: folderExists
' 説明: 指定されたパスにフォルダが存在するかどうかをチェックします。
' パラメータ:
'   - ByVal folderPath As String: チェックするフォルダのパス。
' 戻り値:
'   - Boolean: フォルダが存在する場合は True、存在しない場合は False。
'==============================================================================
Public Function folderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    ' フォルダの属性を取得
    attributes = GetFileAttributesW(StrPtr(getFullPath(folderPath)))
    
    ' 属性が無効でないか、ディレクトリであるかをチェック
    If attributes = INVALID_FILE_ATTRIBUTES Then
        folderExists = False
    Else
        folderExists = (attributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
    End If
End Function

'==============================================================================
' Function: fileExists
' 説明: 指定されたパスにファイルが存在するかどうかをチェックします。
' パラメータ:
'   - ByVal filePath As String: チェックするファイルのパス。
' 戻り値:
'   - Boolean: ファイルが存在する場合は True、存在しない場合は False。
'==============================================================================
Public Function fileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    ' ファイルの属性を取得
    attributes = GetFileAttributesW(StrPtr(getFullPath(filePath)))
    
    ' 属性が無効でないか、ディレクトリではないかをチェック
    If attributes = INVALID_FILE_ATTRIBUTES Then
        fileExists = False
    Else
        fileExists = (attributes And FILE_ATTRIBUTE_DIRECTORY) = 0
    End If
End Function

'==============================================================================
' Function: containsFiles
' 説明: 指定されたフォルダおよびそのサブフォルダ内に特定のパターンを除く
'       ファイルまたはサブフォルダが含まれているかどうかを再帰的にチェックします。
' パラメータ:
'   - ByVal folderPath As String: チェックするフォルダのパス。
'   - Optional ByVal excludePattern As String: チェック対象外とするファイルやフォルダのパターン。
'                                               省略可能で、省略時はすべてのファイルとフォルダをチェックします。
' 戻り値:
'   - Boolean: フォルダまたはサブフォルダにファイルまたはサブフォルダが含まれている場合は True、
'              含まれていない場合は False。
'==============================================================================
Public Function containsFiles(ByVal folderPath As String, Optional ByVal excludePattern As String = "") As Boolean
    Dim fileList As Collection

    ' ファイルのリストを取得
    Set fileList = getAllFiles(folderPath, excludePattern)
    
    ' ファイルリストが空でないかをチェック
    containsFiles = (fileList.Count > 0)
End Function

'==============================================================================
' Function: getAllFiles
' 説明: 指定されたフォルダ内およびそのサブフォルダ内のすべてのファイルを取得します。
' パラメータ:
'   - ByVal folderPath As String: 検索するフォルダのパス。
'   - Optional ByVal excludePattern As String: 除外するファイルやフォルダのパターン（省略可能）。
' 戻り値:
'   - Collection: 見つかったすべてのファイルパスを含むコレクション。
'==============================================================================
Public Function getAllFiles(ByVal folderPath As String, Optional ByVal excludePattern As String = "") As Collection
    Dim fileList As New Collection

    ' 再帰的にファイルをリスト化
    Call listFilesRecursive(folderPath, fileList, excludePattern, False)
    
    Set getAllFiles = fileList
End Function

'==============================================================================
' Subroutine: copyFile
' 説明: 指定されたファイルをコピー元からコピー先にコピーします。
' パラメータ:
'   - ByVal sourceFilePath As String: コピー元のファイルパス。
'   - ByVal destinationFilePath As String: コピー先のファイルパス。
'   - Optional ByVal ensureDestinationFolder As Boolean: コピー先のフォルダが存在しない場合は作成するかどうか（デフォルトは False）。
'==============================================================================
Public Sub copyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal ensureDestinationFolder As Boolean = False)
    ' コピー元ファイルの存在チェック
    If Not fileExists(sourceFilePath) Then
        Call handleError("copyFile", "コピー元ファイルが存在しません。", sourceFilePath)
        Exit Sub
    End If

    ' コピー先のフォルダが存在しない場合に作成する
    If ensureDestinationFolder Then
        Call ensureFolderExists(destinationFilePath)
    End If
    
    ' ファイルのコピーを実行
    If CopyFileW(StrPtr(getFullPath(sourceFilePath)), StrPtr(getFullPath(destinationFilePath)), 0) = 0 Then
        Call handleError("copyFile", "ファイルのコピーに失敗しました。", sourceFilePath & " -> " & destinationFilePath)
    End If
End Sub

'==============================================================================
' Subroutine: moveFile
' 説明: 指定されたファイルを移動元から移動先に移動します。
' パラメータ:
'   - ByVal sourceFilePath As String: 移動元のファイルパス。
'   - ByVal destinationFilePath As String: 移動先のファイルパス。
'   - Optional ByVal ensureDestinationFolder As Boolean: 移動先のフォルダが存在しない場合は作成するかどうか（デフォルトは False）。
'==============================================================================
Public Sub moveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal ensureDestinationFolder As Boolean = False)
    ' 移動元ファイルの存在チェック
    If Not fileExists(sourceFilePath) Then
        Call handleError("moveFile", "移動元ファイルが存在しません。", sourceFilePath)
        Exit Sub
    End If

    ' 移動先のフォルダが存在しない場合に作成する
    If ensureDestinationFolder Then
        Call ensureFolderExists(destinationFilePath)
    End If

    ' ファイルの移動を実行
    If MoveFileW(StrPtr(getFullPath(sourceFilePath)), StrPtr(getFullPath(destinationFilePath))) = 0 Then
        Call handleError("moveFile", "ファイルの移動に失敗しました。", sourceFilePath & " -> " & destinationFilePath)
    End If
End Sub

'==============================================================================
' Subroutine: deleteFile
' 説明: 指定されたファイルを削除します。
' パラメータ:
'   - ByVal filePath As String: 削除するファイルのパス。
'   - Optional ByVal raiseErrorIfNotExists As Boolean: ファイルが存在しない場合にエラーを発生させるかどうか（デフォルトは True）。
'==============================================================================
Public Sub deleteFile(ByVal filePath As String, Optional ByVal raiseErrorIfNotExists As Boolean = True)
    ' ファイルの存在チェック
    If Not fileExists(filePath) Then
        If raiseErrorIfNotExists Then
            Call handleError("deleteFile", "削除対象のファイルが存在しません。", filePath)
        End If
        Exit Sub
    End If
    
    ' ファイルの削除を実行
    If DeleteFileW(StrPtr(getFullPath(filePath))) = 0 Then
        Call handleError("deleteFile", "ファイルの削除に失敗しました。", filePath)
    End If
End Sub

'==============================================================================
' Subroutine: clearFolder
' 説明: 指定されたフォルダ内のすべてのファイルとサブフォルダをクリアします。
' パラメータ:
'   - ByVal folderPath As String: クリアするフォルダのパス。
'   - Optional ByVal excludePattern As String: 除外するファイルやフォルダのパターン（省略可能）。
'==============================================================================
Public Sub clearFolder(ByVal folderPath As String, Optional ByVal excludePattern As String = "")
    Dim fileList As Object
    Dim key As Variant
    
    ' Dictionary を初期化
    Set fileList = CreateObject("Scripting.Dictionary")
    
    ' 再帰的にフォルダ内のすべてのファイルとフォルダを取得（削除用）
    Call listFilesRecursive(folderPath, fileList, excludePattern, True)
    
    ' リストの要素を順に処理（最も内側から削除されるように）
    For Each key In fileList.Keys
        If fileList(key) = True Then ' Trueの場合はディレクトリ
            ' フォルダの場合、RemoveDirectoryWを使用して削除
            If RemoveDirectoryW(StrPtr(getFullPath(key))) = 0 Then
                Call handleError("clearFolder", "サブフォルダの削除に失敗しました。", key)
            End If
        Else
            ' ファイルの場合、DeleteFileWを使用して削除
            If DeleteFileW(StrPtr(getFullPath(key))) = 0 Then
                Call handleError("clearFolder", "ファイルの削除に失敗しました。", key)
            End If
        End If
    Next key
End Sub

'==============================================================================
' Subroutine: deleteFolder
' 説明: 指定されたフォルダとその中身をすべて削除します。
' パラメータ:
'   - ByVal folderPath As String: 削除するフォルダのパス。
'   - Optional ByVal raiseErrorIfNotExists As Boolean: フォルダが存在しない場合にエラーを発生させるかどうか（デフォルトは True）。
'==============================================================================
Public Sub deleteFolder(ByVal folderPath As String, Optional ByVal raiseErrorIfNotExists As Boolean = True)
    ' フォルダの存在チェック
    If Not folderExists(folderPath) Then
        If raiseErrorIfNotExists Then
            Call handleError("deleteFolder", "削除対象のフォルダが存在しません。", folderPath)
        End If
        Exit Sub
    End If

    ' フォルダ内をクリアする
    Call clearFolder(folderPath)
    
    ' フォルダ自体を削除する
    If RemoveDirectoryW(StrPtr(getFullPath(folderPath))) = 0 Then
        Call handleError("deleteFolder", "フォルダの削除に失敗しました。", folderPath)
    End If
End Sub

'==================== Private 関数 ====================

'==============================================================================
' Function: findFiles
' 説明: 指定されたフォルダのファイル検索を初期化し、検索ハンドルを返します。
'       エラーチェックを行い、検索が失敗した場合はエラーを発生させます。
' パラメータ:
'   - ByVal folderPath As String: 検索するフォルダのパス。
'   - ByRef findData As WIN32_FIND_DATA: ファイル情報を格納する構造体（参照渡し）。
' 戻り値:
'   - LongPtr: 検索ハンドル（失敗した場合はエラーを発生させる）。
'==============================================================================
Private Function findFiles(ByVal folderPath As String, ByRef findData As WIN32_FIND_DATA) As LongPtr
    ' FindFirstFileExW を使用してファイル検索を開始
    findFiles = FindFirstFileExW(StrPtr(getFullPath(folderPath & "\*")), _
                                 FindExInfoStandard, findData, FindExSearchNameMatch, _
                                 0, FIND_FIRST_EX_LARGE_FETCH)
    
    ' 検索ハンドルが無効な場合、エラーを発生させる
    If findFiles = INVALID_HANDLE_VALUE Then
        Call handleError("findFiles", "ファイル検索の初期化に失敗しました。", folderPath)
    End If
End Function

'==============================================================================
' Subroutine: ensureFolderExists
' 説明: 指定されたファイルパスからフォルダを抽出し、そのフォルダが存在しない場合は作成します。
' パラメータ:
'   - ByVal filePath As String: チェックするファイルのパス。
'==============================================================================
Private Sub ensureFolderExists(ByVal filePath As String)
    Dim folderPath As String

    ' ファイルパスからフォルダパスを抽出
    folderPath = Left(filePath, InStrRev(filePath, DEFAULT_SEPARATOR) - 1)
    
    ' フォルダが存在しない場合は作成
    If Not folderExists(folderPath) Then
        Call createFolders(folderPath)
    End If
End Sub

'==============================================================================
' Subroutine: listFilesRecursive
' 説明: 指定されたフォルダ内およびそのサブフォルダ内のすべてのファイルを再帰的にリスト化します。
' パラメータ:
'   - ByVal folderPath As String: 検索するフォルダのパス。
'   - ByRef fileList As Object: ファイル情報を格納するDictionaryオブジェクト（参照渡し）。
'   - ByVal excludePattern As String: 除外するファイルやフォルダのパターン（省略可能）。
'   - ByVal includeFolders As Boolean: フォルダもリストに含めるかどうか（デフォルトは False）。
'==============================================================================
Private Sub listFilesRecursive(ByVal folderPath As String, ByRef fileList As Object, ByVal excludePattern As String, ByVal includeFolders As Boolean)
    Dim findData As WIN32_FIND_DATA
    Dim hFind As LongPtr
    Dim fileName As String
    Dim fullPath As String
    Dim isDirectory As Boolean
    
    ' 指定されたフォルダのファイルを検索
    hFind = findFiles(folderPath, findData)
    
    ' 検索されたファイルやフォルダをループで処理
    Do
        ' ファイル名を取得
        fileName = Left$(findData.cFileName, InStr(findData.cFileName, Chr(0)) - 1)
        fullPath = pathCombine(folderPath, fileName)
        
        ' カレントディレクトリと親ディレクトリを除外
        If fileName <> "." And fileName <> ".." Then
            ' 除外パターンに一致しない場合
            If excludePattern = "" Or InStr(fileName, excludePattern) = 0 Then
                isDirectory = (findData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
                
                ' サブフォルダの場合、先に再帰呼び出し
                If isDirectory Then
                    Call listFilesRecursive(fullPath, fileList, excludePattern, includeFolders)
                    
                    ' フォルダもリストに含める場合、リストに追加（フォルダが先に削除されないように）
                    If includeFolders Then
                        fileList.Add fullPath, isDirectory
                    End If
                Else
                    ' ファイルの場合、直接リストに追加
                    fileList.Add fullPath, isDirectory
                End If
            End If
        End If
    Loop While FindNextFileW(hFind, findData)
    
    ' ファイル検索ハンドルを閉じる
    Call FindClose(hFind)
End Sub

'==============================================================================
' Subroutine: handleError
' 説明: エラーハンドリングの統一を行うためのサブプロシージャ。API呼び出しのエラー情報を含む。
' パラメータ:
'   - ByVal methodName As String: エラーが発生したメソッド名。
'   - ByVal errorMessage As String: エラーメッセージ。
'   - ByVal additionalInfo As String: 追加の情報（例: 関連するファイルパス）。
'==============================================================================
Private Sub handleError(ByVal methodName As String, ByVal errorMessage As String, ByVal additionalInfo As String)
    Dim lastError As Long
    Dim errorBuffer As String * 256
    Dim messageLength As Long
    
    ' 最後のエラーコードを取得
    lastError = GetLastError()
    
    ' エラーメッセージをフォーマット
    messageLength = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lastError, 0, StrPtr(errorBuffer), Len(errorBuffer), 0)
    
    If messageLength > 0 Then
        errorMessage = errorMessage & vbCrLf & "詳細: " & additionalInfo & vbCrLf & "システムエラーメッセージ: " & Left$(errorBuffer, messageLength)
    Else
        errorMessage = errorMessage & vbCrLf & "詳細: " & additionalInfo & vbCrLf & "エラーコード: " & lastError
    End If
    
    ' エラーを発生させる
    Err.Raise ERR_FILE_OPERATION, methodName, errorMessage
End Sub
