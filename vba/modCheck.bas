Attribute VB_Name = "modCheck"
Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' �V���[�g�p�X���g���邩�m�F
Sub CheckShortPath()
    Dim longPath As String
    Dim shortPath As String
    Dim bufferSize As Long
    
    ' TODO: ���m�F���钷���p�X���w��
    longPath = "\\ServerName\ShareFolder\DeepPath\�����p�X�̃e�X�g.xlsx"
    
    ' �o�b�t�@�T�C�Y���w��
    bufferSize = 260
    shortPath = String(bufferSize, vbNullChar)
    
    ' �V���[�g�p�X���擾
    bufferSize = GetShortPathName(longPath, shortPath, bufferSize)
    
    If bufferSize > 0 Then
        shortPath = Left(shortPath, bufferSize)
        MsgBox "�V���[�g�p�X�����p�\�ł�: " & shortPath
    Else
        MsgBox "�V���[�g�p�X���擾�ł��܂���ł����B"
    End If
End Sub


