Attribute VB_Name = "mdl_init"
Option Explicit
Option Base 0
Private Function bas2NewBook(FilePath As Variant) As String() 'bas����荞�݁A���̃��W���[������Ԃ�
    Dim wb As Workbook
    Dim i As Long
    Dim Extension As String
    Dim buf() As String
    Set wb = Workbooks.Add
    
    ReDim buf(UBound(FilePath)) As String
    With wb.VBProject.VBComponents
        For i = LBound(FilePath) To UBound(FilePath)
            Extension = LCase(Right(FilePath(i), 4))
            If Extension = ".bas" Or Extension = ".cls" Or Extension = ".frm" Then
                .Import FilePath(i) '���W���[���捞
                buf(i) = .Item(.Count).Name '�捞�σ��W���[�����擾
            End If
        Next i
    End With
    bas2NewBook = buf
    Exit Function
End Function
Private Function GetAllFileNames(Optional strFolderPath As String) As String()  '�w��t�H���_�̃t�@�C���̖��O�����ׂĕԂ�
    Dim i As Long
    Dim buf() As String
    Dim objFile As Object
    
    If strFolderPath = "" Then strFolderPath = ThisWorkbook.Path '�t�H���_���w�肳��Ă��Ȃ���΂��̃t�@�C���̃t�H���_���w��
    
    With CreateObject("Scripting.FileSystemObject").GetFolder(strFolderPath)
        ReDim buf(.Files.Count - 1)
        i = 0
        For Each objFile In .Files
             buf(i) = objFile.Path '�p�X���擾
             i = i + 1
        Next objFile
    End With
    GetAllFileNames = buf
End Function
Public Function main(Optional FolderPath As String) '�w��t�H���_��bas�t�@�C�������ׂēǂݍ���
    If FolderPath = "" Then FolderPath = ThisWorkbook.Path & "\lib"
    Call bas2NewBook(GetAllFileNames(FolderPath))
End Function
