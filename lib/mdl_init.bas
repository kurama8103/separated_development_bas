Attribute VB_Name = "mdl_init"
Option Explicit
Option Base 0
Private Function bas2NewBook(FilePath As Variant) As String() 'basを取り込み、そのモジュール名を返す
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
                .Import FilePath(i) 'モジュール取込
                buf(i) = .Item(.Count).Name '取込済モジュール名取得
            End If
        Next i
    End With
    bas2NewBook = buf
    Exit Function
End Function
Private Function GetAllFileNames(Optional strFolderPath As String) As String()  '指定フォルダのファイルの名前をすべて返す
    Dim i As Long
    Dim buf() As String
    Dim objFile As Object
    
    If strFolderPath = "" Then strFolderPath = ThisWorkbook.Path 'フォルダが指定されていなければこのファイルのフォルダを指定
    
    With CreateObject("Scripting.FileSystemObject").GetFolder(strFolderPath)
        ReDim buf(.Files.Count - 1)
        i = 0
        For Each objFile In .Files
             buf(i) = objFile.Path 'パス名取得
             i = i + 1
        Next objFile
    End With
    GetAllFileNames = buf
End Function
Public Function main(Optional FolderPath As String) '指定フォルダのbasファイルをすべて読み込む
    If FolderPath = "" Then FolderPath = ThisWorkbook.Path & "\lib"
    Call bas2NewBook(GetAllFileNames(FolderPath))
End Function
