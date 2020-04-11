Attribute VB_Name = "mdl_worksheet"
Option Explicit
Option Base 0
'''Function
Public Function GetShapesName() As String() '�V�[�g�̃I�u�W�G�N�g�̖��O�ꗗ��z��ŕԂ�
    Dim objShp As Shape
    Dim i As Long
    Dim buf() As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws
        If .Shapes.Count = 0 Then Exit Function
        
        ReDim buf(.Shapes.Count - 1)
        i = 0
        For Each objShp In .Shapes
            buf(i) = objShp.Name
            i = i + 1
        Next
    End With
    GetShapesName = buf
End Function
Public Function GetRGB_InteriorColor(argRange As Range) As Long() '�����Z����RGB�̒l��Ԃ�
    Dim ColorNum As Long
    Dim RGB(2) As Long
    ColorNum = argRange.Interior.Color
    
    RGB(0) = ColorNum Mod 256 'red
    RGB(1) = Int(ColorNum / 256) Mod 256  'green
    RGB(2) = Int(ColorNum / (256 ^ 2)) 'blue
    GetRGB_InteriorColor = RGB
End Function
Public Function GetEnviron() As String() '���ϐ����擾
    Dim i As Long
    Dim buf(50) As String
    i = 1
    Do Until Environ(i) = ""
        buf(i - 1) = Environ(i)
        i = i + 1
    Loop
    GetEnviron = buf
End Function
Public Function GetFileNameFromPath(FilePath As String) '�t�@�C���p�X����t�@�C�����𒊏o��Dir()
    Stop
    GetFileNameFromPath = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
End Function
Public Function Moment(argArray As Variant, Dimention As Long) As Double '���ϒl����̒������[�����g:E[Z(x-l�)Ad]
    Dim buf As Double
    Dim mu As Double
    Dim i As Long
    buf = 0
    mu = WorksheetFunction.Average(argArray)
    If Dimention = 1 Then '����
        buf = mu
    ElseIf Dimention > 1 Then '���ϒl����̒������[�����g
        For i = 1 To UBound(argArray)
            buf = buf + (argArray(i) - mu) ^ Dimention
        Next i
        buf = buf / (i - 1)
    End If
    Moment = buf
End Function
Public Function MaxDrawDown(argArray As Variant) As Double '�ő�h���[�_�E��
    Dim Peak As Double
    Dim MaxDD As Double
    Dim i As Long
    MaxDD = 0
    Peak = 0
    For i = 0 To UBound(argArray)
        Peak = WorksheetFunction.Max(argArray(i), Peak)
        MaxDD = WorksheetFunction.Max(Peak / argArray(i) - 1, MaxDD)
    Next i
    MaxDrawDown = MaxDD
End Function
'''ScriptControl
Public Function EncodeURL(URL As String) As String '���{��URL�̃G���R�[�h
    '�Q�lhttps://colo-ri.jp/develop/2009/07/urlexcelurl.html
    Dim ScriptControl As Object
    Dim Jscript As Object
    Set ScriptControl = CreateObject("ScriptControl")
    ScriptControl.Language = "Jscript"
    Set Jscript = sc.CodeOdect
    EncodeURL = Jscript.encodeURIComponent(URL)
End Function
