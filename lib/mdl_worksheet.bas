Attribute VB_Name = "mdl_worksheet"
Option Explicit
Option Base 0
'''Function
Public Function GetShapesName() As String() 'シートのオブジエクトの名前一覧を配列で返す
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
Public Function GetRGB_InteriorColor(argRange As Range) As Long() '引数セルのRGBの値を返す
    Dim ColorNum As Long
    Dim RGB(2) As Long
    ColorNum = argRange.Interior.Color
    
    RGB(0) = ColorNum Mod 256 'red
    RGB(1) = Int(ColorNum / 256) Mod 256  'green
    RGB(2) = Int(ColorNum / (256 ^ 2)) 'blue
    GetRGB_InteriorColor = RGB
End Function
Public Function GetEnviron() As String() '環境変数を取得
    Dim i As Long
    Dim buf(50) As String
    i = 1
    Do Until Environ(i) = ""
        buf(i - 1) = Environ(i)
        i = i + 1
    Loop
    GetEnviron = buf
End Function
Public Function GetFileNameFromPath(FilePath As String) 'ファイルパスからファイル名を抽出→Dir()
    Stop
    GetFileNameFromPath = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
End Function
Public Function Moment(argArray As Variant, Dimention As Long) As Double '平均値周りの中央モーメント:E[Z(x-l｣)Ad]
    Dim buf As Double
    Dim mu As Double
    Dim i As Long
    buf = 0
    mu = WorksheetFunction.Average(argArray)
    If Dimention = 1 Then '平均
        buf = mu
    ElseIf Dimention > 1 Then '平均値周りの中央モーメント
        For i = 1 To UBound(argArray)
            buf = buf + (argArray(i) - mu) ^ Dimention
        Next i
        buf = buf / (i - 1)
    End If
    Moment = buf
End Function
Public Function MaxDrawDown(argArray As Variant) As Double '最大ドローダウン
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
Public Function EncodeURL(URL As String) As String '日本語URLのエンコード
    '参考https://colo-ri.jp/develop/2009/07/urlexcelurl.html
    Dim ScriptControl As Object
    Dim Jscript As Object
    Set ScriptControl = CreateObject("ScriptControl")
    ScriptControl.Language = "Jscript"
    Set Jscript = sc.CodeOdect
    EncodeURL = Jscript.encodeURIComponent(URL)
End Function
