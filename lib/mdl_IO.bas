Attribute VB_Name = "mdl_IO"
'VBComponent系、参照設定系、Property、汎用関数系

'''MEMO
'基本的にはFunction、ワークシートから呼び出すコードはSub
'Subで呼び出す関数の引数は、利便性よりActiveBook/Sheetとしている
'事前バインディングはコーディングが楽、実行速度も速い:New Scripting.FileSystemObject
'実行時バインディングは参照設定が不要で配布が楽：CreateObject("Scripting.FileSystemObject")
'現状ADODB関連のみ事前バインディング
Option Explicit
Option Base 0
'''VBComponent系
Public Function VB_ExportModule() As String 'ブックのすべてのモジュールをエクスポート
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Const SaveFolderName As String = "lib"
        
    Dim objVBC As Object 'VBComponent
    Dim FolderPath As String
    Dim flg As Boolean
    
    '保存先フォルダチェック・作成
    If wb.Path = "" Then Exit Function 'ファイルが一度も保存されていないとパスが取れないため
    FolderPath = wb.Path & "/" & SaveFolderName & "/"
    flg = IsFolderExists(FolderPath)
    If flg = False Then MkDir (FolderPath)
    
    'export
    For Each objVBC In wb.VBProject.VBComponents
        Select Case objVBC.Type
            Case 1 '標準モジュール
                objVBC.Export (FolderPath & objVBC.Name & ".bas")
            Case 2 'クラスモジュール
                objVBC.Export (FolderPath & objVBC.Name & ".cls")
            Case 3 'フォーム
                objVBC.Export (FolderPath & objVBC.Name & ".frm")
        End Select
    Next objVBC
    
    VB_ExportModule = FolderPath
End Function
Private Function gitPullAddCommitPush()
    Dim strCMD As String
    Call VB_ExportModule
    
    strCMD = "git pull"
    Debug.Print ExecuteCMD(strCMD)
    
    strCMD = "git add ."
    Debug.Print ExecuteCMD(strCMD)
    
    strCMD = "git commit -m 'revise'"
    Debug.Print ExecuteCMD(strCMD)
    
    strCMD = "git push origin master"
    Debug.Print ExecuteCMD(strCMD)
End Function
Public Function VB_GetModuleNames() As String() 'ブックの全モジュール名を取得
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim objVBC As Object 'VBComponent
    Dim buf() As String
    Dim i As Long
    
    i = 0
    ReDim buf(wb.VBProject.VBComponents.Count - 1) As String
    
    For Each objVBC In wb.VBProject.VBComponents
        With objVBC
            If (.Type = 1) Or (.Type = 2) Or (.Type = 3) Then
                buf(i) = .Name
                i = i + 1
            End If
        End With
    Next objVBC
    
    VB_GetModuleName = buf
End Function
Private Function CloneVBProject()  'クローンを作成して新ブックを作る
    Dim FolderPath As String
    'モジュールエクスポートとファイルパス取得
    FolderPath = VB_ExportModule
    
    'モジュール取り込み&新規作成
    Call mdl_init.main(FolderPath)
End Function
Public Function VB_RunFromModule(ModuleFilePath As String, FunctionName As String) As Variant
    'basファイルから関数を読込・実行
    '実行後はコードを残さない
    Dim VBC As Object 'VBComponent
    Dim buf As Variant
    
    'Module読込
    Set VBC = ThisWorkbook.VBProject.VBComponents.Import(ModuleFilePath)
    'Function実行
    buf = Application.Run(VBC.Name & "." & FunctionName)
    'Module開放
    ThisWorkbook.VBProject.VBComponents.Remove VBC
    
    '関数の返り値を渡す
    VB_RunFromModule = buf
End Function
Public Function VB_ImportModule(ModuleFilePath As String) As String 'ファイルパスからbas等をインポート
    Dim wb As Workbook
    Set wb = ThisWorkbook
    On Error GoTo ErrProc
    
    With wb.VBProject.VBComponents
        .Import ModuleFilePath
        VB_ImportModule = .Item(.Count).Name
    End With
    Exit Function
    
ErrProc:
    MsgBox (Err.Description)
End Function

'''参照設定系
Public Function Ref_chk(ReferencesDescription As String) As Boolean '参照設定が設定されているか名前でチェック
    Dim Ref As Object
    For Each Ref In ActiveWorkbook.VBProject.References
        If Ref.Description = ReferencesDescription Then
            Ref_chk = True
            Exit Function
        End If
    Next Ref
    Ref_chk = False
End Function
Public Function Ref_Add(Ref_FilePath As String) '参照設定をフルパスから追加
    ThisWorkbook.VBProject.References.AddFromFile Ref_FilePath
End Function

'''Property
Public Property Let Application_vbaSpeedUp(flg As Boolean) 'VBA高速化メソッドの一括設定
    'HowtoUse:vbaSpeedUp=True|False
    With Application
        .EnableEvents = (Not flg) 'イベント
        .ScreenUpdating = (Not flg)  '描画
        .DisplayAlerts = (Not flg)  'アラート
        
        .Calculation = IIf(flg, xlCalculationManual, xlCalculationAutomatic)  '再計算
    End With
End Property

'''汎用関数Function系
Public Function IsFolderExists(Optional FolderPath As String) As Boolean 'フォルダの存在チェック
    IsFolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(FolderPath)
End Function
Public Function LoadTextFile(FilePath As String) As String() 'テキストファイルを配列に読み込む
    Dim buf() As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(FilePath).OpenAsTextStream
            buf = Split(.ReadAkk, vbCrLf)
            .CIose
        End With
    End With
    LoadTextFile = buf
End Function
Public Function AddWorksheet(wsName As String) 'ワークシートの存在を調べ、なければ追加
    Dim ws As Worksheet
    
    'wsNameという名のワークシートが存在するか調べる｡存在すれば終了｡
    For Each ws In Worksheets
        If ws.Name = wsName Then
            Exit Function
        End If
    Next ws
    
    '存在しない場合は新規作成
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
        .Name = wsName
    End With
End Function
Public Function GetWorksheetsName(wb As Workbook) As String()  'ブックの全シート名を取得
    Dim ws As Worksheet
    Dim buf() As String
    
    ReDim buf(0 To wb.Worksheets.Count - 1)
    
    For Each ws In wb.Worksheets
        buf(ws.Index - 1) = ws.Name
    Next ws
    GetWorksheetsName = buf
End Function
Public Function TransposeArray(argArray As Variant) As Variant '配列転置(空白も空白として転置)
    Dim buf() As Variant
    Dim i As Long, j As Long
    Dim m As Long, n As Long
    
    m = UBound(argArray, 2)
    n = UBound(argArray)
    ReDim buf(m, n)
    
    For i = 0 To m
        For j = 0 To n
            buf(i, j) = argArray(j, i)
        Next
    Next
    TransposeArray = buf
End Function
Public Function OpenWorkBook(FilePath As String, Optional wbPassword As String) As Workbook 'パスのブックを開き、そのブックオブジェクトを返す
    On Error Resume Next
    Dim wb As Workbook
    
    For Each wb In Workbooks '開いているブックの中に指定のブックがある場合、wbオブジエクトにSet
        If wb.FullName = FilePath Then
            wb.Activate
            Set OpenWorkBook = Workbooks(Dir(FilePath)) 'ファイル名の取得
            Exit Function
        End If
    Next
    
    Set OpenWorkBook = Workbooks.Open(FilePath, UpdateLinks:=3, Password:=wbPassword) 'なければブックを開く
End Function
Public Function MeasureTime(FunctionName As String) As Single '実行時間計測
    Dim T0 As Single, T As Single
    '実行前時刻
    T0 = Timer
    
    '関数実行
    Application.Run FunctionName
    
    '実行後時刻と差分
    T = Timer - T0
    Debug.Print (T & "秒")
    d_MeasureTime = T
End Function
Public Function Array2String(argArray As Variant) As String '配列を改行で区切り、1センテンスにする
    'need:GetDimension
    Dim i As Long, j As Long
    Dim d As Long '次元数
    Dim buf As String
    d = GetDimension(argArray) '次元数
    On Error Resume Next  '配列がBase0かBase1によってd-1かdになるのでこの仕様
    
    For j = 0 To d '再起呼び出しでもよいかも
        For i = LBound(argArray, 1) To UBound(argArray, 1)
            If d > 1 Then
                buf = buf & vbCrLf & argArray(i, j)
            ElseIf d = 1 Then
                buf = buf & vbCrLf & argArray(i)
            End If
        Next i
        If d = 1 Then Exit For
    Next j
    Array2String = buf
End Function
Public Function GetDimension(argArray As Variant) As Long '配列の次元数を返す
    Dim i As Long
    Dim tmp As Long
    On Error Resume Next
    i = 1
    Do While Err.Number = 0
        tmp = UBound(argArray, i)
        i = i + 1
    Loop
    GetDimension = i - 2
End Function
Public Function yyyymmdd2Serial(yyyymmdd As Long) As Date 'yyyymmdd→yyyy/mm/ddヘ変換
    yyyymmdd2Serial = DateSerial(Int(yyyymmdd / 10000), Int((yyyymmdd Mod 10000) / 100), yyyymmdd Mod 100)
End Function
Public Function GetIndexOfArray(TargetArray, FindChar As String) As Long 'FindCharのTargetArrayに対するindex番号を取得
    Dim i As Long, j As Long
    For j = 0 To UBound(TargetArray, 2)
        For i = 0 To UBound(TargetArray, 1)
            If TargetArray(i, j).Value = FindChar Then
                GetIndexOfArray = i
                Exit Function
            End If
        Next i
    Next j
End Function
Public Function PasteRange(argArray As Variant, TargetRange As Range) '配列をサイズ調整してセルに貼り付け
    TargetRange.Resize(UBound(argArray, 1), UBound(argArray, 2)) = argArray
End Function
Public Function SliceArray(argArray, Optional IndS As Long, Optional IndE As Long) As Variant() '配列の部分集合をインデックス番号から作る
    Dim L As Long, U As Long
    Dim i As Long, n As Long
    Dim buf() As Variant
    L = LBound(argArray)
    U = UBound(argArray)
    If IndS = "" Then IndS = L
    If IndE = "" Then IndE = U
    IndS = WorksheetFunction.Max(IndS, L)
    IndE = WorksheetFunction.Min(IndE, U)
    n = IndE - IndS + L
    
    ReDim buf(L To n) '新しい配列はargArrayのLboundから始まる(基本的に0or1)
    For i = L To n
        buf(i) = argArray(IndS + i - L)
    Next i
    Set SliceArray = buf
End Function
Public Function ExecuteCMD(strCMD As String) 'CMDにコマンドを投げる
    Dim buf As String
    With CreateObject("Wscript.Shell")
        .CurrentDirectory = ThisWorkbook.Path
        buf = .Run(strCMD, WaitOnReturn:=True)
    End With
End Function

'''汎用関数Sub系
Public Sub BreakLinkOfWorkBook() '対象ワークブックの全リンクを解除する
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim i As Long
    Dim buf As Variant
    buf = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    If buf = "" Then Exit Sub 'リンクがなければ終了
    
    For i = 1 To UBound(buf)
        wb.BreakLink Name:=buf(i), Type:=xlLinkTypeExcelLinks
    Next i
End Sub
Public Sub PasteOnlyValue() 'ワークシートのセルを全て値貼り状態にする。
    Dim ws As Worksheets
    Set ws = ActiveSheet
    
    Dim rngStart As Range, rngLast As Range
    With ws
        Set rngStart = .Range("a1")
        Set rngLast = rngStart.SpecialCeIIs(xILastCell)
    '値貼り
        .Range(rngStart, rngLast).Value = .Range(rngStart, rngLast).Value
    End With
End Sub
