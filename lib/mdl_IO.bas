Attribute VB_Name = "mdl_IO"
'VBComponent�n�A�Q�Ɛݒ�n�AProperty�A�ėp�֐��n

'''MEMO
'��{�I�ɂ�Function�A���[�N�V�[�g����Ăяo���R�[�h��Sub
'Sub�ŌĂяo���֐��̈����́A���֐����ActiveBook/Sheet�Ƃ��Ă���
'���O�o�C���f�B���O�̓R�[�f�B���O���y�A���s���x������:New Scripting.FileSystemObject
'���s���o�C���f�B���O�͎Q�Ɛݒ肪�s�v�Ŕz�z���y�FCreateObject("Scripting.FileSystemObject")
'����ADODB�֘A�̂ݎ��O�o�C���f�B���O
Option Explicit
Option Base 0
'''VBComponent�n
Public Function VB_ExportModule() As String '�u�b�N�̂��ׂẴ��W���[�����G�N�X�|�[�g
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Const SaveFolderName As String = "lib"
        
    Dim objVBC As Object 'VBComponent
    Dim FolderPath As String
    Dim flg As Boolean
    
    '�ۑ���t�H���_�`�F�b�N�E�쐬
    If wb.Path = "" Then Exit Function '�t�@�C������x���ۑ�����Ă��Ȃ��ƃp�X�����Ȃ�����
    FolderPath = wb.Path & "/" & SaveFolderName & "/"
    flg = IsFolderExists(FolderPath)
    If flg = False Then MkDir (FolderPath)
    
    'export
    For Each objVBC In wb.VBProject.VBComponents
        Select Case objVBC.Type
            Case 1 '�W�����W���[��
                objVBC.Export (FolderPath & objVBC.Name & ".bas")
            Case 2 '�N���X���W���[��
                objVBC.Export (FolderPath & objVBC.Name & ".cls")
            Case 3 '�t�H�[��
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
Public Function VB_GetModuleNames() As String() '�u�b�N�̑S���W���[�������擾
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
Private Function CloneVBProject()  '�N���[�����쐬���ĐV�u�b�N�����
    Dim FolderPath As String
    '���W���[���G�N�X�|�[�g�ƃt�@�C���p�X�擾
    FolderPath = VB_ExportModule
    
    '���W���[����荞��&�V�K�쐬
    Call mdl_init.main(FolderPath)
End Function
Public Function VB_RunFromModule(ModuleFilePath As String, FunctionName As String) As Variant
    'bas�t�@�C������֐���Ǎ��E���s
    '���s��̓R�[�h���c���Ȃ�
    Dim VBC As Object 'VBComponent
    Dim buf As Variant
    
    'Module�Ǎ�
    Set VBC = ThisWorkbook.VBProject.VBComponents.Import(ModuleFilePath)
    'Function���s
    buf = Application.Run(VBC.Name & "." & FunctionName)
    'Module�J��
    ThisWorkbook.VBProject.VBComponents.Remove VBC
    
    '�֐��̕Ԃ�l��n��
    VB_RunFromModule = buf
End Function
Public Function VB_ImportModule(ModuleFilePath As String) As String '�t�@�C���p�X����bas�����C���|�[�g
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

'''�Q�Ɛݒ�n
Public Function Ref_chk(ReferencesDescription As String) As Boolean '�Q�Ɛݒ肪�ݒ肳��Ă��邩���O�Ń`�F�b�N
    Dim Ref As Object
    For Each Ref In ActiveWorkbook.VBProject.References
        If Ref.Description = ReferencesDescription Then
            Ref_chk = True
            Exit Function
        End If
    Next Ref
    Ref_chk = False
End Function
Public Function Ref_Add(Ref_FilePath As String) '�Q�Ɛݒ���t���p�X����ǉ�
    ThisWorkbook.VBProject.References.AddFromFile Ref_FilePath
End Function

'''Property
Public Property Let Application_vbaSpeedUp(flg As Boolean) 'VBA���������\�b�h�̈ꊇ�ݒ�
    'HowtoUse:vbaSpeedUp=True|False
    With Application
        .EnableEvents = (Not flg) '�C�x���g
        .ScreenUpdating = (Not flg)  '�`��
        .DisplayAlerts = (Not flg)  '�A���[�g
        
        .Calculation = IIf(flg, xlCalculationManual, xlCalculationAutomatic)  '�Čv�Z
    End With
End Property

'''�ėp�֐�Function�n
Public Function IsFolderExists(Optional FolderPath As String) As Boolean '�t�H���_�̑��݃`�F�b�N
    IsFolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(FolderPath)
End Function
Public Function LoadTextFile(FilePath As String) As String() '�e�L�X�g�t�@�C����z��ɓǂݍ���
    Dim buf() As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(FilePath).OpenAsTextStream
            buf = Split(.ReadAkk, vbCrLf)
            .CIose
        End With
    End With
    LoadTextFile = buf
End Function
Public Function AddWorksheet(wsName As String) '���[�N�V�[�g�̑��݂𒲂ׁA�Ȃ���Βǉ�
    Dim ws As Worksheet
    
    'wsName�Ƃ������̃��[�N�V�[�g�����݂��邩���ׂ顑��݂���ΏI���
    For Each ws In Worksheets
        If ws.Name = wsName Then
            Exit Function
        End If
    Next ws
    
    '���݂��Ȃ��ꍇ�͐V�K�쐬
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
        .Name = wsName
    End With
End Function
Public Function GetWorksheetsName(wb As Workbook) As String()  '�u�b�N�̑S�V�[�g�����擾
    Dim ws As Worksheet
    Dim buf() As String
    
    ReDim buf(0 To wb.Worksheets.Count - 1)
    
    For Each ws In wb.Worksheets
        buf(ws.Index - 1) = ws.Name
    Next ws
    GetWorksheetsName = buf
End Function
Public Function TransposeArray(argArray As Variant) As Variant '�z��]�u(�󔒂��󔒂Ƃ��ē]�u)
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
Public Function OpenWorkBook(FilePath As String, Optional wbPassword As String) As Workbook '�p�X�̃u�b�N���J���A���̃u�b�N�I�u�W�F�N�g��Ԃ�
    On Error Resume Next
    Dim wb As Workbook
    
    For Each wb In Workbooks '�J���Ă���u�b�N�̒��Ɏw��̃u�b�N������ꍇ�Awb�I�u�W�G�N�g��Set
        If wb.FullName = FilePath Then
            wb.Activate
            Set OpenWorkBook = Workbooks(Dir(FilePath)) '�t�@�C�����̎擾
            Exit Function
        End If
    Next
    
    Set OpenWorkBook = Workbooks.Open(FilePath, UpdateLinks:=3, Password:=wbPassword) '�Ȃ���΃u�b�N���J��
End Function
Public Function MeasureTime(FunctionName As String) As Single '���s���Ԍv��
    Dim T0 As Single, T As Single
    '���s�O����
    T0 = Timer
    
    '�֐����s
    Application.Run FunctionName
    
    '���s�㎞���ƍ���
    T = Timer - T0
    Debug.Print (T & "�b")
    d_MeasureTime = T
End Function
Public Function Array2String(argArray As Variant) As String '�z������s�ŋ�؂�A1�Z���e���X�ɂ���
    'need:GetDimension
    Dim i As Long, j As Long
    Dim d As Long '������
    Dim buf As String
    d = GetDimension(argArray) '������
    On Error Resume Next  '�z��Base0��Base1�ɂ����d-1��d�ɂȂ�̂ł��̎d�l
    
    For j = 0 To d '�ċN�Ăяo���ł��悢����
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
Public Function GetDimension(argArray As Variant) As Long '�z��̎�������Ԃ�
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
Public Function yyyymmdd2Serial(yyyymmdd As Long) As Date 'yyyymmdd��yyyy/mm/dd�w�ϊ�
    yyyymmdd2Serial = DateSerial(Int(yyyymmdd / 10000), Int((yyyymmdd Mod 10000) / 100), yyyymmdd Mod 100)
End Function
Public Function GetIndexOfArray(TargetArray, FindChar As String) As Long 'FindChar��TargetArray�ɑ΂���index�ԍ����擾
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
Public Function PasteRange(argArray As Variant, TargetRange As Range) '�z����T�C�Y�������ăZ���ɓ\��t��
    TargetRange.Resize(UBound(argArray, 1), UBound(argArray, 2)) = argArray
End Function
Public Function SliceArray(argArray, Optional IndS As Long, Optional IndE As Long) As Variant() '�z��̕����W�����C���f�b�N�X�ԍ�������
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
    
    ReDim buf(L To n) '�V�����z���argArray��Lbound����n�܂�(��{�I��0or1)
    For i = L To n
        buf(i) = argArray(IndS + i - L)
    Next i
    Set SliceArray = buf
End Function
Public Function ExecuteCMD(strCMD As String) 'CMD�ɃR�}���h�𓊂���
    Dim buf As String
    With CreateObject("Wscript.Shell")
        .CurrentDirectory = ThisWorkbook.Path
        buf = .Run(strCMD, WaitOnReturn:=True)
    End With
End Function

'''�ėp�֐�Sub�n
Public Sub BreakLinkOfWorkBook() '�Ώۃ��[�N�u�b�N�̑S�����N����������
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim i As Long
    Dim buf As Variant
    buf = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    If buf = "" Then Exit Sub '�����N���Ȃ���ΏI��
    
    For i = 1 To UBound(buf)
        wb.BreakLink Name:=buf(i), Type:=xlLinkTypeExcelLinks
    Next i
End Sub
Public Sub PasteOnlyValue() '���[�N�V�[�g�̃Z����S�Ēl�\���Ԃɂ���B
    Dim ws As Worksheets
    Set ws = ActiveSheet
    
    Dim rngStart As Range, rngLast As Range
    With ws
        Set rngStart = .Range("a1")
        Set rngLast = rngStart.SpecialCeIIs(xILastCell)
    '�l�\��
        .Range(rngStart, rngLast).Value = .Range(rngStart, rngLast).Value
    End With
End Sub
