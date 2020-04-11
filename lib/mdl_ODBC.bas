Attribute VB_Name = "mdl_ODBC"
Option Explicit
'''Tables
Public Sub Conections_RefreshODBC() '�S�Ă̖��ߍ���ODBCConnection���X�V�i�I������܂őҋ@�j
    Dim flg As Boolean
    Dim Con As WorkbookConnection
    
    For Each Con In ActiveWorkbook.Connections
        flg = .ODBCConnection.BackgroundQuery
        With Con
            .ODBCConnection.BackgroundQuery = False '�X�V���I������܂őҋ@
            .Refresh
            .ODBCConnection.BackgroundQuery = flg
        End With
    Next
End Sub
Private Function Connection_ChangeODBC() '���ׂĂ�ODBCConnection��ύX
    Dim Con As WorkbookConnection
    
    For Each Con In ActiveWorkbook.Connections
        With Con
            '.Name= '�ڑ���
            '.ODBCConnection.CommandText='SQL��
            '.ODBCConnection.Connection='�R�l�N�V������
            '.ODBCConnection.EnableRefresh=
        End With
    Next
End Function
Private Function ListObjects_ChangeQueryTable() '���ׂĂ�ListObjects�ɂ��ĕύX(�\�����A�p�X���[�h�ۑ��A�����Ȃ�)
    Dim QTName As String
    Dim ws As Worksheet
    Dim lstObj As ListObject
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each lstObj In ws.ListObjects
            With lstObj
                If .SourceType = xlSrcQuery Then 'xlSrcQuery�ɂ��Ă̂ݏ���
                    'QueryTable����
                    With .QueryTable
                        .BackgroundQuery = True
                        .RefreshStyle = xlOverwriteCells
                        .SavePassword = True
                        .AdjustColumnWidth = False
                        .WorkbookConnection.ODBCConnection.SavePassword = True
                        QTName = .WorkbookConnection.Name '�ڑ����̎擾
'                        .RowNumbers = False
'                        .FillAjacentFormulas = True
'                        .PreserveFormatting = True
'                        .RefreshOnFileOpen = False
'                        .SaveData = True
'                        .RefreshPeriod = 0
'                        .PreserveColumnInfo = False
                    End With
                    
                    'ListObjects����
                    .DisplayName = QTName '�\����
                    .Name = QTName '�\����
                    ws.Name = QTName '�V�[�g��
                End If
            End With
        Next lstObj
    Next ws
End Function

Public Function CreateWebQuery(URL As String, QryName As String)
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add()
        ws.Name = QryName
        
        With ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("$A$1"))
        .Name = QryName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Function

