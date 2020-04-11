Attribute VB_Name = "mdl_ODBC"
Option Explicit
'''Tables
Public Sub Conections_RefreshODBC() '全ての埋め込みODBCConnectionを更新（終了するまで待機）
    Dim flg As Boolean
    Dim Con As WorkbookConnection
    
    For Each Con In ActiveWorkbook.Connections
        flg = .ODBCConnection.BackgroundQuery
        With Con
            .ODBCConnection.BackgroundQuery = False '更新が終了するまで待機
            .Refresh
            .ODBCConnection.BackgroundQuery = flg
        End With
    Next
End Sub
Private Function Connection_ChangeODBC() 'すべてのODBCConnectionを変更
    Dim Con As WorkbookConnection
    
    For Each Con In ActiveWorkbook.Connections
        With Con
            '.Name= '接続名
            '.ODBCConnection.CommandText='SQL文
            '.ODBCConnection.Connection='コネクション文
            '.ODBCConnection.EnableRefresh=
        End With
    Next
End Function
Private Function ListObjects_ChangeQueryTable() 'すべてのListObjectsについて変更(表示名、パスワード保存、書式など)
    Dim QTName As String
    Dim ws As Worksheet
    Dim lstObj As ListObject
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each lstObj In ws.ListObjects
            With lstObj
                If .SourceType = xlSrcQuery Then 'xlSrcQueryについてのみ処理
                    'QueryTable処理
                    With .QueryTable
                        .BackgroundQuery = True
                        .RefreshStyle = xlOverwriteCells
                        .SavePassword = True
                        .AdjustColumnWidth = False
                        .WorkbookConnection.ODBCConnection.SavePassword = True
                        QTName = .WorkbookConnection.Name '接続名の取得
'                        .RowNumbers = False
'                        .FillAjacentFormulas = True
'                        .PreserveFormatting = True
'                        .RefreshOnFileOpen = False
'                        .SaveData = True
'                        .RefreshPeriod = 0
'                        .PreserveColumnInfo = False
                    End With
                    
                    'ListObjects処理
                    .DisplayName = QTName '表示名
                    .Name = QTName '表示名
                    ws.Name = QTName 'シート名
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

