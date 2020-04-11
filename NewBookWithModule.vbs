FilePath = Replace(WScript.ScriptFullName, ".vbs", ".xlsm")
BookName = Replace(WScript.ScriptName, ".vbs", ".xlsm")
FunctionName = BookName & "!mdl_init.main"

With WScript.CreateObject("Excel.Application")
    .Visible = True 'TrueならExcelの画面を表示
    .Workbooks.Open FilePath
    .Application.Run FunctionName
    .Workbooks(BookName).Close
End With
