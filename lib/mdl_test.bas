Attribute VB_Name = "mdl_test"
Option Explicit
Option Base 0
Option Private Module
Private Function test_IO_Ref_chk()
    Const Ref_Path As String = "Microsoft ActiveX Data Objects 2.6 Library"
    Debug.Print mdl_IO.Ref_chk(Ref_Path)
End Function
Private Function test_IO_Ref_Add()
    Const Ref_FilePath As String = "C:\Program Files\Common Files\System\ado\msado26.tlb"
    Debug.Print mdl_IO.Ref_Add(Ref_FilePath)
End Function
Private Function test_IO_Application_vbaSpeedUp()
    Application_vbaSpeedUp = True
    Debug.Print Application.ScreenUpdating
    Application_vbaSpeedUp = False
End Function
Private Function test_VB_ImportModule()
    Dim ModuleName As String
    ModuleName = mdl_IO.VB_ImportModule(ThisWorkbook.Path & "\lib\mdl_init.bas") '実行
    
    Stop
    'インポートしたものを削除
    With ThisWorkbook.VBProject
        .VBComponents.Remove .VBComponents(ModuleName)
    End With
End Function
Private Function test_VB_RunFromModule()
    Debug.Print mdl_IO.VB_RunFromModule(ThisWorkbook.Path & "\lib\mdl_init.bas", "main") '実行
End Function
Private Function test_ExecuteCMD()
    Debug.Print mdl_IO.ExecuteCMD(ThisWorkbook.Path & "\lib\mdl_init.bas")
End Function

