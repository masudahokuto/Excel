Attribute VB_Name = "Module1"
Sub ExportAllModulesToDesktop()
    Dim vbComp As Object
    Dim exportPath As String
    
    ' 保存先のパスを指定（デスクトップにある "VBA" フォルダ）
    exportPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\"
    
    ' フォルダが存在しない場合は作成
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    ' 全てのモジュールをエクスポート
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Dim fileType As String
        Select Case vbComp.Type
            Case 1: fileType = "bas" ' 標準モジュール
            Case 2: fileType = "cls" ' クラスモジュール
            Case 3: fileType = "frm" ' ユーザーフォーム
            Case Else: fileType = "txt" ' その他
        End Select
        vbComp.Export exportPath & vbComp.Name & "." & fileType
    Next vbComp
    
    MsgBox "VBAコードをエクスポートしました: " & exportPath
End Sub

