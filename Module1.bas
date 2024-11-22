Attribute VB_Name = "Module1"
Sub ExportAllModulesToDesktop()
    Dim vbComp As Object
    Dim exportPath As String
    
    ' �ۑ���̃p�X���w��i�f�X�N�g�b�v�ɂ��� "VBA" �t�H���_�j
    exportPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\"
    
    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    ' �S�Ẵ��W���[�����G�N�X�|�[�g
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Dim fileType As String
        Select Case vbComp.Type
            Case 1: fileType = "bas" ' �W�����W���[��
            Case 2: fileType = "cls" ' �N���X���W���[��
            Case 3: fileType = "frm" ' ���[�U�[�t�H�[��
            Case Else: fileType = "txt" ' ���̑�
        End Select
        vbComp.Export exportPath & vbComp.Name & "." & fileType
    Next vbComp
    
    MsgBox "VBA�R�[�h���G�N�X�|�[�g���܂���: " & exportPath
End Sub

