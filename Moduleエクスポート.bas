Attribute VB_Name = "�G�N�X�|�[�g"
Sub Export()
    Set wb = ThisWorkbook
    Dim sh As Worksheet
    Set sh = wb.Sheets(C_WBS_SHNM)
    
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename(InitialFileName:="WBS_" & Format(Now, "yyyymmdd") & ".xlsx", FileFilter:="Excel�t�@�C�� (*.xlsx),*.xlsx*")
    If fileName = False Then
        MsgBox "�ۑ��Ɏ��s���܂����B�������I�����܂��B"
        Exit Sub
    Else
        Dim tmpSh As Worksheet
        Dim tmpShName As String: tmpShName = "WBS_" & Format(Now, "yyyymmdd")
        
        ' WBS��ʃV�[�g�ɃR�s�[
        sh.Copy After:=wb.Sheets(C_WBS_SHNM)
        Set tmpSh = ActiveSheet
        tmpSh.Name = tmpShName
        
        ' �}�N���{�^���폜
        Dim Btn As Object
        For Each Btn In tmpSh.Buttons
            Btn.Delete
        Next Btn
        
        ' �V�u�b�N�ɃV�[�g�R�s�[
        tmpSh.Copy
        ' �V�[�g�폜
        Application.DisplayAlerts = False
        tmpSh.Delete
        Application.DisplayAlerts = True
        
        ' �w�肵���t�@�C�����ŕۑ�
        ActiveWorkbook.SaveAs fileName
        ActiveWorkbook.Close
        
        ' ����WBS�Ƀt�H�[�J�X�����킹��
        sh.Select
        
        MsgBox "�G�N�X�|�[�g���������܂����B"
        
        Set tmpSh = Nothing
    End If
    
    Set sh = Nothing
    Set wb = Nothing

End Sub
