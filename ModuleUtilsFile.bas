Attribute VB_Name = "ModuleUtilsFile"
'�t�H���_�̐�΃p�X�ƃt�@�C���̑��΃p�X���������āA�ړI�̃t�@�C���̐�΃p�X���擾����֐�
'fso.GetAbsolutePathName(fso.BuildPath(basePath, refPath))��ėp�������֐�
Function GetAbsolutePathNameEx(ByVal basePath As String, ByVal RefPath As String) As String
    Dim i As Long
    
    basePath = Replace(basePath, "/", "\")
    basePath = Left(basePath, Len(basePath) - IIf(Right(basePath, 1) = "\", 1, 0))
    
    RefPath = Replace(RefPath, "/", "\")
    
    Dim retVal As String
    Dim rpArr() As String
    rpArr = Split(RefPath, "\")
    
    For i = LBound(rpArr) To UBound(rpArr)
        Select Case rpArr(i)
            Case "", "."
                If retVal = "" Then retVal = basePath
                rpArr(i) = ""
            Case ".."
                If retVal = "" Then retVal = basePath
                If InStrRev(retVal, "\") = 0 Then
                    Err.Raise 8888, "GetAbsolutePathNameEx", "���B�ł��Ȃ��p�X���w�肵�Ă��܂��B"
                    GetAbsolutePathNameEx = ""
                    Exit Function
                End If
                retVal = Left(retVal, InStrRev(retVal, "\") - 1)
                rpArr(i) = ""
            Case Else
                retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
                rpArr(i) = ""
        End Select
        '���΃p�X�������󗓁A.\�A..\�ŏI��������A������\���s������̂ŕ⊮���K�v
        If i = UBound(rpArr) Then
            If RefPath <> "" Then
                If Right(RefPath, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '�A��\�̏����ƃl�b�g���[�N�p�X�΍�
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    GetAbsolutePathNameEx = retVal
End Function
