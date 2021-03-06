Attribute VB_Name = "ModuleUtilsFile"
'フォルダの絶対パスとファイルの相対パスを合成して、目的のファイルの絶対パスを取得する関数
'fso.GetAbsolutePathName(fso.BuildPath(basePath, refPath))を汎用化した関数
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
                    Err.Raise 8888, "GetAbsolutePathNameEx", "到達できないパスを指定しています。"
                    GetAbsolutePathNameEx = ""
                    Exit Function
                End If
                retVal = Left(retVal, InStrRev(retVal, "\") - 1)
                rpArr(i) = ""
            Case Else
                retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
                rpArr(i) = ""
        End Select
        '相対パス部分が空欄、.\、..\で終わった時、末尾の\が不足するので補完が必要
        If i = UBound(rpArr) Then
            If RefPath <> "" Then
                If Right(RefPath, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '連続\の消去とネットワークパス対策
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    GetAbsolutePathNameEx = retVal
End Function
