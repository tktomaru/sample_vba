Attribute VB_Name = "ModuleFileCSV"
Sub PT_ReadCsv()
    Dim csvFileName As String
    Dim csvArray() As Variant
    
    csvFileName = ThisWorkbook.Path & "\" & "sample_csv.csv"
    Call ReadCsvToOutput(csvFileName, csvArray)
    
    Dim i1 As Long, i2 As Long
    Dim fileName As String
    Dim hexText As String
    For i1 = 1 To UBound(csvArray)
       fileName = csvArray(i1, 1)
       hexText = csvArray(i1, 2)
    Next
    
End Sub

Sub PT_CSVoutput()
    Dim fileName As String  ' CSV �t�@�C��
    Dim csv(0, 1) As String ' CSV �ɏ������ޑS�f�[�^
    
    fileName = ThisWorkbook.Path & "\" & "sample_csv.csv"
    csv(0, 0) = ThisWorkbook.Path & "\" & "output.png"
    csv(0, 1) = "010203"
    Call CSVoutput(fileName, csv)
End Sub


Sub PT_CSVoutputCell()
   Dim fileName As String
   fileName = ThisWorkbook.Path & "\" & "sample_csv.csv"
   Call CSVoutputCell(fileName, "D6:E6")
End Sub

Sub CSVoutputCell(fileName As String, hani As String)
   Call CSV�o��(ActiveSheet, fileName, range(hani))
End Sub

Sub CSVoutput(fileName As String, outputData As Variant)
    Dim csv As String  ' CSV �ɏ������ޑS�f�[�^
    Dim line As String ' 1 �s���̃f�[�^
    
    Dim i1 As Long, i2 As Long
    For i1 = LBound(outputData, 1) To UBound(outputData, 1)
        For i2 = LBound(outputData, 2) To UBound(outputData, 2)
            item = outputData(i1, i2)
            If line = "" Then
                line = item
            Else
                line = line & "," & item
            End If
        Next
           ' �s������
           If csv = "" Then
               csv = line
           Else
               csv = csv & vbCrLf & line
           End If
    Next
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(fileName, ForWriting, True)
             
    ts.Write (csv) ' ��������
    
    ts.Close ' �t�@�C�������
    
    ' ��n��
    Set ts = Nothing
    Set fso = Nothing
End Sub

Sub CSV�o��(ByVal sht As Worksheet, varFile As String, Optional ByVal Selection As range = Nothing)
    Application.DisplayAlerts = False
    
    '�����ݑI�����Ă���Z������rng�Ɋi�[
    Set Rng = Selection
    
    '���V�K�u�b�N�쐬��rng��A1�ɃR�s�[��CSV�ۑ���CSV����
    Workbooks.Add
    Rng.Copy ActiveSheet.range("A1")
  
    ActiveWorkbook.SaveAs fileName:=varFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close
    
    Application.DisplayAlerts = True
  
    'MsgBox ("CSV�o�͂��܂����B" & vbLf & vbLf & varFile)
End Sub

Sub ReadCsvToOutput(fileName As String, ByRef output As Variant)
    Dim result As String
    result = ReadCsvRetString(fileName)

    'CsvToJagged�ōs�E�t�B�[���h�ɕ������ăW���O�z���
    Dim jagArray() As Variant
    Dim csvArray() As Variant
    jagArray = CsvToJagged(result)
    ' Max��
    Dim maxCol As Long
    maxCol = JaggedMaxColumnCount(jagArray)
    '
    If maxCol < 1 Then
       GoTo functionEnd
    End If
    'JaggedTo2D�ŃW���O�z���2�����z��ɕϊ�
    Call JaggedTo2D(jagArray, csvArray)
    
    output = csvArray
    
    GoTo functionEnd
functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:
End Sub

'�����R�[�h���������ʂ��A�S�s��CrLf��؂�ɓ��ꂵ��String�ɓ����
'�Q�Ɛݒ�
'  Microsoft Scripting Runtime
'  Microsoft ActiveX Data Objects x.x Library
'  Microsoft Html Object Library
'https://excel-ubara.com/excelvba5/EXCELVBA268.html

Private Function JaggedMaxColumnCount(jagArray As Variant) As Long
    '�W���O�z��̍ő�񐔎擾
    Dim maxCol As Long, v As Variant
    maxCol = 0
    For Each v In jagArray
        If UBound(v) > maxCol Then
            maxCol = UBound(v)
        End If
    Next
    JaggedMaxColumnCount = maxCol
End Function


'�W���O�z���2�����z��ɕϊ�
Private Sub JaggedTo2D(ByRef jagArray() As Variant, _
                       ByRef twoDArray As Variant)
    '�W���O�z��̍ő�񐔎擾
    Dim maxCol As Long, v As Variant
    maxCol = 0
    For Each v In jagArray
        If UBound(v) > maxCol Then
            maxCol = UBound(v)
        End If
    Next
  
    '�W���O�z��2�����z��
    Dim i1 As Long, i2 As Long
    ReDim twoDArray(1 To UBound(jagArray), 1 To maxCol)
    For i1 = 1 To UBound(jagArray)
        For i2 = 1 To UBound(jagArray(i1))
            twoDArray(i1, i2) = jagArray(i1)(i2)
        Next
    Next
End Sub

Private Function CsvToJagged(ByVal strRec As String) As Variant()
    Dim childArray() As Variant '�W���O�z��̎q�z��
    Dim lngQuate As Long '�_�u���N�H�[�e�[�V������
    Dim strCell As String '1�t�B�[���h������
    Dim blnCrLf As Boolean '���s����
    Dim i As Long '�s�ʒu
    Dim j As Long '��ʒu
    Dim k As Long
 
    ReDim CsvToJagged(1 To 1) '�W���O�z��̏�����
    ReDim childArray(1 To 1) '�W���O�z��̎q�z��̏�����
  
    i = 1 '�V�[�g��1�s�ڂ���o��
    j = 0 '��ʒu��putChildArray�ŃJ�E���g�A�b�v
    lngQuate = 0 '�_�u���N�H�[�e�[�V�����̐�
    strCell = ""
    For k = 1 To Len(strRec)
        Select Case Mid(strRec, k, 1)
            Case vbLf, vbCr '�u"�v�������Ȃ���s�A��Ȃ炽���̕���
                If lngQuate Mod 2 = 0 Then
                    blnCrLf = False
                    If k > 1 Then '���s��CrLf��Cr�ŉ��s����ςȂ̂Ŗ�������
                        If Mid(strRec, k - 1, 2) = vbCrLf Then
                            blnCrLf = True
                        End If
                    End If
                    If blnCrLf = False Then
                        Call putChildArray(childArray, j, strCell, lngQuate)
                        '���ꂪ���s�ƂȂ�
                        Call putjagArray(CsvToJagged, childArray, _
                                         i, j, lngQuate, strCell)
                    End If
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case ",", vbTab '�u"�v�������Ȃ��؂�A��Ȃ炽���̕���
                If lngQuate Mod 2 = 0 Then
                    Call putChildArray(childArray, j, strCell, lngQuate)
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case """" '�u"�v�̃J�E���g���Ƃ�
                lngQuate = lngQuate + 1
                strCell = strCell & Mid(strRec, k, 1)
            Case Else
                strCell = strCell & Mid(strRec, k, 1)
        End Select
    Next
  
    '�ŏI�s�̍ŏI��̏���
    If j > 0 And strCell <> "" Then
        Call putChildArray(childArray, j, strCell, lngQuate)
        Call putjagArray(CsvToJagged, childArray, _
                         i, j, lngQuate, strCell)
    End If
End Function

Private Sub putjagArray(ByRef jagArray() As Variant, _
                        ByRef childArray() As Variant, _
                        ByRef i As Long, _
                        ByRef j As Long, _
                        ByRef lngQuate As Long, _
                        ByRef strCell As String)
    If i > UBound(jagArray) Then '��ɐ������邪�ꉞ�L�q
        ReDim Preserve jagArray(1 To i)
    End If
    jagArray(i) = childArray '�q�z����W���O�z��ɓ����
    ReDim childArray(1 To 1) '�q�z��̏�����
    i = i + 1 '��ʒu
    j = 0 '��ʒu
    lngQuate = 0 '�_�u���N�H�[�e�[�V������
    strCell = "" '1�t�B�[���h������
End Sub

'1�t�B�[���h���ƂɃZ���ɏo��
Private Sub putChildArray(ByRef childArray() As Variant, _
                          ByRef j As Long, _
                          ByRef strCell As String, _
                          ByRef lngQuate As Long)
    j = j + 1
    '�u""�v���u"�v�Œu��
    strCell = Replace(strCell, """""", """")
    '�O��́u"�v���폜
    If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
        If Len(strCell) <= 2 Then
            strCell = ""
        Else
            strCell = Mid(strCell, 2, Len(strCell) - 2)
        End If
    End If
    If j > UBound(childArray) Then
        ReDim Preserve childArray(1 To j)
    End If
    childArray(j) = strCell
    strCell = ""
    lngQuate = 0
End Sub
'�����R�[�h���������ʂ��A�S�s��CrLf��؂�ɓ��ꂵ��String�ɓ����
Private Function ReadCsvRetString(ByVal strFile As String, _
                         Optional ByVal CharSet As String = "Auto") As String
'�@�@Dim objFSO As New FileSystemObject
'�@�@Dim inTS As TextStream
'�@�@Dim adoSt As New ADODB.Stream
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim inTS As Object
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
  
    Dim strRec As String
    Dim i As Long
    Dim aryRec() As String
 
    If CharSet = "Auto" Then CharSet = getCharSet(strFile)
    Select Case UCase(CharSet)
        Case "UTF-8", "UTF-8N"
            'ADO���g���ēǍ��A���̌�̏����𓝈ꂷ�邽�ߑS���R�[�h��CrLf�Ō���
            'Set inTS = objFSO.OpenTextFile(strFile, ForAppending)
            Set inTS = objFSO.OpenTextFile(strFile, 8)
            i = inTS.line - 1
            inTS.Close
            ReDim aryRec(i)
            With adoSt
                '.Type = adTypeText
                .Type = 2
                .CharSet = "UTF-8"
                .Open
                .LoadFromFile strFile
                i = 0
                Do While Not (.EOS)
                    'aryRec(i) = .ReadText(adReadLine)
                    aryRec(i) = .ReadText(-2)
                    i = i + 1
                Loop
                .Close
                strRec = Join(aryRec, vbCrLf)
            End With
        Case "UTF-16 LE", "UTF-16 BE"
            'Set inTS = objFSO.OpenTextFile(strFile, , , TristateTrue)
            Set inTS = objFSO.OpenTextFile(strFile, , , -1)
            strRec = inTS.ReadAll
            inTS.Close
        Case "SHIFT_JIS"
            Set inTS = objFSO.OpenTextFile(strFile)
            strRec = inTS.ReadAll
            inTS.Close
        Case Else
            'EUC-JP�AUTF-32�ɂ��Ă͖��e�X�g
            MsgBox "�����R�[�h���m�F���Ă��������B" & vbLf & CharSet
            Stop
    End Select
    Set inTS = Nothing
    Set objFSO = Nothing
    ReadCsvRetString = strRec
End Function

'�����R�[�h�̎�������
Private Function getCharSet(strFileName As String) As String
    Dim bytes() As Byte
    Dim intFileNo As Integer
    ReDim bytes(FileLen(strFileName))
    intFileNo = FreeFile
    Open strFileName For Binary As #intFileNo
    Get #intFileNo, , bytes
    Close intFileNo
  
    'BOM�ɂ�锻�f
    getCharSet = getCharFromBOM(bytes)
  
    'BOM�Ȃ����f�[�^�̕����R�[�h�Ŕ���
    If getCharSet = "" Then
        getCharSet = getCharFromCode(bytes)
    End If
  
    Debug.Print strFileName & " : " & getCharSet
End Function

'BOM�ɂ�锻�f
Private Function getCharFromBOM(ByRef bytes() As Byte) As String
    getCharFromBOM = ""
    If UBound(bytes) < 3 Then Exit Function
  
    Select Case True
        Case bytes(0) = &HEF And _
             bytes(1) = &HBB And _
             bytes(2) = &HBF
            getCharFromBOM = "UTF-8"
            Exit Function
        Case bytes(0) = &HFF And _
             bytes(1) = &HFE
             If bytes(2) = &H0 And _
                bytes(3) = &H0 Then
                getCharFromBOM = "UTF-32 LE"
                Exit Function
            End If
            getCharFromBOM = "UTF-16 LE"
            Exit Function
        Case bytes(0) = &HFE And _
             bytes(1) = &HFF
            getCharFromBOM = "UTF-16 BE"
            Exit Function
        Case bytes(0) = &H0 And _
             bytes(1) = &H0 And _
             bytes(2) = &HFE And _
             bytes(3) = &HFF
            getCharFromBOM = "UTF-32 BE"
            Exit Function
    End Select
End Function

'�ȉ��͉��L�T�C�g�̃R�[�h��VBA�ɈڐA
'https://dobon.net/vb/dotnet/string/detectcode.html

'BOM�Ȃ����f�[�^�̕����R�[�h�Ŕ���
Private Function getCharFromCode(ByRef bytes() As Byte) As String
    Const bEscape As Byte = &H1B
    Const bAt As Byte = &H40
    Const bDollar As Byte = &H24
    Const bAnd As Byte = &H26
    Const bOpen As Byte = &H28
    Const bB As Byte = &H42
    Const bD As Byte = &H44
    Const bJ As Byte = &H4A
    Const bI As Byte = &H49

    Dim bLen As Long: bLen = UBound(bytes)
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
    Dim isBinary As Boolean: isBinary = False
    Dim i As Integer
  
    For i = 0 To bLen - 1
        b1 = bytes(i)
        If b1 <= &H6 Or b1 = &H7F Or b1 = &HFF Then
            isBinary = True
            If b1 = &H0 And i < bLen - 1 And bytes(i + 1) <= &H7F Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            End If
        End If
    Next
    If isBinary Then
        getCharFromCode = ""
        Exit Function
    End If

    For i = 0 To bLen - 3
        b1 = bytes(i)
        b2 = bytes(i + 1)
        b3 = bytes(i + 2)

        If b1 = bEscape Then
            If b2 = bDollar And b3 = bAt Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bDollar And b3 = bB Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bOpen And (b3 = bB Or b3 = bJ) Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bOpen And b3 = bI Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            End If
            If i < bLen - 3 Then
                b4 = bytes(i + 3)
                If b2 = bDollar And b3 = bOpen And b4 = bD Then
                    getCharFromCode = "Shift_JIS"
                    Exit Function
                End If
                If i < bLen - 5 And _
                    b2 = bAnd And b3 = bAt And b4 = bEscape And _
                    bytes(i + 4) = bDollar And bytes(i + 5) = bB Then
                    getCharFromCode = "Shift_JIS"
                    Exit Function
                End If
            End If
        End If
    Next

    Dim sjis As Integer: sjis = 0
    Dim euc As Integer: euc = 0
    Dim utf8 As Integer: utf8 = 0
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If ((&H81 <= b1 And b1 <= &H9F) Or (&HE0 <= b1 And b1 <= &HFC)) And _
           ((&H40 <= b2 And b2 <= &H7E) Or (&H80 <= b2 And b2 <= &HFC)) Then
            sjis = sjis + 2
            i = i + 1
        End If
    Next
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If ((&HA1 <= b1 And b1 <= &HFE) And _
            (&HA1 <= b2 And b2 <= &HFE)) Or _
            (b1 = &H8E And (&HA1 <= b2 And b2 <= &HDF)) Then
            euc = euc + 2
            i = i + 1
        ElseIf i < bLen - 2 Then
            b3 = bytes(i + 2)
            If b1 = &H8F And (&HA1 <= b2 And b2 <= &HFE) And _
                (&HA1 <= b3 And b3 <= &HFE) Then
                euc = euc + 3
                i = i + 2
            End If
        End If
    Next
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If (&HC0 <= b1 And b1 <= &HDF) And _
            (&H80 <= b2 And b2 <= &HBF) Then
            utf8 = utf8 + 2
            i = i + 1
        ElseIf i < bLen - 2 Then
            b3 = bytes(i + 2)
            If (&HE0 <= b1 And b1 <= &HEF) And _
                (&H80 <= b2 And b2 <= &HBF) And _
                (&H80 <= b3 And b3 <= &HBF) Then
                utf8 = utf8 + 3
                i = i + 2
            End If
        End If
    Next
  
    Select Case True
        Case euc > sjis And euc > utf8
            getCharFromCode = "EUC-JP"
        Case utf8 > euc And utf8 > sjis
            getCharFromCode = "UTF-8N"
        Case sjis > euc And sjis > utf8
            getCharFromCode = "SHIFT-JIS"
        Case Else '����ł���
            getCharFromCode = "Shift_JIS"
    End Select
End Function
