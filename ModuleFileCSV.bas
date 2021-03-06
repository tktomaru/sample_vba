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
    Dim fileName As String  ' CSV ファイル
    Dim csv(0, 1) As String ' CSV に書き込む全データ
    
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
   Call CSV出力(ActiveSheet, fileName, range(hani))
End Sub

Sub CSVoutput(fileName As String, outputData As Variant)
    Dim csv As String  ' CSV に書き込む全データ
    Dim line As String ' 1 行分のデータ
    
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
           ' 行を結合
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
             
    ts.Write (csv) ' 書き込み
    
    ts.Close ' ファイルを閉じる
    
    ' 後始末
    Set ts = Nothing
    Set fso = Nothing
End Sub

Sub CSV出力(ByVal sht As Worksheet, varFile As String, Optional ByVal Selection As range = Nothing)
    Application.DisplayAlerts = False
    
    '■現在選択しているセル情報をrngに格納
    Set Rng = Selection
    
    '■新規ブック作成→rngをA1にコピー→CSV保存→CSV閉じる
    Workbooks.Add
    Rng.Copy ActiveSheet.range("A1")
  
    ActiveWorkbook.SaveAs fileName:=varFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close
    
    Application.DisplayAlerts = True
  
    'MsgBox ("CSV出力しました。" & vbLf & vbLf & varFile)
End Sub

Sub ReadCsvToOutput(fileName As String, ByRef output As Variant)
    Dim result As String
    result = ReadCsvRetString(fileName)

    'CsvToJaggedで行・フィールドに分割してジャグ配列に
    Dim jagArray() As Variant
    Dim csvArray() As Variant
    jagArray = CsvToJagged(result)
    ' Max列数
    Dim maxCol As Long
    maxCol = JaggedMaxColumnCount(jagArray)
    '
    If maxCol < 1 Then
       GoTo functionEnd
    End If
    'JaggedTo2Dでジャグ配列を2次元配列に変換
    Call JaggedTo2D(jagArray, csvArray)
    
    output = csvArray
    
    GoTo functionEnd
functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:
End Sub

'文字コードを自動判別し、全行をCrLf区切りに統一してStringに入れる
'参照設定
'  Microsoft Scripting Runtime
'  Microsoft ActiveX Data Objects x.x Library
'  Microsoft Html Object Library
'https://excel-ubara.com/excelvba5/EXCELVBA268.html

Private Function JaggedMaxColumnCount(jagArray As Variant) As Long
    'ジャグ配列の最大列数取得
    Dim maxCol As Long, v As Variant
    maxCol = 0
    For Each v In jagArray
        If UBound(v) > maxCol Then
            maxCol = UBound(v)
        End If
    Next
    JaggedMaxColumnCount = maxCol
End Function


'ジャグ配列を2次元配列に変換
Private Sub JaggedTo2D(ByRef jagArray() As Variant, _
                       ByRef twoDArray As Variant)
    'ジャグ配列の最大列数取得
    Dim maxCol As Long, v As Variant
    maxCol = 0
    For Each v In jagArray
        If UBound(v) > maxCol Then
            maxCol = UBound(v)
        End If
    Next
  
    'ジャグ配列→2次元配列
    Dim i1 As Long, i2 As Long
    ReDim twoDArray(1 To UBound(jagArray), 1 To maxCol)
    For i1 = 1 To UBound(jagArray)
        For i2 = 1 To UBound(jagArray(i1))
            twoDArray(i1, i2) = jagArray(i1)(i2)
        Next
    Next
End Sub

Private Function CsvToJagged(ByVal strRec As String) As Variant()
    Dim childArray() As Variant 'ジャグ配列の子配列
    Dim lngQuate As Long 'ダブルクォーテーション数
    Dim strCell As String '1フィールド文字列
    Dim blnCrLf As Boolean '改行判定
    Dim i As Long '行位置
    Dim j As Long '列位置
    Dim k As Long
 
    ReDim CsvToJagged(1 To 1) 'ジャグ配列の初期化
    ReDim childArray(1 To 1) 'ジャグ配列の子配列の初期化
  
    i = 1 'シートの1行目から出力
    j = 0 '列位置はputChildArrayでカウントアップ
    lngQuate = 0 'ダブルクォーテーションの数
    strCell = ""
    For k = 1 To Len(strRec)
        Select Case Mid(strRec, k, 1)
            Case vbLf, vbCr '「"」が偶数なら改行、奇数ならただの文字
                If lngQuate Mod 2 = 0 Then
                    blnCrLf = False
                    If k > 1 Then '改行のCrLfはCrで改行判定済なので無視する
                        If Mid(strRec, k - 1, 2) = vbCrLf Then
                            blnCrLf = True
                        End If
                    End If
                    If blnCrLf = False Then
                        Call putChildArray(childArray, j, strCell, lngQuate)
                        'これが改行となる
                        Call putjagArray(CsvToJagged, childArray, _
                                         i, j, lngQuate, strCell)
                    End If
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case ",", vbTab '「"」が偶数なら区切り、奇数ならただの文字
                If lngQuate Mod 2 = 0 Then
                    Call putChildArray(childArray, j, strCell, lngQuate)
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case """" '「"」のカウントをとる
                lngQuate = lngQuate + 1
                strCell = strCell & Mid(strRec, k, 1)
            Case Else
                strCell = strCell & Mid(strRec, k, 1)
        End Select
    Next
  
    '最終行の最終列の処理
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
    If i > UBound(jagArray) Then '常に成立するが一応記述
        ReDim Preserve jagArray(1 To i)
    End If
    jagArray(i) = childArray '子配列をジャグ配列に入れる
    ReDim childArray(1 To 1) '子配列の初期化
    i = i + 1 '列位置
    j = 0 '列位置
    lngQuate = 0 'ダブルクォーテーション数
    strCell = "" '1フィールド文字列
End Sub

'1フィールドごとにセルに出力
Private Sub putChildArray(ByRef childArray() As Variant, _
                          ByRef j As Long, _
                          ByRef strCell As String, _
                          ByRef lngQuate As Long)
    j = j + 1
    '「""」を「"」で置換
    strCell = Replace(strCell, """""", """")
    '前後の「"」を削除
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
'文字コードを自動判別し、全行をCrLf区切りに統一してStringに入れる
Private Function ReadCsvRetString(ByVal strFile As String, _
                         Optional ByVal CharSet As String = "Auto") As String
'　　Dim objFSO As New FileSystemObject
'　　Dim inTS As TextStream
'　　Dim adoSt As New ADODB.Stream
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
            'ADOを使って読込、その後の処理を統一するため全レコードをCrLfで結合
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
            'EUC-JP、UTF-32については未テスト
            MsgBox "文字コードを確認してください。" & vbLf & CharSet
            Stop
    End Select
    Set inTS = Nothing
    Set objFSO = Nothing
    ReadCsvRetString = strRec
End Function

'文字コードの自動判別
Private Function getCharSet(strFileName As String) As String
    Dim bytes() As Byte
    Dim intFileNo As Integer
    ReDim bytes(FileLen(strFileName))
    intFileNo = FreeFile
    Open strFileName For Binary As #intFileNo
    Get #intFileNo, , bytes
    Close intFileNo
  
    'BOMによる判断
    getCharSet = getCharFromBOM(bytes)
  
    'BOMなしをデータの文字コードで判別
    If getCharSet = "" Then
        getCharSet = getCharFromCode(bytes)
    End If
  
    Debug.Print strFileName & " : " & getCharSet
End Function

'BOMによる判断
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

'以下は下記サイトのコードをVBAに移植
'https://dobon.net/vb/dotnet/string/detectcode.html

'BOMなしをデータの文字コードで判別
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
        Case Else '判定できず
            getCharFromCode = "Shift_JIS"
    End Select
End Function
