Attribute VB_Name = "ModuleAES"

Public Sub aesEncodeCell(exeCell As String, inputCell As String, outCell As String, keyCell As String, _
                   Optional rowCount As Integer = 1, _
                   Optional ivCell As String = "A2", _
                   Optional cypherMode As String = "1", _
                   Optional paddimgMode As String = "1", _
                   Optional keySize As String = "128", _
                   Optional blockSize As String = "128")
   Call aesCell(exeCell, "1", inputCell, outCell, keyCell, rowCount, ivCell, cypherMode, paddimgMode, keySize, blockSize)
End Sub

Public Sub aesDecodeCell(exeCell As String, inputCell As String, outCell As String, keyCell As String, _
                   Optional rowCount As Integer = 1, _
                   Optional ivCell As String = "A2", _
                   Optional cypherMode As String = "1", _
                   Optional paddimgMode As String = "1", _
                   Optional keySize As String = "128", _
                   Optional blockSize As String = "128")
   Call aesCell(exeCell, "0", inputCell, outCell, keyCell, rowCount, ivCell, cypherMode, paddimgMode, keySize, blockSize)
End Sub


   ' flg
   '    0 復号
   '    1 暗号
   ' inputCell  入力セル（例："A1"）
   ' outCell    出力セル（例："A2"）
   ' keyCell    鍵セル（例："A3"）
   ' 以降は省略可能
   ' rowCount   入力セルの下方向にrowCount行繰り返す（default = 指定セルのみ）
   ' ivCell    初期セル(default ALL 0x00)
   ' cypherMode
   '    CBC = 1(default)
   '    ECB = 2
   '    OFB = 3
   '    CFB = 4
   '    CTS = 5
   ' paddingMode
   '    None = 1(default = None)
   '    PKCS7 = 2
   '    Zeros = 3
   '    ANSIX923 = 4
   '    ISO10126 = 5
   ' keySize キーサイズ（default = 128）
   ' blockSize ブロックサイズ（default = 128）
Public Sub aesCell(exeFileCell As String, flg As String, _
                   inputCell As String, _
                   outCell As String, _
                   keyCell As String, _
                   Optional rowCount As Integer = 1, _
                   Optional ivCell As String = "A2", _
                   Optional cypherMode As String = "1", _
                   Optional paddimgMode As String = "1", _
                   Optional keySize As String = "128", _
                   Optional blockSize As String = "128")
   
   Dim rowIn As String
   Dim columnIn As String
   Dim rowOut As String
   Dim columnOut As String
   Dim ii As Long
   Dim rowEnd As Long
   Dim planeHexText As String
   Dim keyHexText As String
   Dim ivHexText As String
   Dim exeText As String
   
   rowIn = CellStringToCellRowNumber(inputCell)
   columnIn = CellStringToCellColumnAlpha(inputCell)
   rowOut = CellStringToCellRowNumber(outCell)
   columnOut = CellStringToCellColumnAlpha(outCell)
   
   keyHexText = CellStringToCellValue(keyCell)
   ivHexText = CellStringToCellValue(ivCell)
   exeText = CellStringToCellValue(exeFileCell)
   
   ' 変換最終行
   rowEnd = rowCount - 1
   
   For ii = 0 To rowEnd
       planeHexText = Cells(ii + rowIn, columnIn).Value
       outputHexText = aes(exeText, flg, planeHexText, keyHexText, ivHexText, cypherMode, paddimgMode, keySize, blockSize)
       Cells(ii + rowOut, columnOut) = outputHexText
   Next ii
End Sub

' 動作確認用の関数
Public Sub PT_aesEncode()
   Dim result As String
   result = aesEncode("D:\aes\aes.exe", "000102030405060708090A0B0C0D0E0F", "000102030405060708090A0B0C0D0E0F", "00000000000000000000000000000000")
   ' メッセージボックスにencode結果を表示
   MsgBox result
End Sub
' 動作確認用の関数
Public Sub PT_aesDecode()
   Dim result As String
   result = aesDecode("D:\aes\aes.exe", "0A940BB5416EF045F1C39458C653EA5A", "000102030405060708090A0B0C0D0E0F", "00000000000000000000000000000000")
   ' メッセージボックスにdecode結果を表示
   MsgBox result
End Sub

Public Function aesEncode(exeFileText As String, inputHexText As String, keyHexText As String, _
                    Optional ivHexText As String = "", _
                    Optional cypherMode As String = "1", _
                    Optional paddimgMode As String = "1", _
                    Optional keySize As String = "128", _
                    Optional blockSize As String = "128") As String
    aesEncode = aes(exeFileText, exeFileText, "1", inputHexText, keyHexText, ivHexText, cypherMode, paddimgMode, keySize, blockSize)

End Function

Public Function aesDecode(exeFileText As String, inputHexText As String, keyHexText As String, _
                    Optional ivHexText As String = "", _
                    Optional cypherMode As String = "1", _
                    Optional paddimgMode As String = "1", _
                    Optional keySize As String = "128", _
                    Optional blockSize As String = "128") As String
    aesDecode = aes(exeFileText, exeFileText, "0", inputHexText, keyHexText, ivHexText, cypherMode, paddimgMode, keySize, blockSize)
End Function

   ' encdecflg
   '    0 復号
   '    1 暗号
   ' inputHexText  入力Hexテキスト
   ' keyHexText   鍵Hexテキスト
   ' 出力ファイル
   ' 以降は省略可能
   ' ivHexText    初期ベクトルHexテキスト(default ALL 0x00)
   ' cypherMode
   '    CBC = 1(default)
   '    ECB = 2
   '    OFB = 3
   '    CFB = 4
   '    CTS = 5
   ' paddingMode
   '    None = 1(default)
   '    PKCS7 = 2
   '    Zeros = 3
   '    ANSIX923 = 4
   '    ISO10126 = 5
   ' keySize キーサイズ（default = 128）
   ' blockSize ブロックサイズ（default = 128）
Public Function aes(exeFileText As String, encdecflg As String, inputHexText As String, keyHexText As String, _
                    Optional ivHexText As String = "00000000000000000000000000000000", _
                    Optional cypherMode As String = "1", _
                    Optional paddimgMode As String = "1", _
                    Optional keySize As String = "128", _
                    Optional blockSize As String = "128") As String
                    
   Dim ExeName As String
   Dim tmpInputDataFile As String
   Dim tmpInputKeyFile As String
   Dim tmpInputIVFile As String
   Dim tmpOutputFile As String
   Dim executeCommand As String
   Dim outputHexText As String
   Dim result As String
                    
   On Error GoTo functionEnd

   ' 引数チェック
   If inputHexText = "" Then
     GoTo functionEnd
   End If
   
   
   tmpInputDataFile = ThisWorkbook.Path & "\" & "aes\tmp_input.dat"
   tmpInputKeyFile = ThisWorkbook.Path & "\" & "aes\tmp_key.dat"
   tmpInputIVFile = ThisWorkbook.Path & "\" & "aes\tmp_iv.dat"
   tmpOutputFile = ThisWorkbook.Path & "\" & "aes\tmp_output.dat"
   
   result = WriteBinaryFile(tmpInputDataFile, inputHexText)
   If Not (keyHexText = "") Then
      result = WriteBinaryFile(tmpInputKeyFile, keyHexText)
   End If
   If Not (ivHexText = "") Then
      result = WriteBinaryFile(tmpInputIVFile, ivHexText)
   End If
   
   ' aes EncDecFlg(0=dec,1=enc) dataFile keyFile OutputFile IVFile cypherMode paddingMode keySize blockSize
   ' encdecflg
   '    0 復号
   '    1 暗号
   ' tmpInputDataFile  入力ファイル
   ' tmpInputKeyFile   鍵ファイル
   ' 出力ファイル
   ' 以降は省略可能
   ' tmpInputIVFile    初期ベクトルファイル
   ' cypherMode
   '    CBC = 1(default)
   '    ECB = 2
   '    OFB = 3
   '    CFB = 4
   '    CTS = 5
   ' paddingMode
   '    None = 1(default)
   '    PKCS7 = 2
   '    Zeros = 3
   '    ANSIX923 = 4
   '    ISO10126 = 5
   ' keySize キーサイズ（default = 128）
   ' blockSize ブロックサイズ（default = 128）
   '
   ExeName = exeFileText ' aesを行うexeファイル名
   executeCommand = ExeName & " " & encdecflg & " " & _
                    tmpInputDataFile & " " & _
                    tmpInputKeyFile & " " & _
                    tmpOutputFile & " " & _
                    tmpInputIVFile & " " & _
                    cypherMode & " " & _
                    paddimgMode & " " & _
                    keySize & " " & _
                    blockSize
   runShellCommand (executeCommand)
   aes = ReadBinaryFile(tmpOutputFile)
   
   GoTo functionEnd

functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:

End Function
