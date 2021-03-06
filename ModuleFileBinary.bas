Attribute VB_Name = "ModuleFileBinary"
' 動作確認用の関数
Public Sub PT_WriteBinaryFile()
   Dim outputFileName As String
   Dim result As String
   ' エクセルのあるフォルダのoutput.datを読み込む
   outputFileName = ThisWorkbook.Path & "\" & "output.dat"
   result = WriteBinaryFile(outputFileName, "010203040506070809101112AABBCCDDEEFF")
   ' 書き込み結果はファイルを参照すること
End Sub

' 動作確認用の関数
Public Sub PT_ReadBinaryFile()
   Dim outputFileName As String
   Dim outputHexText As String
   ' エクセルのあるフォルダのoutput.datを読み込む
   outputFileName = ThisWorkbook.Path & "\" & "output.dat"
   outputHexText = ReadBinaryFile(outputFileName)
   ' メッセージボックスに読み込み結果を表示
   MsgBox outputHexText
End Sub


' 下記を参考にして実装した
' 【VBA】Openステートメントでバイナリファイルを読み書きする | やさしいプログラミング備忘録
' http://pg-sample.sagami-ss.net/?eid=9
'********************************************
' バイナリデータをテストファイルに出力
' param strfil 入力ファイル名
' param strHexText 16進数のテキスト（例："010203FF"）
'
' return 成功="1" 失敗="0"
'********************************************
Function WriteBinaryFile(ByVal strfil As String, strHexText As String) As String
    '//バイナリファイルの1バイト毎の入出力にはByte型を用いる
    Dim buff() As Byte
    Dim i As Integer
    Dim fp As Long
    Dim outputLen As Integer
    
    ' 出力サイズ
    outputLen = (Len(strHexText) / 2) - 1
    ' 出力領域
    ReDim buff(0 To outputLen) As Byte
    '書き込みデータをセット
    For i = 0 To outputLen
        buff(i) = HexTextToDec(Mid(strHexText, i * 2 + 1, 2))
    Next

    ' FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile

    ' ファイルが存在する場合は指定アドレスが上書きされるだけのため
    ' 書き込み前にファイルを削除するか中身を一旦クリアする
    Open strfil For Output As #fp
    Close (fp)

    ' ファイルオープン(バイナリ書き込みでオープン、ファイルが存在しない場合は新規作成)
    ' Openステートメントを用いてファイルの入出力を行います
    ' モードに下記のいずれかが指定されていればファイルが存在しない場合、新規作成されます
    ' 追加モード(Append)、バイナリモード(Binary)、出力モード(Output)、ランダムアクセスモード(Random)
    ' ※https://msdn.microsoft.com/ja-jp/library/office/gg264163.aspx
    Open strfil For Binary Access Write As #fp

       '//ファイルに書き込み(ファイル先頭からの書き込みを明示)
       Put #fp, 1, buff

    '//ファイルを閉じる
    Close (fp)
    ' 成功
    WriteBinaryFile = "1"
End Function

 

'********************************************
'テストファイルからバイナリデータを読み込み
' param strfil 入力ファイル名
'
' return ReadBinaryFile 16進数のテキスト（例："010203FF"）
'********************************************
Function ReadBinaryFile(ByVal strfil As String) As String
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim strOutput As String

    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile

    '//ファイルを開く
    Open strfil For Binary As #fp

    '//ファイルサイズ分の読み込み領域を確保して読み込む場合の実装例
    'ReDim buff(FileLen(strfil))
    'Get #fp, 1, buff
    '//実装例ここまで

    ' 出力初期化
    strOutput = ""
    '//ファイルの終端まで指定サイズ(最大16バイト)繰り返し読み込む
    Do While NowLoc < LOF(fp)

        '//最大16バイト分の領域を確保し初期化
        If (LOF(fp) - NowLoc) >= 16 Then
            '//残りのファイルサイズが16バイト以上のとき
            ReDim buff(15)
        Else
            '//最終読み込み時(497バイト～500バイト目)は残りのファイルサイズが16未満
            ReDim buff(LOF(fp) - NowLoc - 1)
        End If

        '//データを読み込み
        Get #fp, , buff

        '//現在位置をを保持する(ループBreak判定用)
        NowLoc = Loc(fp)

        '//出力文字列を生成
        For idx = 0 To UBound(buff)
            strBinary = strBinary + Right("00" & Hex(buff(idx)), 2)
        Next

        '//シートの1列目に結果を表示
        strOutput = strOutput + strBinary
        gyo = gyo + 1
    Loop

    '//ファイルを閉じる
    Close (fp)
    ReadBinaryFile = strOutput
End Function
