Attribute VB_Name = "ModuleDimentionImage"

Public Sub PT_makeImagefromFile()
   Dim exeFile As String
   Dim csvFile As String
   
   exeFile = "C:\QR\ZxingQRWriter.jar"
   csvFile = ThisWorkbook.Path & "\" & "sample_csv.csv"
   
   Call makeImagefromFile(exeFile, csvFile)
End Sub


' Param exeCells セル文字列（QRを作成するjarファイルを指定）（例："A1"）
' Param hani     範囲文字列（QRを作成するCSV作成範囲）（例："D6:E6"）
' Param outCell  QRイメージ出力先のセル文字列（例："F6"Z)
' Param width QRイメージの横幅（例：80）
' Param height QRイメージの高さ（例：80）
' Param excelHeight QRイメージの高さに加算するpixel数（例：20）
Public Sub makeImagefromCell(exeCells As String, hani As String, _
                          Optional outCell As String, _
                          Optional width As Integer = 0, Optional height As Integer = 0, _
                          Optional excelHeight As Integer = 0)
   Dim csvFile As String
   csvFile = "./tmp_qr_input.csv"
   csvFile = GetAbsolutePathNameEx(ThisWorkbook.Path, csvFile)
   ' QR生成の設定ファイルを生成
   Call CSVoutputCell(csvFile, hani)
   ' QR生成
   Call makeImagefromFile(CellStringToCellValue(exeCells), csvFile)
   
   Dim paste As range
   Dim p As Long
   Dim inputCell As String
   Set paste = range(hani)
   p = InStr(hani, ":")
   inputCell = Left(hani, p - 1)
   
   Dim rowCount As Integer
   Dim rowCountStart As Integer
   Dim rowCountEnd As Integer
   rowCountStart = CellStringToCellRowNumber(inputCell)
   rowCountEnd = CellStringToCellRowNumber(Mid(hani, p))
   ' CSV行数（Range範囲の行数）
   rowCount = rowCountEnd - rowCountStart + 1
   
   Dim rowNumber As String
   Dim columnAlpha As String
    
   rowNumber = CellStringToCellRowNumber(inputCell)
   columnAlpha = CellStringToCellColumnAlpha(inputCell)

   Call アクティブシートの画像をすべて削除する
   
   Dim rowEnd As Long
   rowEnd = rowCount - 1
   
   Dim outColumn As String
   Dim outRow As String
   outColumn = CellStringToCellColumnAlpha(outCell)
   outRow = CellStringToCellRowNumber(outCell)
   
   For ii = 0 To rowEnd
      Dim heightImage As Integer
      Dim imageOutCell As String
      imageOutCell = outColumn & str(ii + outRow)
      imageOutCell = Replace(imageOutCell, " ", "")
      ' QR読み込み
      heightImage = ImportPicture(Cells(rowNumber + ii, columnAlpha), imageOutCell, width, height)
      ' エクセルの行高さをイメージの高さとする
      Rows(rowNumber + ii).RowHeight = heightImage + excelHeight
   Next
End Sub

Public Sub makeImagefromFile(exeFile As String, inputCSV As String)

   Dim ExeName As String
   Dim executeCommand As String
   Dim result As String
                    
   On Error GoTo functionErr

   ' 引数チェック
   If exeFile = "" Then
     GoTo functionEnd
   End If
   ' 引数チェック
   If inputCSV = "" Then
     GoTo functionEnd
   End If
   
   inputCSV = GetAbsolutePathNameEx(ThisWorkbook.Path, inputCSV)
   
   executeCommand = "java -jar " & exeFile & " " & _
                    inputCSV
   runShellCommand (executeCommand)
   
   GoTo functionEnd

functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:

End Sub




