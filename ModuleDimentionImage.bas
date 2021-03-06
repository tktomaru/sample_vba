Attribute VB_Name = "ModuleDimentionImage"

Public Sub PT_makeImagefromFile()
   Dim exeFile As String
   Dim csvFile As String
   
   exeFile = "C:\QR\ZxingQRWriter.jar"
   csvFile = ThisWorkbook.Path & "\" & "sample_csv.csv"
   
   Call makeImagefromFile(exeFile, csvFile)
End Sub


' Param exeCells �Z��������iQR���쐬����jar�t�@�C�����w��j�i��F"A1"�j
' Param hani     �͈͕�����iQR���쐬����CSV�쐬�͈́j�i��F"D6:E6"�j
' Param outCell  QR�C���[�W�o�͐�̃Z��������i��F"F6"Z)
' Param width QR�C���[�W�̉����i��F80�j
' Param height QR�C���[�W�̍����i��F80�j
' Param excelHeight QR�C���[�W�̍����ɉ��Z����pixel���i��F20�j
Public Sub makeImagefromCell(exeCells As String, hani As String, _
                          Optional outCell As String, _
                          Optional width As Integer = 0, Optional height As Integer = 0, _
                          Optional excelHeight As Integer = 0)
   Dim csvFile As String
   csvFile = "./tmp_qr_input.csv"
   csvFile = GetAbsolutePathNameEx(ThisWorkbook.Path, csvFile)
   ' QR�����̐ݒ�t�@�C���𐶐�
   Call CSVoutputCell(csvFile, hani)
   ' QR����
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
   ' CSV�s���iRange�͈͂̍s���j
   rowCount = rowCountEnd - rowCountStart + 1
   
   Dim rowNumber As String
   Dim columnAlpha As String
    
   rowNumber = CellStringToCellRowNumber(inputCell)
   columnAlpha = CellStringToCellColumnAlpha(inputCell)

   Call �A�N�e�B�u�V�[�g�̉摜�����ׂč폜����
   
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
      ' QR�ǂݍ���
      heightImage = ImportPicture(Cells(rowNumber + ii, columnAlpha), imageOutCell, width, height)
      ' �G�N�Z���̍s�������C���[�W�̍����Ƃ���
      Rows(rowNumber + ii).RowHeight = heightImage + excelHeight
   Next
End Sub

Public Sub makeImagefromFile(exeFile As String, inputCSV As String)

   Dim ExeName As String
   Dim executeCommand As String
   Dim result As String
                    
   On Error GoTo functionErr

   ' �����`�F�b�N
   If exeFile = "" Then
     GoTo functionEnd
   End If
   ' �����`�F�b�N
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




