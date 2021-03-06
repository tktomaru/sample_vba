Attribute VB_Name = "ModuleQR"

' Param exeCells セル文字列（QRを作成するjarファイルを指定）（例："A1"）
' Param hani     範囲文字列（QRを作成するCSV作成範囲）（例："D6:E6"）
' Param outCell  QRイメージ出力先のセル文字列（例："F6"Z)
' Param width QRイメージの横幅（例：80）
' Param height QRイメージの高さ（例：80）
' Param excelHeight QRイメージの高さに加算するpixel数（例：20）
Public Sub makeQRfromCell(exeCells As String, hani As String, _
                          Optional outCell As String, _
                          Optional width As Integer = 0, Optional height As Integer = 0, _
                          Optional excelHeight As Integer = 0)
   Call makeImagefromCell(exeCells, hani, outCell, width, height, excelHeight)
End Sub
