Attribute VB_Name = "ModuleQR"

' Param exeCells �Z��������iQR���쐬����jar�t�@�C�����w��j�i��F"A1"�j
' Param hani     �͈͕�����iQR���쐬����CSV�쐬�͈́j�i��F"D6:E6"�j
' Param outCell  QR�C���[�W�o�͐�̃Z��������i��F"F6"Z)
' Param width QR�C���[�W�̉����i��F80�j
' Param height QR�C���[�W�̍����i��F80�j
' Param excelHeight QR�C���[�W�̍����ɉ��Z����pixel���i��F20�j
Public Sub makeQRfromCell(exeCells As String, hani As String, _
                          Optional outCell As String, _
                          Optional width As Integer = 0, Optional height As Integer = 0, _
                          Optional excelHeight As Integer = 0)
   Call makeImagefromCell(exeCells, hani, outCell, width, height, excelHeight)
End Sub
