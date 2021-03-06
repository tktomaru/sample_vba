Attribute VB_Name = "ModuleUtils"
' �����񁩁�16�i�@����2�i�@�̑��ݕϊ�
' https://excel.syogyoumujou.com/memorandum/hex_binary.html
Private i As Long
Private varBinary As Variant
Private colHValue As New Collection '�A�z�z��ACollection�I�u�W�F�N�g�̍쐬
Private lngNu() As Long

Sub UtilsInit() '�{���W���[�����g�p����ۂɂ́A�擪�ł��Ȃ炸�R�[�����邱��
    Dim strData As String
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '������
    For i = 0 To 15 '�A�z�z���varBinary�̊e�l���L�[�Ƃ��āA16�i�@�u0�`F�v�̒l���i�[
        colHValue.Add CStr(Hex$(i)), varBinary(i)
    Next
End Sub


Sub Interconversion_Main() '�����񁩁�16�i�@����2�i�@�̑��ݕϊ�
    Dim strData As String
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '������
    For i = 0 To 15 '�A�z�z���varBinary�̊e�l���L�[�Ƃ��āA16�i�@�u0�`F�v�̒l���i�[
        colHValue.Add CStr(Hex$(i)), varBinary(i)
    Next
'**********************
'������16�i�@��2�i�@
'**********************
    strData = "��]�@��"
    MsgBox strData, , "�������"
    strData = StrToHex(strData)
    MsgBox strData, , "������16�i�@"
    strData = HtoB(strData)
    MsgBox strData, , "16�i�@��2�i�@"
'**********************
'2�i�@��16�i�@��������
'**********************
    strData = BtoH(strData)
    MsgBox strData, , "2�i�@��16�i�@"
    strData = HexToStr(strData)
    MsgBox strData, , "16�i�@��������"
    Erase lngNu
End Sub

Private Function StrToHex(ByVal strData As String) As String '������16�i�@
    Dim strChar As String
    ReDim strHex(1 To Len(strData)) As String
    ReDim lngNu(1 To Len(strData))
    For i = 1 To Len(strData)
        strChar = Mid$(strData, i, 1)
        strHex(i) = Hex$(Asc(strChar))
        lngNu(i) = Len(strHex(i)) '16�i�@�̒l�̌������i�[
    Next
    StrToHex = Join$(strHex, vbNullString)
End Function

Private Function HtoB(ByVal strH As String) As String '16�i�@��2�i�@
    ReDim strHtoB(1 To Len(strH)) As String
    For i = 1 To Len(strH)
        strHtoB(i) = varBinary(val("&h" & Mid$(strH, i, 1)))
    Next
    HtoB = Join$(strHtoB, vbNullString)
End Function

Private Function BtoH(ByVal strB As String) As String '2�i�@��16�i�@
    ReDim strBtoH(1 To Len(strB) / 4) As String
    For i = 1 To Len(strB) / 4 '2�i�@(4bit��)��16�i�@�ɕϊ�
        strBtoH(i) = colHValue.item(Mid$(strB, (i - 1) * 4 + 1, 4))
    Next
    BtoH = Join$(strBtoH, vbNullString)
End Function

Private Function HexToStr(ByVal strData As String) As String '16�i�@��������
    Dim lngLen As Long
    Dim strHex As String
    ReDim strChar(1 To UBound(lngNu)) As String
    lngLen = 1
    For i = 1 To UBound(lngNu)
        strHex = Mid$(strData, lngLen, lngNu(i))
        strChar(i) = Chr$(val("&h" & strHex))
        lngLen = lngLen + lngNu(i)
    Next
    HexToStr = Join$(strChar, vbNullString)
End Function

' 16�i���̃e�L�X�g�i��2�F01 / ��1�FFF�j��Long�ɕϊ�����
Function HexTextToDec(str As String) As Long
    Dim l As Long
    l = val("&H" & str)
    HexTextToDec = l
End Function

' param inputCell "A1"�Ȃǂ̃Z���w�蕶����
' return inputCell��Column�l�i��F"A"�j
Function CellStringToCellColumnAlpha(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    tmp = StrReverse(val(StrReverse(inputCell & "9")))
    rowNumber = Left(tmp, Len(tmp) - 1)
    columnAlpah = Replace(inputCell, rowNumber, "")
    
    CellStringToCellColumnAlpha = columnAlpah

End Function

' param inputCell "A1"�Ȃǂ̃Z���w�蕶����
' return inputCell��Row�l�i��F1�j
Function CellStringToCellRowNumber(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    tmp = StrReverse(val(StrReverse(inputCell & "9")))
    rowNumber = Left(tmp, Len(tmp) - 1)
    
    CellStringToCellRowNumber = rowNumber

End Function


' param inputCell "A1"�Ȃǂ̃Z���w�蕶����
' return inputCell�̒l
Function CellStringToCellValue(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    rowNumber = CellStringToCellRowNumber(inputCell)
    columnAlpah = CellStringToCellColumnAlpha(inputCell)
    
    CellStringToCellValue = Cells(rowNumber, columnAlpah).Value

End Function

