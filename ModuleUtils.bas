Attribute VB_Name = "ModuleUtils"
' 文字列←→16進法←→2進法の相互変換
' https://excel.syogyoumujou.com/memorandum/hex_binary.html
Private i As Long
Private varBinary As Variant
Private colHValue As New Collection '連想配列、Collectionオブジェクトの作成
Private lngNu() As Long

Sub UtilsInit() '本モジュールを使用する際には、先頭でかならずコールすること
    Dim strData As String
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '初期化
    For i = 0 To 15 '連想配列にvarBinaryの各値をキーとして、16進法「0～F」の値を格納
        colHValue.Add CStr(Hex$(i)), varBinary(i)
    Next
End Sub


Sub Interconversion_Main() '文字列←→16進法←→2進法の相互変換
    Dim strData As String
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '初期化
    For i = 0 To 15 '連想配列にvarBinaryの各値をキーとして、16進法「0～F」の値を格納
        colHValue.Add CStr(Hex$(i)), varBinary(i)
    Next
'**********************
'文字列→16進法→2進法
'**********************
    strData = "坂江　保"
    MsgBox strData, , "基準文字列"
    strData = StrToHex(strData)
    MsgBox strData, , "文字列→16進法"
    strData = HtoB(strData)
    MsgBox strData, , "16進法→2進法"
'**********************
'2進法→16進法→文字列
'**********************
    strData = BtoH(strData)
    MsgBox strData, , "2進法→16進法"
    strData = HexToStr(strData)
    MsgBox strData, , "16進法→文字列"
    Erase lngNu
End Sub

Private Function StrToHex(ByVal strData As String) As String '文字列→16進法
    Dim strChar As String
    ReDim strHex(1 To Len(strData)) As String
    ReDim lngNu(1 To Len(strData))
    For i = 1 To Len(strData)
        strChar = Mid$(strData, i, 1)
        strHex(i) = Hex$(Asc(strChar))
        lngNu(i) = Len(strHex(i)) '16進法の値の桁数を格納
    Next
    StrToHex = Join$(strHex, vbNullString)
End Function

Private Function HtoB(ByVal strH As String) As String '16進法→2進法
    ReDim strHtoB(1 To Len(strH)) As String
    For i = 1 To Len(strH)
        strHtoB(i) = varBinary(val("&h" & Mid$(strH, i, 1)))
    Next
    HtoB = Join$(strHtoB, vbNullString)
End Function

Private Function BtoH(ByVal strB As String) As String '2進法→16進法
    ReDim strBtoH(1 To Len(strB) / 4) As String
    For i = 1 To Len(strB) / 4 '2進法(4bit分)を16進法に変換
        strBtoH(i) = colHValue.item(Mid$(strB, (i - 1) * 4 + 1, 4))
    Next
    BtoH = Join$(strBtoH, vbNullString)
End Function

Private Function HexToStr(ByVal strData As String) As String '16進法→文字列
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

' 16進数のテキスト（例2：01 / 例1：FF）をLongに変換する
Function HexTextToDec(str As String) As Long
    Dim l As Long
    l = val("&H" & str)
    HexTextToDec = l
End Function

' param inputCell "A1"などのセル指定文字列
' return inputCellのColumn値（例："A"）
Function CellStringToCellColumnAlpha(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    tmp = StrReverse(val(StrReverse(inputCell & "9")))
    rowNumber = Left(tmp, Len(tmp) - 1)
    columnAlpah = Replace(inputCell, rowNumber, "")
    
    CellStringToCellColumnAlpha = columnAlpah

End Function

' param inputCell "A1"などのセル指定文字列
' return inputCellのRow値（例：1）
Function CellStringToCellRowNumber(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    tmp = StrReverse(val(StrReverse(inputCell & "9")))
    rowNumber = Left(tmp, Len(tmp) - 1)
    
    CellStringToCellRowNumber = rowNumber

End Function


' param inputCell "A1"などのセル指定文字列
' return inputCellの値
Function CellStringToCellValue(inputCell As String) As String
    Dim tmp As String
    Dim rowNumber As String
    Dim columnAlpah As String
    
    rowNumber = CellStringToCellRowNumber(inputCell)
    columnAlpah = CellStringToCellColumnAlpha(inputCell)
    
    CellStringToCellValue = Cells(rowNumber, columnAlpah).Value

End Function

