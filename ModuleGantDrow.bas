Attribute VB_Name = "ModuleGantDrow"

Sub testDrow()
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    Dim ii As Integer
    Dim jj As Integer
    Dim row As Integer
    Dim col As Integer
    
    Dim startDate As Date
    Dim startRng As Range
    Dim nowDate As Date
    Dim nowRng As Range
    
    Dim beforeRow As Integer
    
    Call del_shape
    
    Set startRng = WS.Range("R3")
    If (IsDate(startRng)) Then
       startDate = startRng
    Else
       GoTo LEXIT
    End If
    
    Set nowRng = WS.Range("G2")
    If (IsDate(nowRng)) Then
       nowDate = nowRng
    Else
       GoTo LEXIT
    End If
    
    row = 5
    
    For jj = row To 500
       Dim keikaku1StartDate As Date
       Dim keikaku1StartRng As Range
       Dim keikaku1EndDate As Date
       Dim keikaku1EndRng As Range
       
       Dim keikaku2StartDate As Date
       Dim keikaku2StartRng As Range
       Dim keikaku2EndDate As Date
       Dim keikaku2EndRng As Range
       
       Dim percent1 As Integer
       Dim percent2 As Integer
       
       percent1 = Cells(jj, "Q")
       percent2 = Cells(jj + 1, "Q")
       
       ' 開始日を取得
       Set keikaku1StartRng = Cells(jj, "D")
       If (IsDate(keikaku1StartRng)) Then
          keikaku1StartDate = keikaku1StartRng
       Else
          GoTo LLOOPEND
       End If
       
       ' 終了日を取得
       Set keikaku1EndRng = Cells(jj, "E")
       If (IsDate(keikaku1EndRng)) Then
          keikaku1EndDate = keikaku1EndRng
       Else
          GoTo LLOOPEND
       End If
       
       ' 開始日を取得
       Set keikaku2StartRng = Cells(jj + 1, "D")
       If (IsDate(keikaku2StartRng)) Then
          keikaku2StartDate = keikaku2StartRng
       Else
          GoTo LLOOPEND
       End If
       
       ' 終了日を取得
       Set keikaku2EndRng = Cells(jj + 1, "E")
       If (IsDate(keikaku2EndRng)) Then
          keikaku2EndDate = keikaku2EndRng
       Else
          GoTo LLOOPEND
       End If
       
       ' 進捗率をもとに描画するか判定する
       If (percent1 = 0) Then
           keikaku1StartDate = nowRng
           keikaku1EndDate = nowRng
       End If
       
       If (percent2 = 0) Then
           keikaku2StartDate = nowRng
           keikaku2EndDate = nowRng
       End If
       
       If (percent1 = 100) Then
           keikaku1StartDate = nowRng
           keikaku1EndDate = nowRng
       End If
       
       If (percent2 = 100) Then
           keikaku2StartDate = nowRng
           keikaku2EndDate = nowRng
       End If
       
       'If (percent2 = 100) Then
       '   If (beforeRow = 0) Then
       '      beforeRow = jj
       '   End If
       '   GoTo LLOOPEND
       'End If
       
       ' 描画位置変数
       Dim draw1StartRng As Range
       Dim draw1EndRng As Range
       
       Dim draw2StartRng As Range
       Dim draw2EndRng As Range
       
       If (beforeRow > 0) Then
          
          percent1 = Cells(beforeRow, "Q")
       
          ' 開始日を取得
          Set keikaku1StartRng = Cells(beforeRow, "D")
          If (IsDate(keikaku1StartRng)) Then
             keikaku1StartDate = keikaku1StartRng
          Else
             GoTo LLOOPEND
          End If
       
          ' 終了日を取得
          Set keikaku1EndRng = Cells(beforeRow, "E")
          If (IsDate(keikaku1EndRng)) Then
             keikaku1EndDate = keikaku1EndRng
          Else
             GoTo LLOOPEND
          End If
       
          ' 進捗率をもとに描画するか判定する
          If (percent1 = 0) Then
              keikaku1StartDate = nowRng
              keikaku1EndDate = nowRng
          End If
       
          Call addline(msearchDate(keikaku1StartDate, beforeRow), _
                    msearchDate(keikaku1EndDate, beforeRow), _
                    percent1, _
                    msearchDate(keikaku2StartDate, jj + 1), _
                    msearchDate(keikaku2EndDate, jj + 1), _
                    percent2)
          beforeRow = 0
       Else
          Call addline(msearchDate(keikaku1StartDate, jj), _
                    msearchDate(keikaku1EndDate, jj), _
                    percent1, _
                    msearchDate(keikaku2StartDate, jj + 1), _
                    msearchDate(keikaku2EndDate, jj + 1), _
                    percent2)
       End If
       
LLOOPEND:
    Next jj


LEXIT:

End Sub

 Sub del_shape()
    For Each myshape In ActiveSheet.Shapes
        If myshape.Type <> 8 Then 'フォーム(ボタンなど)以外
            myshape.Delete
        End If
    Next
 End Sub
 
Function msearchDate(inputDate As Date, row As Integer) As Range
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    Dim ii As Integer
    Dim jj As Integer
    Dim col As Integer
    Dim day As Integer
    Dim dayS As String
    
    Dim diffDay As Integer
   
    Dim kijyunDate As Date
    
    kijyunDate = WS.Range("R3")
    
    diffDay = DateDiff("d", kijyunDate, inputDate)
    
    day = diffDay + CNumAlp("R")
    dayS = CStr(CNumAlp(day)) & CStr(row)
    Set msearchDate = WS.Range(dayS)
End Function


Function addline(rngStart1 As Range, rngEnd1 As Range, percent1 As Integer, _
                 rngStart2 As Range, rngEnd2 As Range, percent2 As Integer)
    ' Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    
    'Shapeを配置するための基準となるセル
    ' Set rngStart = Range("B2")
    ' Set rngEnd = Range("J2")
    
    'セルのLeft、Top、Widthプロパティを利用して位置決め
    BX = rngStart1.Left + (rngEnd1.Left - rngStart1.Left + rngEnd1.Width) / 100 * percent1
    BY = rngStart1.Top + rngStart1.Height / 2
    EX = rngStart2.Left + (rngEnd2.Left - rngStart2.Left + rngEnd2.Width) / 100 * percent2
    EY = rngEnd2.Top + rngEnd2.Height / 2
    
    '赤色・太さ1.5ポイントの矢印線
    With ActiveSheet.Shapes.addline(BX, BY, EX, EY).Line
        .ForeColor.RGB = RGB(255, 0, 0)
        .Weight = 2
        .EndArrowheadStyle = msoArrowheadNone ' msoArrowheadTriangle
        .DashStyle = msoLineSysDot
    End With
End Function
