Attribute VB_Name = "ModuleGantCalcWorkday"

' 開始日と作業日数と優先順位から計画を立てる
Sub updateKeikakuDate()
    Dim startDate As Date
    Dim taskDbl As Double
    
    Dim holidayRange As Range
    Dim holidayDate() As Date
    
    Dim youbiInt() As Integer
    Dim youbiString As String
    
    Dim priNString As String
    Dim nameA As String
    Dim tmp As String
    
    Dim personalDate() As Date
    
    Dim ii As Integer
    Dim jj As Integer
    Dim row As Integer
    Dim col As Integer
    
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    Dim priMax As Integer
    
    Set nameC = CreateObject("Scripting.Dictionary")
    Set nameCSum = CreateObject("Scripting.Dictionary")
    
    taskDbl = WS.Range("K5")
    startDate = WS.Range("E2")
  
    row = 5
    priNString = "N"
    nameA = "P"
        
    ' 優先順位の最大値
    priMax = WorksheetFunction.Max(Range("N5", "N500"))
      
    ' 担当の割り出し
    For ii = row To WS.Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(WS.Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
           nameCSum.Add tmp, 0
        End If
    Next ii
    
    ' 優先度順に繰り返し優先順位を求める
    For ii = 1 To priMax
    
       For jj = row To 500
    
          If ("" = WS.Cells(jj, "N")) Then
             GoTo LFEND
          End If
       
          ' 探索中の優先度と一致するか
          If (ii = Cells(jj, priNString)) Then
             ' 設定シートのカラム番号を取得
             For col = CNumAlp("I") To WSConfig.Cells(2, Columns.Count).End(xlToLeft).Column
                If WS.Cells(jj, "P") = WSConfig.Cells(2, col) Then
                   GoTo LEnterName
                End If
            Next col
LEnterName:
          tmp = CStr(Cells(jj, nameA))
          holidayDate = conbertRangeToDateWithout(WSConfig.Range("B3:D500"), tmp)
          personalDate = conbertRangeToDate(WSConfig.Range(CNumAlp(col) & "6:" & CNumAlp(col) & "500"))
          youbiString = WSConfig.Range(CStr(CNumAlp(col)) & "5")
          youbiInt = convertYoubi(youbiString)
    
          ' 名前から現在の作業日数を取得
          Dim task As Double
           
           
          ' 開始日からの日数を算出
          Dim sumtask As Double
           
           task = nameC.item(tmp)
          ' 開始日からの日数を算出（開始日）
          sumtask = calcWorkday(startDate, task, _
                     holidayDate, _
                     youbiInt, _
                     personalDate)
           Cells(jj, "D") = startDate + sumtask
                     
          task = nameC.item(tmp) + WS.Cells(jj, "K")
          Cells(jj, "L") = task
           nameC(tmp) = task
               
          ' 開始日からの日数を算出（終了日）
          sumtask = calcWorkday(startDate, task, _
                     holidayDate, _
                     youbiInt, _
                     personalDate)
                     
           nameCSum(tmp) = sumtask
           ' 開始日以外、かつ、 整数の場合にはちょうどその日にタスクが収まるため、加算をなくす
           If (0 <> sumtask And (sumtask = Int(sumtask))) Then
              sumtask = sumtask - 1
           End If
           Cells(jj, "E") = startDate + sumtask
          End If
       Next jj
LFEND:
    Next ii

End Sub

' 担当の除外日以外の祝日を返す
Function conbertRangeToDateWithout(inputRange As Range, name As String) As Date()
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    
   Dim rngDate As String
   Dim rngName As String
   Dim removeName As String
   Dim num As Integer
   Dim tmpdate As Date
   Dim ret() As Date
   num = 0
   
   Dim r As Long, c As Long
   
   With inputRange
      'For c = 1 To .Column.Count
         For r = 1 To .Rows.Count
           rngDate = .item(r, 1).Address(False, False)
           rngName = .item(r, 3).Address(False, False)
           tmpdate = WSConfig.Range(rngDate)
           removeName = WSConfig.Range(rngName)
           If InStr(removeName, name) = 0 Then
              If IsDate(tmpdate) Then
                 ReDim Preserve ret(num)
                 ret(num) = CDate(tmpdate)
                 num = num + 1
              End If
           End If
         Next r
      'Next c
   End With
   conbertRangeToDateWithout = ret
End Function

' RangeからDate配列に変換する
Function conbertRangeToDate(inputRange As Range) As Date()
   Dim rng As Range
   Dim num As Integer
   Dim ret() As Date
   num = 0
   For Each rng In inputRange
     If IsDate(rng) Then
        ReDim Preserve ret(num)
        ret(num) = CDate(rng)
        num = num + 1
     End If
   Next rng
   conbertRangeToDate = ret
End Function

' 例："月,火"という文字列を与えると、ret=[2,3]というInteger配列で返す
Function convertYoubi(youbi As String) As Integer()
   Dim ret() As Integer
   Dim tmp As Variant
   Dim retNum As Integer
   
   retNum = 0
   tmp = Split(youbi, ",")

'０：vbUseSystemDayOfWeek(PCのOSのシステム時間)
'１：vbSunday (日曜日)
'２：vbMonday (月曜日)
'３：vbTuesday (火曜日)
'４：vbWednesday (水曜日)
'５：vbThursday (木曜日)
'６：vbFriday (金曜日)
'７：vbSaturday (土曜日)

    For ii = LBound(tmp) To UBound(tmp)
       Select Case tmp(ii)
       Case "日"
           ReDim Preserve ret(retNum)
           ret(retNum) = 1
           retNum = retNum + 1
       Case "月"
           ReDim Preserve ret(retNum)
           ret(retNum) = 2
           retNum = retNum + 1
       Case "火"
           ReDim Preserve ret(retNum)
           ret(retNum) = 3
           retNum = retNum + 1
       Case "水"
           ReDim Preserve ret(retNum)
           ret(retNum) = 4
           retNum = retNum + 1
       Case "木"
           ReDim Preserve ret(retNum)
           ret(retNum) = 5
           retNum = retNum + 1
       Case "金"
           ReDim Preserve ret(retNum)
           ret(retNum) = 6
           retNum = retNum + 1
       Case "土"
           ReDim Preserve ret(retNum)
           ret(retNum) = 7
           retNum = retNum + 1
       End Select
    Next ii
    
    convertYoubi = ret
End Function

Sub calcPriority()
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    
    Dim row As Integer
    Dim priA As String
    Dim nameA As String
    Dim outA As String
    Dim nameC As Object
    Set nameC = CreateObject("Scripting.Dictionary")
    Dim tmp As String
    
    Dim ii As Integer
    Dim jj As Integer
    
    priA = "M"
    outA = "N"
    nameA = "P"
    row = 5
    
    ' 担当の割り出し
    For ii = row To Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
        End If
    Next ii
        
    ' 優先度の最大値
    Set rng = Range("M5", "M500")
    priMax = WorksheetFunction.Max(rng)
      
    ' 優先度順に繰り返し優先順位を求める
    For ii = 1 To priMax
       For jj = row To 500
          ' 探索中の優先度と一致するか
          If (ii = Cells(jj, CNumAlp(priA))) Then
             ' 名前から現在の作業日数を取得
             Dim task As Integer
             tmp = CStr(Cells(jj, nameA))
             task = nameC.item(tmp)
             task = task + 1
             nameC(tmp) = task
             Cells(jj, outA) = task
          End If
       Next jj
    Next ii
End Sub

' "R"列より右の休日をピンク色にする
Sub YoubiColor()
    Dim holidayRange As Range
    Dim holidayDate() As Date
    
    Dim youbiInt() As Integer
    Dim youbiString As String
    Dim personalDate() As Date
    
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    Dim ii As Integer
    Dim jj As Integer
    Dim kk As Integer
    Dim row As Integer
    Dim col As Integer
    Dim colMax As Integer
    
    Dim nameA As String
    
    Dim keikakuDate As Date
    
    nameA = "P"
    row = 5
    
    colMax = WS.Cells(3, Columns.Count).End(xlToLeft).Column

   For jj = row To 500
       If ("" = WS.Cells(jj, nameA)) Then
          GoTo LOOPEND
       End If
       
             ' 設定シートのカラム番号を取得
       For col = CNumAlp("I") To WSConfig.Cells(2, Columns.Count).End(xlToLeft).Column
          If WS.Cells(jj, "P") = WSConfig.Cells(2, col) Then
             GoTo LEnterName
          End If
       Next col
LEnterName:
       holidayDate = conbertRangeToDateWithout(WSConfig.Range("B3:D500"), WS.Cells(jj, nameA))
       personalDate = conbertRangeToDate(WSConfig.Range(CNumAlp(col) & "6:" & CNumAlp(col) & "500"))
       youbiString = WSConfig.Range(CStr(CNumAlp(col)) & "5")
       youbiInt = convertYoubi(youbiString)
          
       ' 行の色分け
       For kk = CNumAlp("R") To colMax
         
         keikakuDate = WS.Cells(3, kk)
      
         ' 色指定をクリア（白を指定）
         WS.Cells(jj, kk).Interior.color = RGB(255, 255, 255) ' 背景色
         
         ' 祝日
         If (True = isHoliday(holidayDate, youbiInt, personalDate, keikakuDate)) Then
                WS.Cells(jj, kk).Interior.color = RGB(255, 200, 200) ' 背景色をピンクにする
         End If
         
       Next kk
LOOPEND:
    Next jj
End Sub

' keikakuDate が祝日かどうかを判定する
' @param holidayDate 祝日
' @param youbiInt    個人の休む曜日
' @param personalDate 個人の休日
' @return True=祝日　False=平日
Function isHoliday(holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date, keikakuDate As Date) As Boolean
    Dim WS As Worksheet
    Set WS = Worksheets("工程表")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("工程表Config")
    Dim ii As Integer
    Dim ret As Boolean
    

         ' 祝日
         If (CalcArrayLength(holidayDate) >= 1) Then
         For ii = LBound(holidayDate) To UBound(holidayDate)
           ' 祝日と一致するか？
            If ((keikakuDate) = holidayDate(ii)) Then
               ' 一致していたら
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
   
         ' 個人の休日
         If (CalcArrayLength(personalDate) >= 1) Then
         For ii = LBound(personalDate) To UBound(personalDate)
           ' 個人の休日と一致するか？
            If ((keikakuDate) = personalDate(ii)) Then
               ' 一致していたら
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
   
         ' 個人の曜日
         If (CalcArrayLength(youbiInt) >= 1) Then
         For ii = LBound(youbiInt) To UBound(youbiInt)
           ' 個人の曜日と一致するか？
            If (Weekday(keikakuDate) = youbiInt(ii)) Then
               ' 一致していたら
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
         ret = False
LOOPEND:
    isHoliday = ret
End Function

' @param startDate    開始日
' @param taskDbl      作業日数
' @param holidayDate  祝日
' @param youbiInt     非稼働曜日
' @param personalDate 非稼働日
Function calcWorkday(startDate As Date, _
                     taskDbl As Double, _
                     holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date) As Double

   ' Dim startDayInt As Integer
   ' Dim endDayInt As Integer
   ' 切り捨て
   ' startDayInt = taskDbl
   ' 切り上げ
   ' endDayInt = WorksheetFunction.RoundUp(taskDbl, 0)
   Dim tmpTask  As Integer
   
   tmpTask = 0
   Do
         ' 祝日
         If (True = isHoliday(holidayDate, youbiInt, personalDate, startDate + tmpTask)) Then
            ' 休日ならば稼働最終日を延長する
            taskDbl = taskDbl + 1
            
         End If
         ' 現在日を1日経過させる
         tmpTask = tmpTask + 1
   Loop While tmpTask < taskDbl
LFEND:
   calcWorkday = taskDbl

End Function

