Attribute VB_Name = "WBS作成"
Option Explicit

Dim C_NOW_COLOR As Variant
Dim C_SATURDAY_COLOR As Variant
Dim C_SUNDAY_COLOR As Variant
Dim C_NOWORKDAY_COLOR As Variant
Dim C_WBSLINE_COLOR As Variant
Dim C_WHITE_COLOR As Variant
Dim C_BORDER_COLOR_RED As Variant
Dim C_BORDER_COLOR_BLUE As Variant
Dim C_STATUS_COLOR_NOTSTART As Variant
Dim C_STATUS_COLOR_PROGRESS As Variant
Dim C_STATUS_COLOR_DONE As Variant

Dim wb As Workbook
Dim sh As Worksheet
Dim shConf As Worksheet

Dim minDate, maxDate As Date
Dim maxRow As Long

Sub Init()
    Set wb = ThisWorkbook
    Set sh = wb.Sheets(C_WBS_SHNM)
    Set shConf = wb.Sheets("config")
    
    C_NOW_COLOR = RGB(255, 204, 204)
    C_SATURDAY_COLOR = RGB(183, 222, 232)
    C_SUNDAY_COLOR = RGB(242, 220, 219)
    C_NOWORKDAY_COLOR = RGB(217, 217, 217)
    C_WBSLINE_COLOR = RGB(128, 128, 128)
    C_WHITE_COLOR = RGB(255, 255, 255)
    C_BORDER_COLOR_RED = RGB(255, 0, 0)
    C_BORDER_COLOR_BLUE = RGB(0, 0, 255)
    C_STATUS_COLOR_NOTSTART = RGB(253, 233, 217)
    C_STATUS_COLOR_PROGRESS = RGB(218, 238, 243)
    C_STATUS_COLOR_DONE = RGB(235, 241, 222)
    
    ' 最下行取得
    maxRow = sh.Cells(Rows.Count, ConvertNumAlp(C_NO_COL)).End(xlUp).Row

    ' WBSクリア
    Dim k As Long
    For k = C_HEADER_ROW + 1 To maxRow
        ' 開始予定日のフォント色取得
        If sh.Range(C_STARTPLAN_COL & k).Font.FontStyle <> C_FONT_BOLD Then
            sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).ClearContents
        ElseIf sh.Range(C_STARTPLAN_COL & k).Value = "" And _
            IsNull(sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).Font.FontStyle) = True Then
            ' 文字の太さを戻す（戻し忘れ防止）
            sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).Font.FontStyle = C_FONT_REGULAR
        Else
            sh.Range(C_ENDPLAN_COL & k).ClearContents
        End If
    Next
    ' 年：縮小して全体表示
    sh.Range(C_STARTWBS_COL & ":" & "XFD").Clear
    sh.Range(C_STARTWBS_COL & C_YEAR_ROW & ":" & "XFD" & C_YEAR_ROW).ShrinkToFit = True
    
End Sub

Sub CreateWBS()
    ' 初期化
    Call Init
    
    '========================================================================================
    ' カレンダー作成
    '========================================================================================
    
    ' 開始予定日MIN取得
    minDate = sh.Cells(C_MONTH_ROW, C_PJSTARTDAY_COL).Value
    If IsDate(minDate) = False Then
        MsgBox "開始予定日検索でエラーとなりました。処理を終了します。"
        Exit Sub
    End If
    
    ' 完了予定日MAX取得
    maxDate = sh.Cells(C_MONTH_ROW, C_PJENDDAY_COL).Value
    If IsDate(maxDate) = False Then
        MsgBox "完了予定日検索でエラーとなりました。処理を終了します。"
        Exit Sub
    End If
    
    ' 開始日MAXの月初～最終日MAXの月末までチャートを作る
    Dim startWbsColNum As Long
    startWbsColNum = ConvertNumAlp(C_STARTWBS_COL)
    
    ' 開始予定日MINの月と完了予定日MAXの月の差分月数を取得
    Dim monthCount As Long
    monthCount = DateDiff("m", minDate, maxDate)
    
    Dim curColNum As Long: curColNum = 0
    Dim i As Integer
    For i = 0 To monthCount
        ' 対象月取得
        Dim targetDate As Date
        targetDate = DateAdd("m", i, minDate)
        
        ' 月の日数を取得
        Dim lastDate As Date
        Dim lastDay As Integer
        lastDate = DateSerial(Year(targetDate), Month(targetDate) + 1, 0)
        lastDay = Format(lastDate, "d")
        
        Dim j As Integer
        For j = 1 To lastDay
            If j = 1 Then
                sh.Cells(C_MONTH_ROW, startWbsColNum + curColNum).Value = Month(targetDate)
                ' 年を記入
                If i = 0 Or Month(targetDate) = 1 Then
                    sh.Cells(C_YEAR_ROW, startWbsColNum + curColNum).Value = Year(targetDate)
                End If
            End If
            
            ' 日を記入
            sh.Cells(C_DAY_ROW, startWbsColNum + curColNum).Value = j
            
            Dim curDate As Date
            curDate = DateSerial(Year(targetDate), Month(targetDate), j)
            
            Dim curCol As String
            curCol = ConvertNumAlp(startWbsColNum + curColNum)
            
            With sh.Cells(C_HEADER_ROW, startWbsColNum + curColNum)
                ' 曜日（数値）を日本語に変換
                .Value = ConvWeekDayJp(curDate)
                Select Case Weekday(curDate)
                Case vbSunday
                    .Interior.Color = C_SUNDAY_COLOR
                    sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
                Case vbSaturday
                    .Interior.Color = C_SATURDAY_COLOR
                    sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
                End Select
            End With
            
            ' 土日祝日判定
            Dim findRng1, findRng2 As Range
            Set findRng1 = shConf.Range(C_HOLIDAY_COL & ":" & C_HOLIDAY_COL).Find(curDate, LookAt:=xlWhole)
            If Not findRng1 Is Nothing Then
                ' 土日祝日の場合は列をグレーにする
                sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
            End If
            ' 非稼働日判定
            Set findRng2 = shConf.Range(C_NOWORKDAY_COL & ":" & C_NOWORKDAY_COL).Find(curDate, LookAt:=xlWhole)
            If Not findRng2 Is Nothing Then
                ' 非稼働日の場合は列をグレーにする
                sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
            End If
            
            ' 担当者別の非稼働曜日、日のセルをグレーにする
            Call SetChargeNoWorkDay("", curDate, ConvertNumAlp(curCol))
            
            curColNum = curColNum + 1
            
            Set findRng1 = Nothing
            Set findRng2 = Nothing
        Next
    Next
    
    ' 今日の日付列に色を付ける
    Dim nowRng As Range
    Set nowRng = GetDayRange(Date)
    If Not nowRng Is Nothing Then
        Dim nowRngCol As String: nowRngCol = ConvertNumAlp(nowRng.Column)
        sh.Range(nowRngCol & C_DAY_ROW & ":" & nowRngCol & maxRow).Interior.Color = C_NOW_COLOR
        nowRng.Select
    End If
    
    ' 最も右の列（アルファベット）を取得
    Dim maxcolAlp As String
    maxcolAlp = ConvertNumAlp(startWbsColNum + curColNum - 1)
    
    ' 罫線を引く
    With sh.Range(C_STARTWBS_COL & C_DAY_ROW & ":" & maxcolAlp & maxRow).Borders
        .LineStyle = xlContinuous
        .Color = C_WBSLINE_COLOR
    End With
    sh.Range(C_STARTWBS_COL & C_MONTH_ROW & ":" & maxcolAlp & maxRow).BorderAround Color:=C_WBSLINE_COLOR
    
    With sh.Range(C_STARTWBS_COL & ":" & maxcolAlp)
        ' 列幅調整
        .EntireColumn.ColumnWidth = 2.45
        ' 中央揃え
        .HorizontalAlignment = xlCenter
        ' フォント設定
        .Font.Name = "メイリオ"
        .Font.Size = "9"
    End With
        
    '========================================================================================
    ' WBS作成
    '========================================================================================

    ' 開始予定日取得
    Dim startPlanDate As Date
        
    ' グループ別にガントチャートを引く
    Dim maxGroup As Long
    maxGroup = Application.WorksheetFunction.Max(sh.Range(C_GROUP_COL & (C_HEADER_ROW + 1) & _
                ":" & C_GROUP_COL & maxRow))
    
    Dim chargeName As String
    Dim chargeColor As Variant
    Dim chargeNameColNum As Long
    chargeNameColNum = ConvertNumAlp(C_CONF_CHARGENAME_COL)
    ' 担当者マスタ最下行取得
    Dim maxChargeRow As Long
    maxChargeRow = shConf.Cells(Rows.Count, chargeNameColNum).End(xlUp).Row
    
    Dim g As Integer
    For g = 1 To maxGroup
        ' WBSシートから担当者名を検索
        Dim findGroupStRng, findGroupEdRng As Range
        ' 上から検索
        Set findGroupStRng = sh.Range(C_GROUP_COL & C_HEADER_ROW & ":" & _
                                C_GROUP_COL & maxRow).Find(g, LookAt:=xlWhole)
        ' 下から検索
        Set findGroupEdRng = sh.Range(C_GROUP_COL & C_HEADER_ROW & ":" & _
                                C_GROUP_COL & maxRow).Find(g, LookAt:=xlWhole, SearchDirection:=xlPrevious)

        ' 検索結果がNothingの場合は次ループへ
        If findGroupStRng Is Nothing Or findGroupEdRng Is Nothing Then
            GoTo Continue
        End If

        ' 担当者名を取得
        chargeName = sh.Range(C_CHARGE_COL & findGroupStRng.Row).Value
        
        ' 担当者のタスク末端日付を取得
        startPlanDate = GetChargeEndDate(chargeName)
        
        Dim h As Long
        For h = findGroupStRng.Row To findGroupEdRng.Row
            ' グループ番号を取得
            Dim orderCellVal As String
            orderCellVal = sh.Range(C_GROUP_COL & h).Value
            
            ' 現在のグループ番号と処理対象のグループ番号が異なる場合は次ループへ
            If g <> orderCellVal Then GoTo ContinueGroup
            
            ' ステータスを取得
            Dim statusVal As String
            statusVal = sh.Range(C_STATUS_COL & h).Value
            Select Case statusVal
            Case C_STATUS_NOTSTART
                sh.Range(C_STATUS_COL & h).Interior.Color = C_STATUS_COLOR_NOTSTART
            Case C_STATUS_PROGRESS
                sh.Range(C_STATUS_COL & h).Interior.Color = C_STATUS_COLOR_PROGRESS
            Case C_STATUS_DONE
                sh.Range(C_STATUS_COL & h).Interior.Color = C_STATUS_COLOR_DONE
            Case Else
                sh.Range(C_STATUS_COL & h).Interior.ColorIndex = 0
            End Select
                        
            ' 予定工数を取得
            Dim manHourPlanVal As String
            manHourPlanVal = sh.Range(C_MANHOUR_COL & h).Value
            ' 担当者名を取得
            chargeName = sh.Range(C_CHARGE_COL & h).Value
            ' 担当者マスタから担当者の色を取得
            Dim tmpChargeRng As Range
            Dim chargeEndCol As String
            ' 担当者マスタの列末端を取得
            chargeEndCol = shConf.Cells(C_CONFHEADER_ROW, Columns.Count).End(xlToLeft).Column
            Set tmpChargeRng = shConf.Range(C_CONF_CHARGENAME_COL & C_CONFHEADER_ROW & ":" & _
                                    ConvertNumAlp(CInt(chargeEndCol)) & C_CONFHEADER_ROW).Find(chargeName, LookAt:=xlWhole)
            If Not tmpChargeRng Is Nothing Then
                ' 担当者の色を取得
                chargeColor = GetRGBValue(shConf.Range(ConvertNumAlp(tmpChargeRng.Column) & C_CONF_CHARGECLR_ROW).Interior.Color)
                
                ' グループ単位でガントチャートを引く
                ' 予定工数とグループ番号の数値入力、開始予定日に日付が入っていない行のみ
                If IsNumeric(manHourPlanVal) = True And IsNumeric(orderCellVal) = True Then
                    If sh.Range(C_STARTPLAN_COL & h).Font.FontStyle = C_FONT_BOLD Then
                        ' 開始予定日が手動入力（日付が太字）されている場合は優先
                        startPlanDate = CDate(sh.Range(C_STARTPLAN_COL & h).Value)
                    Else
                        Dim tmpChargeEndDate, tmpGroupEndDate As Date
                        ' 担当者のタスク末端日付
                        If chargeName = "清海" Then
                            Dim test As String
                            test = "test"
                        End If
                        tmpChargeEndDate = GetChargeEndDate(chargeName)
                        ' 現在のグループのタスク末端日付
                        tmpGroupEndDate = startPlanDate
                        
                        ' 直近の自分のタスク終了日を起点にタスク末端日付を考慮するか判定する
                        Dim myTaskStr As String: myTaskStr = sh.Range(C_MYTASK_COL & h).Value
                        If myTaskStr <> "" Then
                            ' 担当者のタスク末端日付の後に繋げる
                            startPlanDate = DateAdd("d", 1, tmpChargeEndDate)
                            
                        Else
                            ' より末端の日付以降に続けてガントチャートを引く
                            If tmpChargeEndDate > tmpGroupEndDate Then
                                ' 担当者のタスク末端日付の後に繋げる
                                startPlanDate = DateAdd("d", 1, tmpChargeEndDate)
                            Else
                                ' 現在のグループのタスク末端日付の後に繋げる
                                startPlanDate = DateAdd("d", 1, tmpGroupEndDate)
                            End If
                        
                        End If
                        
                    End If
                    
                    ' ガントチャートを引く
                    startPlanDate = SearchCharge(chargeName, chargeColor, g, startPlanDate, h)
                    ' ガントチャートを引いた結果、完了予定日を超過している場合は処理終了
                    If startPlanDate = 0 Then
                        Exit Sub
                    End If
                    
                End If
            End If
            Set tmpChargeRng = Nothing
ContinueGroup:
        Next
Continue:
        Set findGroupStRng = Nothing
        Set findGroupEdRng = Nothing
    Next
        
    MsgBox "ガントチャートの作成が完了しました。"

GetDayRangeError:
    Set wb = Nothing
    Set sh = Nothing
    Set shConf = Nothing

End Sub

' 曜日を数値⇒日本語に変換して返却
Function ConvWeekDayJp(ByVal curDate As String) As String
    Select Case Weekday(curDate)
    Case vbSunday
        ConvWeekDayJp = "日"
    Case vbMonday
        ConvWeekDayJp = "月"
    Case vbTuesday
        ConvWeekDayJp = "火"
    Case vbWednesday
        ConvWeekDayJp = "水"
    Case vbThursday
        ConvWeekDayJp = "木"
    Case vbFriday
        ConvWeekDayJp = "金"
    Case vbSaturday
        ConvWeekDayJp = "土"
    End Select
End Function

' 列のアルファベットを数値へ、数値をアルファベットへ変換
Function ConvertNumAlp(ByVal va As Variant) As Variant
    Dim al As String
    
    If IsNumeric(va) = True Then '数値の場合
        al = Cells(1, va).Address(RowAbsolute:=False, ColumnAbsolute:=False) '$無しでAddress取得
        ConvertNumAlp = Left(al, Len(al) - 1)
    Else 'アルファベットの場合
        ConvertNumAlp = Range(va & "1").Column '列番号を取得
    End If
     
End Function


' Long⇒RGBに変換
Function GetRGBValue(ByVal lColorValue As Long) As Variant
    Dim Red, Green, Blue As Long
    Red = lColorValue Mod 256
    Green = Int(lColorValue / 256) Mod 256
    Blue = Int(lColorValue / 256 / 256)
    
    GetRGBValue = RGB(Red, Green, Blue)
    
End Function


' 指定範囲内で処理対象グループにおける担当者名をすべて検索し、ガントチャートを引く
Function SearchCharge(ByVal chargeName As String, ByVal chargeColor As Variant, ByVal chargeRank As Long, _
                        ByVal startPlanDate As Date, ByVal curRow As Long) As Date
    ' シートから担当者名を検索
    
    ' 対象の行のグループ番号
    ' グループ単位
    Dim searchChargeRng, findChargeRng As Range
    Dim curRank As Long
    Set searchChargeRng = sh.Range(C_CHARGE_COL & (curRow - 1) & ":" & C_CHARGE_COL & maxRow)
    Set findChargeRng = searchChargeRng.Find(chargeName, LookAt:=xlWhole)
    If findChargeRng Is Nothing Then GoTo SearchChargeError
    
    ' 該当行のグループ番号
    curRank = sh.Range(C_GROUP_COL & findChargeRng.Row).Value
    
    Do
        ' 該当行のグループ番号と処理対象のグループ番号が異なる場合は次ループへ
        If curRank <> chargeRank Then GoTo NextDo
        
        ' 開始予定日が太字か判定
        Dim startPlanBoldFlg As Boolean: startPlanBoldFlg = False
        Dim setTaskLineCompFlg As Boolean: setTaskLineCompFlg = False
        If sh.Range(C_STARTPLAN_COL & findChargeRng.Row).Font.FontStyle = C_FONT_BOLD Then
            startPlanBoldFlg = True
        ElseIf sh.Range(C_STARTPLAN_COL & findChargeRng.Row).Value <> "" Then
            ' 開始予定日が細字で既に記載されている場合
            setTaskLineCompFlg = True
        End If
        
        ' 予定工数取得
        Dim manHourPlanVal As String
        manHourPlanVal = sh.Range(C_MANHOUR_COL & CStr(findChargeRng.Row)).Value
            
        If IsNumeric(manHourPlanVal) = True And setTaskLineCompFlg = False Then
            ' 予定工数を数字に変換
            Dim manHourPlanNum As Long
            manHourPlanNum = CLng(manHourPlanVal)
        
            ' 開始日取得
            Dim startResVal As String
            Dim startResDate As Date: startResDate = 0
            startResVal = sh.Range(C_STARTRES_COL & CStr(findChargeRng.Row)).Value
            If IsDate(startResVal) Then
                startResDate = CDate(startResVal)
            End If
            ' 完了日取得
            Dim endResVal As String
            Dim endResDate As Date: endResDate = 0
            endResVal = sh.Range(C_ENDRES_COL & CStr(findChargeRng.Row)).Value
            If IsDate(endResVal) Then
                endResDate = CDate(endResVal)
            End If

            ' ガントチャートを引く（計画）
            Dim tmpStartPlanDate As Date
            tmpStartPlanDate = SetTaskLine(findChargeRng.Row, chargeName, chargeColor, startPlanDate, startPlanBoldFlg, manHourPlanNum)
            If tmpStartPlanDate = 0 Then
                ' ガントチャートを引いた結果、完了予定日を超過している場合
                SearchCharge = 0
                Exit Function
            End If
            
            ' ガントチャートを引く（実績）
            Dim tmpStartResDate As Date
            tmpStartResDate = SetTaskLine(findChargeRng.Row, chargeName, chargeColor, startPlanDate, startPlanBoldFlg, 0)
            
            startPlanDate = tmpStartPlanDate
            
        End If

NextDo:
        Dim findChargeRngNext As Range
        Set findChargeRngNext = searchChargeRng.FindNext(findChargeRng)
        If findChargeRngNext Is Nothing Then Exit Do
        
        If findChargeRng.Address <> findChargeRngNext.Address Then
            Set findChargeRng = findChargeRngNext
            ' グループを取得
            curRank = CLng(sh.Range(C_GROUP_COL & findChargeRng.Row).Value)
        Else
            Exit Do
        End If
        
    Loop
    
SearchChargeError:
    SearchCharge = startPlanDate
    
End Function


' ガントチャートを引く
Function SetTaskLine(ByVal curRow As Long, ByVal chargeName As String, ByVal chargeColor As Variant, _
                             ByVal startDate As Date, ByVal startPlanBoldFlg As Boolean, ByVal manHourPlanNum As Long) As Date

    Dim tmpStartDate, tmpEndDate As Date
    Dim tmpStartResDate, tmpEndResDate As Date
    Dim startMonth, startDay As Long
    Dim startMonthRng, startDayRng, endDayRng, startDayResRng, endDayResRng As Range
    startMonth = Month(startDate)
    
    Set startMonthRng = sh.Range(C_MONTH_ROW & ":" & C_MONTH_ROW).Find(startMonth, LookAt:=xlWhole)
    If Not startMonthRng Is Nothing Then
        Dim startMonthColNum, nextMonthColNum As Long

        ' 土日祝日、非稼働日を加味した開始予定日を取得
        ' tmpStartDate = GetWorkDay(curRow, chargeName, startDate, startPlanBoldFlg)
        tmpStartDate = GetWorkDay(curRow, chargeName, startDate)
        tmpEndDate = tmpStartDate

        ' 工数を加算していき、完了日を算出
        If manHourPlanNum <> 0 Then
            ' 計画の場合
            
            ' 予定工数が【30の倍数】の場合、月末までの日数を計算
            'If manHourPlanNum Mod 30 = 0 Then
            '    Dim tmpFirstDay As String: tmpFirstDay = Format(startDate, "yyyy/mm/01")
            '    Dim monthCount As String: monthCount = CStr(manHourPlanNum / 30)
            '    Dim tmpNextMonth As String: tmpNextMonth = DateAdd("m", monthCount, tmpFirstDay)
            '    Dim tmpLastDay As String: tmpLastDay = DateAdd("d", -1, tmpNextMonth)
            '    Dim dayCount As String: dayCount = Format(tmpLastDay, "d")
            '
            '    manHourPlanNum = dayCount
            'End If
            
            Dim i As Integer
            For i = 1 To manHourPlanNum - 1
                ' 1日加算するごとに土日祝日・非稼働日チェックを行う
                ' tmpEndDate = GetWorkDay(curRow, chargeName, DateAdd("d", 1, tmpEndDate), startPlanBoldFlg)
                tmpEndDate = GetWorkDay(curRow, chargeName, DateAdd("d", 1, tmpEndDate))
                
                ' 完了予定日を超えていないかチェック
                If tmpEndDate > maxDate Then
                    MsgBox "完了予定日以内に全タスクを消化できません。" & vbCrLf & _
                                "完了予定日またはタスクを見直してください。"
                    SetTaskLine = 0
                    Exit Function
                End If
            Next
        Else
            ' 実績の場合
            ' 開始日取得
            Dim tmpStartResDateVal As String
            tmpStartResDateVal = sh.Range(C_STARTRES_COL & curRow).Value
            If IsDate(tmpStartResDateVal) Then
                tmpStartResDate = CDate(tmpStartResDateVal)
                ' 開始日のセル位置を取得
                Set startDayResRng = GetDayRange(tmpStartResDate)
            Else
                Set startDayResRng = Nothing
            End If
        
            ' 完了日取得
            Dim tmpEndResDateVal As String
            tmpEndResDateVal = sh.Range(C_ENDRES_COL & curRow).Value
            If IsDate(tmpEndResDateVal) Then
                tmpEndResDate = CDate(tmpEndResDateVal)
                ' 完了日のセル位置を取得
                Set endDayResRng = GetDayRange(tmpEndResDate)
            Else
                Set endDayResRng = Nothing
            End If
            
        End If
        
        ' 開始予定日のセル位置を取得
        Set startDayRng = GetDayRange(tmpStartDate)
        
        ' 完了予定日のセル位置を取得
        Set endDayRng = GetDayRange(tmpEndDate)
        
        If startDayRng Is Nothing Or endDayRng Is Nothing Then
            GoTo SetTaskLineError
        End If
        
        ' セル着色
        If manHourPlanNum <> 0 Then
            Dim j As Long
            For j = startDayRng.Column To endDayRng.Column
                Dim bkcolor As Variant
                bkcolor = GetRGBValue(sh.Cells(curRow, j).Interior.Color)
        
                ' 開始予定日が太字の場合、またはセル背景色＝白の場合のみ着色
                ' If startPlanBoldFlg = True Or bkcolor = C_WHITE_COLOR Or bkcolor = C_NOW_COLOR Then
                If bkcolor = C_WHITE_COLOR Or bkcolor = C_NOW_COLOR Then
                    ' 計画の場合
                    ' 開始/完了予定日を入力（フォント太字の場合は手動入力として優先）
                    If sh.Cells(curRow, C_STARTPLAN_COL).Font.FontStyle <> C_FONT_BOLD Then
                        sh.Cells(curRow, C_STARTPLAN_COL).Value = tmpStartDate
                    End If
                    sh.Cells(curRow, C_ENDPLAN_COL).Value = tmpEndDate
                    sh.Cells(curRow, j).Interior.Color = chargeColor
                End If
            Next
        Else
            ' 実績の場合
            If Not startDayResRng Is Nothing And Not endDayResRng Is Nothing Then
                ' 完了日まで太枠線で囲う
                Dim k As Long
                For k = startDayResRng.Column To endDayResRng.Column
                    With sh.Cells(curRow, k)
                        Dim borderColor As Variant
                        Dim leftBorder, rightBorder As Border
                        
                        ' 開始日が太字の場合、進行中（未完了）として枠の色を変える
                        If sh.Range(C_STARTRES_COL & curRow).Font.FontStyle = C_FONT_BOLD Then
                            borderColor = C_BORDER_COLOR_BLUE
                        Else
                            borderColor = C_BORDER_COLOR_RED
                        End If
                        
                        Set leftBorder = .Borders(xlEdgeLeft)
                        Set rightBorder = .Borders(xlEdgeRight)
                        
                        ' 罫線を引く
                        .Borders(xlEdgeTop).Weight = xlThick
                        .Borders(xlEdgeBottom).Weight = xlThick
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).Color = borderColor
                        .Borders(xlEdgeBottom).Color = borderColor
                        If leftBorder.Color = borderColor And leftBorder.Weight = xlThick Then
                            .Borders(xlEdgeLeft).Weight = xlThin
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeLeft).Color = C_WBSLINE_COLOR
                        Else
                            .Borders(xlEdgeLeft).Weight = xlThick
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeLeft).Color = borderColor
                        End If
                        If rightBorder.Color = borderColor And rightBorder.Weight = xlThick Then
                            .Borders(xlEdgeRight).Weight = xlThin
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).Color = C_WBSLINE_COLOR
                        Else
                            .Borders(xlEdgeRight).Weight = xlThick
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).Color = borderColor
                        End If
                        Set leftBorder = Nothing
                        Set rightBorder = Nothing
                    End With
                Next
            End If
        End If
        
    End If
    
    ' 完了日を返却
    SetTaskLine = tmpEndDate

SetTaskLineError:
    
End Function


' 土日祝日,非稼働日を加味した稼働日を返却
' Function GetWorkDay(ByVal curRow As Long, ByVal chargeName, ByVal tmpDate As Date, ByVal startPlanBoldFlg As Boolean) As Date
Function GetWorkDay(ByVal curRow As Long, ByVal chargeName, ByVal tmpDate As Date) As Date
    Dim tmpStSuDateRng, tmpHoliDateRng, tmpHoliDayRng, tmpNoWkDateRng, tmpMonthRng, tmpDayRng As Range
    Dim checkOkFlg As Boolean: checkOkFlg = False
        
    ' If startPlanBoldFlg = True Then
    '     GetWorkDay = tmpDate
    '     Exit Function
    ' End If
    
    Do
        ' 土日チェック
        Do
            Select Case Weekday(tmpDate)
                Case vbSaturday, vbSunday
                    ' tmpDateが土日に該当する
                    ' 開始予定日が太字の場合は日付加算しない（以下同様）
                    tmpDate = DateAdd("d", 1, tmpDate)
                Case Else
                    Exit Do
            End Select
        Loop
    
        ' 祝日チェック
        Do
            ' tmpDateが祝日に該当するか祝日一覧から検索
            Set tmpHoliDateRng = shConf.Range(C_HOLIDAY_COL & ":" & C_HOLIDAY_COL).Find(tmpDate, LookAt:=xlWhole)
            If Not tmpHoliDateRng Is Nothing Then
                ' tmpDateが祝日に該当する
                tmpDate = DateAdd("d", 1, tmpDate)
            Else
                ' tmpDateが祝日に該当しない
                Exit Do
            End If
        Loop While tmpHoliDateRng.Address <> ""
        
        ' 非稼働日チェック
        Do
            ' tmpDateが非稼働日に該当するか非稼働日一覧から検索
            Set tmpNoWkDateRng = shConf.Range(C_NOWORKDAY_COL & ":" & C_NOWORKDAY_COL).Find(tmpDate, LookAt:=xlWhole)
            If Not tmpNoWkDateRng Is Nothing Then
                ' tmpDateが非稼働日に該当する
                tmpDate = DateAdd("d", 1, tmpDate)
            Else
                ' tmpDateが非稼働日に該当しない
                Exit Do
            End If
        Loop While tmpNoWkDateRng.Address <> ""
    
        ' 担当者別の非稼働曜日、日を考慮
        Dim resultDate As Date
        resultDate = SetChargeNoWorkDay(chargeName, tmpDate, 0)
        If resultDate <> tmpDate Then
            tmpDate = resultDate
        Else
            Exit Do
        End If
    Loop
    
    GetWorkDay = tmpDate
    
End Function


' 日付の列位置（3行目）を取得
Function GetDayRange(ByVal tmpDate As Variant) As Range
    Dim tmpYear, tmpMonth, tmpDay As Long
    Dim tmpMonthColNum, tmpDayColNum, nextMonthColNum As Long
    tmpYear = Year(tmpDate)
    tmpMonth = Month(tmpDate)
    tmpDay = Day(tmpDate)
    
    Dim tmpYearRng, tmpCurYearRng, tmpNextYearRng, tmpMonthRng, tmpDayRng As Range
    Dim curYearCol, nextYearCol As Variant
    ' 対象日付の年を考慮
    ' 2019/8/1を2019年の日付として認識する。（2020/8/1と区別）
    Set tmpYearRng = sh.Range(C_YEAR_ROW & ":" & C_YEAR_ROW).Find(tmpYear, LookAt:=xlWhole)
    ' ガントチャート上に対象年度が見つからない場合
    If tmpYearRng Is Nothing Then
        GoTo GetDayRangeError
    End If
    
    Set tmpNextYearRng = sh.Range(C_YEAR_ROW & ":" & C_YEAR_ROW).Find(tmpYear + 1, LookAt:=xlWhole)
    If tmpNextYearRng Is Nothing Then
        ' 右端まで検索
        nextYearCol = "XFD"
    Else
        nextYearCol = ConvertNumAlp(tmpNextYearRng.Column)
    End If
    curYearCol = ConvertNumAlp(tmpYearRng.Column)
    Set tmpMonthRng = sh.Range(curYearCol & C_MONTH_ROW & ":" & nextYearCol & C_MONTH_ROW).Find(tmpMonth, LookAt:=xlWhole)
    
    If Not tmpMonthRng Is Nothing Then
        ' 当月のセル位置取得
        tmpMonthColNum = tmpMonthRng.Column
        ' 翌月のセル位置取得
        ' 翌月月初のセル位置が末端の場合は？
        nextMonthColNum = tmpMonthRng.End(xlToRight).Column - 1
        
        ' セル検索（当月から翌月月初までの間で日付検索）
        Dim st, ed As Variant
        st = ConvertNumAlp(tmpMonthColNum)
        ed = ConvertNumAlp(nextMonthColNum)
        Set GetDayRange = sh.Range(st & C_DAY_ROW & ":" & ed & C_DAY_ROW).Find(tmpDay, LookAt:=xlWhole)
        
    End If
    
GetDayRangeError:

End Function


' 担当者のガントチャート末端の日付を取得
Function GetChargeEndDate(ByVal chargeName As String) As Date
    Dim tmpRng As Range
    Dim tmpRngAdr, tmpCellVal As String
    Dim tmpDate, tmpMaxDate As Date
    Set tmpRng = sh.Range(C_CHARGE_COL & C_HEADER_ROW & ":" & C_CHARGE_COL & maxRow).Find(chargeName, LookAt:=xlWhole)
    tmpRngAdr = tmpRng.Address
    
    Do
        If Not tmpRng Is Nothing Then
            tmpCellVal = sh.Range(C_ENDPLAN_COL & tmpRng.Row).Value
            ' セルの値が日付型の場合のみ
            If IsDate(tmpCellVal) = True Then
                tmpDate = CDate(tmpCellVal)
                If tmpMaxDate < tmpDate Then
                    ' 日付の最大値を保持
                    tmpMaxDate = tmpDate
                End If
            End If
        Else
            Exit Do
        End If
        
        Set tmpRng = sh.Range(C_CHARGE_COL & C_HEADER_ROW & ":" & C_CHARGE_COL & maxRow).FindNext(tmpRng)
        If tmpRng Is Nothing Then
            Exit Do
        End If

    Loop While tmpRngAdr <> tmpRng.Address

    If tmpMaxDate = 0 Then
        GetChargeEndDate = minDate
    Else
        GetChargeEndDate = tmpMaxDate
    End If

    
End Function


' 担当者の非稼働曜日、日をWBSに反映
' chargeName: 担当者名, tmpDate: 対象日, curCol: 対象日の列番号
' chargeNameあり: ガントチャート設定時、なし：カレンダー作成時
Function SetChargeNoWorkDay(ByVal chargeName As String, ByVal tmpDate As Date, ByVal curCol As Long) As Date
    
    If chargeName <> "" Then
        ' WBS作成時
        SetChargeNoWorkDay = GetChargeWorkDay(chargeName, tmpDate, 0)
    Else
        ' カレンダー設定時
        Dim confChargeColEnd As Long
        confChargeColEnd = shConf.Cells(C_CONFHEADER_ROW, Columns.Count).End(xlToLeft).Column
        
        Dim i As Long
        For i = ConvertNumAlp(C_CONF_CHARGENAME_COL) To confChargeColEnd
            ' 担当者名を取得
            chargeName = shConf.Cells(C_CONFHEADER_ROW, i).Value
            
            Call GetChargeWorkDay(chargeName, tmpDate, curCol)
        Next
    
    End If
        
End Function


' 担当者別の非稼働曜日、日を考慮した稼働日を算出
Function GetChargeWorkDay(ByVal chargeName As String, ByVal tmpDate As Date, ByVal curCol As Long) As Date
    ' chargeFromWbsRngに対するFindNextが効かないため、Forループで対応
    Dim maxRow As Long
    
     ' WBSから担当者の行を取得
    Dim chargeFromWbsRng, chargeFromWbsRngMax As Range
    Set chargeFromWbsRng = sh.Range(C_CHARGE_COL & ":" & C_CHARGE_COL).Find(chargeName, LookAt:=xlWhole)
    If chargeFromWbsRng Is Nothing Then
        ' ！！！要見直し！！！
        GetChargeWorkDay = tmpDate
        Exit Function
    Else
        Set chargeFromWbsRngMax = sh.Range(C_CHARGE_COL & ":" & C_CHARGE_COL).FindPrevious(chargeFromWbsRng)
        maxRow = chargeFromWbsRngMax.Row
    End If

    Dim i As Long
    For i = C_HEADER_ROW + 1 To maxRow
        If chargeName <> sh.Cells(i, ConvertNumAlp(C_CHARGE_COL)).Value Then GoTo NextLoop
    
        ' 担当者の非稼働曜日に該当しないかチェック

        ' configシートから担当者検索
        Dim chargeNameRng As Range
        Set chargeNameRng = shConf.Range(C_CONFHEADER_ROW & ":" & C_CONFHEADER_ROW).Find(chargeName, LookAt:=xlWhole)
        
        If chargeNameRng Is Nothing Then
            ' ！！！要見直し！！！
            GetChargeWorkDay = tmpDate
            Exit Function
        End If

        ' 担当者の非稼働曜日チェック
        Dim chargeNoWorkWeekList As Variant
        Dim chargeNoWorkWeekStr As String
        chargeNoWorkWeekStr = shConf.Cells(C_CONF_CHARGENOWKWEEK_ROW, chargeNameRng.Column).Value
        chargeNoWorkWeekList = Split(chargeNoWorkWeekStr, ",")
        
        Dim s As Integer
        For s = LBound(chargeNoWorkWeekList) To UBound(chargeNoWorkWeekList)
            ' tmpDateの曜日を取得
            Dim tmpDateWeekDay As String
            tmpDateWeekDay = ConvWeekDayJp(tmpDate)
            If tmpDateWeekDay = chargeNoWorkWeekList(s) Then
                ' 担当者の非稼働曜日に該当する
                If curCol <> 0 Then
                    ' 該当行のセルの色をグレーにする
                    sh.Cells(i, curCol).Interior.Color = C_NOWORKDAY_COLOR
                Else
                    ' 日付加算
                    tmpDate = IIf(chargeName <> "", DateAdd("d", 1, tmpDate), tmpDate)
                End If
            End If
        Next s
    
        ' 担当者の非稼働日チェック
        Dim chargeNoWorkDayRng As Range
        Set chargeNoWorkDayRng = shConf.Range(ConvertNumAlp(chargeNameRng.Column) & ":" & _
                                        ConvertNumAlp(chargeNameRng.Column)).Find(tmpDate, LookAt:=xlWhole)
        Do
            ' 担当者の非稼働日に該当しないかチェック
            If chargeNoWorkDayRng Is Nothing Then Exit Do
            
            ' 担当者の非稼働日に該当する
            If curCol <> 0 Then
                ' 該当行のセルの色をグレーにする
                sh.Cells(i, curCol).Interior.Color = C_NOWORKDAY_COLOR
            Else
                ' 日付加算
                tmpDate = IIf(chargeName <> "", DateAdd("d", 1, tmpDate), tmpDate)
            End If
            
            Dim chargeNoWorkDayRngNext As Range
            Set chargeNoWorkDayRngNext = shConf.Range(ConvertNumAlp(chargeNameRng.Column) & ":" & _
                                            ConvertNumAlp(chargeNameRng.Column)).FindNext(After:=chargeNoWorkDayRng)
            If chargeNoWorkDayRngNext Is Nothing Then Exit Do
            If chargeNoWorkDayRng.Address <> chargeNoWorkDayRngNext.Address Then
                Set chargeNoWorkDayRng = chargeNoWorkDayRngNext
            Else
                Exit Do
            End If
        Loop
        
        ' WBS作成時、1ループで終了
        If curCol = 0 Then Exit For
NextLoop:
        
    Next
    
    GetChargeWorkDay = tmpDate

End Function
