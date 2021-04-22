Attribute VB_Name = "WBS�쐬"
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
    
    ' �ŉ��s�擾
    maxRow = sh.Cells(Rows.Count, ConvertNumAlp(C_NO_COL)).End(xlUp).Row

    ' WBS�N���A
    Dim k As Long
    For k = C_HEADER_ROW + 1 To maxRow
        ' �J�n�\����̃t�H���g�F�擾
        If sh.Range(C_STARTPLAN_COL & k).Font.FontStyle <> C_FONT_BOLD Then
            sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).ClearContents
        ElseIf sh.Range(C_STARTPLAN_COL & k).Value = "" And _
            IsNull(sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).Font.FontStyle) = True Then
            ' �����̑�����߂��i�߂��Y��h�~�j
            sh.Range(C_STARTPLAN_COL & k & ":" & C_ENDPLAN_COL & k).Font.FontStyle = C_FONT_REGULAR
        Else
            sh.Range(C_ENDPLAN_COL & k).ClearContents
        End If
    Next
    ' �N�F�k�����đS�̕\��
    sh.Range(C_STARTWBS_COL & ":" & "XFD").Clear
    sh.Range(C_STARTWBS_COL & C_YEAR_ROW & ":" & "XFD" & C_YEAR_ROW).ShrinkToFit = True
    
End Sub

Sub CreateWBS()
    ' ������
    Call Init
    
    '========================================================================================
    ' �J�����_�[�쐬
    '========================================================================================
    
    ' �J�n�\���MIN�擾
    minDate = sh.Cells(C_MONTH_ROW, C_PJSTARTDAY_COL).Value
    If IsDate(minDate) = False Then
        MsgBox "�J�n�\��������ŃG���[�ƂȂ�܂����B�������I�����܂��B"
        Exit Sub
    End If
    
    ' �����\���MAX�擾
    maxDate = sh.Cells(C_MONTH_ROW, C_PJENDDAY_COL).Value
    If IsDate(maxDate) = False Then
        MsgBox "�����\��������ŃG���[�ƂȂ�܂����B�������I�����܂��B"
        Exit Sub
    End If
    
    ' �J�n��MAX�̌����`�ŏI��MAX�̌����܂Ń`���[�g�����
    Dim startWbsColNum As Long
    startWbsColNum = ConvertNumAlp(C_STARTWBS_COL)
    
    ' �J�n�\���MIN�̌��Ɗ����\���MAX�̌��̍����������擾
    Dim monthCount As Long
    monthCount = DateDiff("m", minDate, maxDate)
    
    Dim curColNum As Long: curColNum = 0
    Dim i As Integer
    For i = 0 To monthCount
        ' �Ώی��擾
        Dim targetDate As Date
        targetDate = DateAdd("m", i, minDate)
        
        ' ���̓������擾
        Dim lastDate As Date
        Dim lastDay As Integer
        lastDate = DateSerial(Year(targetDate), Month(targetDate) + 1, 0)
        lastDay = Format(lastDate, "d")
        
        Dim j As Integer
        For j = 1 To lastDay
            If j = 1 Then
                sh.Cells(C_MONTH_ROW, startWbsColNum + curColNum).Value = Month(targetDate)
                ' �N���L��
                If i = 0 Or Month(targetDate) = 1 Then
                    sh.Cells(C_YEAR_ROW, startWbsColNum + curColNum).Value = Year(targetDate)
                End If
            End If
            
            ' �����L��
            sh.Cells(C_DAY_ROW, startWbsColNum + curColNum).Value = j
            
            Dim curDate As Date
            curDate = DateSerial(Year(targetDate), Month(targetDate), j)
            
            Dim curCol As String
            curCol = ConvertNumAlp(startWbsColNum + curColNum)
            
            With sh.Cells(C_HEADER_ROW, startWbsColNum + curColNum)
                ' �j���i���l�j����{��ɕϊ�
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
            
            ' �y���j������
            Dim findRng1, findRng2 As Range
            Set findRng1 = shConf.Range(C_HOLIDAY_COL & ":" & C_HOLIDAY_COL).Find(curDate, LookAt:=xlWhole)
            If Not findRng1 Is Nothing Then
                ' �y���j���̏ꍇ�͗���O���[�ɂ���
                sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
            End If
            ' ��ғ�������
            Set findRng2 = shConf.Range(C_NOWORKDAY_COL & ":" & C_NOWORKDAY_COL).Find(curDate, LookAt:=xlWhole)
            If Not findRng2 Is Nothing Then
                ' ��ғ����̏ꍇ�͗���O���[�ɂ���
                sh.Range(curCol & (C_HEADER_ROW + 1) & ":" & curCol & maxRow).Interior.Color = C_NOWORKDAY_COLOR
            End If
            
            ' �S���ҕʂ̔�ғ��j���A���̃Z�����O���[�ɂ���
            Call SetChargeNoWorkDay("", curDate, ConvertNumAlp(curCol))
            
            curColNum = curColNum + 1
            
            Set findRng1 = Nothing
            Set findRng2 = Nothing
        Next
    Next
    
    ' �����̓��t��ɐF��t����
    Dim nowRng As Range
    Set nowRng = GetDayRange(Date)
    If Not nowRng Is Nothing Then
        Dim nowRngCol As String: nowRngCol = ConvertNumAlp(nowRng.Column)
        sh.Range(nowRngCol & C_DAY_ROW & ":" & nowRngCol & maxRow).Interior.Color = C_NOW_COLOR
        nowRng.Select
    End If
    
    ' �ł��E�̗�i�A���t�@�x�b�g�j���擾
    Dim maxcolAlp As String
    maxcolAlp = ConvertNumAlp(startWbsColNum + curColNum - 1)
    
    ' �r��������
    With sh.Range(C_STARTWBS_COL & C_DAY_ROW & ":" & maxcolAlp & maxRow).Borders
        .LineStyle = xlContinuous
        .Color = C_WBSLINE_COLOR
    End With
    sh.Range(C_STARTWBS_COL & C_MONTH_ROW & ":" & maxcolAlp & maxRow).BorderAround Color:=C_WBSLINE_COLOR
    
    With sh.Range(C_STARTWBS_COL & ":" & maxcolAlp)
        ' �񕝒���
        .EntireColumn.ColumnWidth = 2.45
        ' ��������
        .HorizontalAlignment = xlCenter
        ' �t�H���g�ݒ�
        .Font.Name = "���C���I"
        .Font.Size = "9"
    End With
        
    '========================================================================================
    ' WBS�쐬
    '========================================================================================

    ' �J�n�\����擾
    Dim startPlanDate As Date
        
    ' �O���[�v�ʂɃK���g�`���[�g������
    Dim maxGroup As Long
    maxGroup = Application.WorksheetFunction.Max(sh.Range(C_GROUP_COL & (C_HEADER_ROW + 1) & _
                ":" & C_GROUP_COL & maxRow))
    
    Dim chargeName As String
    Dim chargeColor As Variant
    Dim chargeNameColNum As Long
    chargeNameColNum = ConvertNumAlp(C_CONF_CHARGENAME_COL)
    ' �S���҃}�X�^�ŉ��s�擾
    Dim maxChargeRow As Long
    maxChargeRow = shConf.Cells(Rows.Count, chargeNameColNum).End(xlUp).Row
    
    Dim g As Integer
    For g = 1 To maxGroup
        ' WBS�V�[�g����S���Җ�������
        Dim findGroupStRng, findGroupEdRng As Range
        ' �ォ�猟��
        Set findGroupStRng = sh.Range(C_GROUP_COL & C_HEADER_ROW & ":" & _
                                C_GROUP_COL & maxRow).Find(g, LookAt:=xlWhole)
        ' �����猟��
        Set findGroupEdRng = sh.Range(C_GROUP_COL & C_HEADER_ROW & ":" & _
                                C_GROUP_COL & maxRow).Find(g, LookAt:=xlWhole, SearchDirection:=xlPrevious)

        ' �������ʂ�Nothing�̏ꍇ�͎����[�v��
        If findGroupStRng Is Nothing Or findGroupEdRng Is Nothing Then
            GoTo Continue
        End If

        ' �S���Җ����擾
        chargeName = sh.Range(C_CHARGE_COL & findGroupStRng.Row).Value
        
        ' �S���҂̃^�X�N���[���t���擾
        startPlanDate = GetChargeEndDate(chargeName)
        
        Dim h As Long
        For h = findGroupStRng.Row To findGroupEdRng.Row
            ' �O���[�v�ԍ����擾
            Dim orderCellVal As String
            orderCellVal = sh.Range(C_GROUP_COL & h).Value
            
            ' ���݂̃O���[�v�ԍ��Ə����Ώۂ̃O���[�v�ԍ����قȂ�ꍇ�͎����[�v��
            If g <> orderCellVal Then GoTo ContinueGroup
            
            ' �X�e�[�^�X���擾
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
                        
            ' �\��H�����擾
            Dim manHourPlanVal As String
            manHourPlanVal = sh.Range(C_MANHOUR_COL & h).Value
            ' �S���Җ����擾
            chargeName = sh.Range(C_CHARGE_COL & h).Value
            ' �S���҃}�X�^����S���҂̐F���擾
            Dim tmpChargeRng As Range
            Dim chargeEndCol As String
            ' �S���҃}�X�^�̗񖖒[���擾
            chargeEndCol = shConf.Cells(C_CONFHEADER_ROW, Columns.Count).End(xlToLeft).Column
            Set tmpChargeRng = shConf.Range(C_CONF_CHARGENAME_COL & C_CONFHEADER_ROW & ":" & _
                                    ConvertNumAlp(CInt(chargeEndCol)) & C_CONFHEADER_ROW).Find(chargeName, LookAt:=xlWhole)
            If Not tmpChargeRng Is Nothing Then
                ' �S���҂̐F���擾
                chargeColor = GetRGBValue(shConf.Range(ConvertNumAlp(tmpChargeRng.Column) & C_CONF_CHARGECLR_ROW).Interior.Color)
                
                ' �O���[�v�P�ʂŃK���g�`���[�g������
                ' �\��H���ƃO���[�v�ԍ��̐��l���́A�J�n�\����ɓ��t�������Ă��Ȃ��s�̂�
                If IsNumeric(manHourPlanVal) = True And IsNumeric(orderCellVal) = True Then
                    If sh.Range(C_STARTPLAN_COL & h).Font.FontStyle = C_FONT_BOLD Then
                        ' �J�n�\������蓮���́i���t�������j����Ă���ꍇ�͗D��
                        startPlanDate = CDate(sh.Range(C_STARTPLAN_COL & h).Value)
                    Else
                        Dim tmpChargeEndDate, tmpGroupEndDate As Date
                        ' �S���҂̃^�X�N���[���t
                        If chargeName = "���C" Then
                            Dim test As String
                            test = "test"
                        End If
                        tmpChargeEndDate = GetChargeEndDate(chargeName)
                        ' ���݂̃O���[�v�̃^�X�N���[���t
                        tmpGroupEndDate = startPlanDate
                        
                        ' ���߂̎����̃^�X�N�I�������N�_�Ƀ^�X�N���[���t���l�����邩���肷��
                        Dim myTaskStr As String: myTaskStr = sh.Range(C_MYTASK_COL & h).Value
                        If myTaskStr <> "" Then
                            ' �S���҂̃^�X�N���[���t�̌�Ɍq����
                            startPlanDate = DateAdd("d", 1, tmpChargeEndDate)
                            
                        Else
                            ' ��薖�[�̓��t�ȍ~�ɑ����ăK���g�`���[�g������
                            If tmpChargeEndDate > tmpGroupEndDate Then
                                ' �S���҂̃^�X�N���[���t�̌�Ɍq����
                                startPlanDate = DateAdd("d", 1, tmpChargeEndDate)
                            Else
                                ' ���݂̃O���[�v�̃^�X�N���[���t�̌�Ɍq����
                                startPlanDate = DateAdd("d", 1, tmpGroupEndDate)
                            End If
                        
                        End If
                        
                    End If
                    
                    ' �K���g�`���[�g������
                    startPlanDate = SearchCharge(chargeName, chargeColor, g, startPlanDate, h)
                    ' �K���g�`���[�g�����������ʁA�����\����𒴉߂��Ă���ꍇ�͏����I��
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
        
    MsgBox "�K���g�`���[�g�̍쐬���������܂����B"

GetDayRangeError:
    Set wb = Nothing
    Set sh = Nothing
    Set shConf = Nothing

End Sub

' �j���𐔒l�˓��{��ɕϊ����ĕԋp
Function ConvWeekDayJp(ByVal curDate As String) As String
    Select Case Weekday(curDate)
    Case vbSunday
        ConvWeekDayJp = "��"
    Case vbMonday
        ConvWeekDayJp = "��"
    Case vbTuesday
        ConvWeekDayJp = "��"
    Case vbWednesday
        ConvWeekDayJp = "��"
    Case vbThursday
        ConvWeekDayJp = "��"
    Case vbFriday
        ConvWeekDayJp = "��"
    Case vbSaturday
        ConvWeekDayJp = "�y"
    End Select
End Function

' ��̃A���t�@�x�b�g�𐔒l�ցA���l���A���t�@�x�b�g�֕ϊ�
Function ConvertNumAlp(ByVal va As Variant) As Variant
    Dim al As String
    
    If IsNumeric(va) = True Then '���l�̏ꍇ
        al = Cells(1, va).Address(RowAbsolute:=False, ColumnAbsolute:=False) '$������Address�擾
        ConvertNumAlp = Left(al, Len(al) - 1)
    Else '�A���t�@�x�b�g�̏ꍇ
        ConvertNumAlp = Range(va & "1").Column '��ԍ����擾
    End If
     
End Function


' Long��RGB�ɕϊ�
Function GetRGBValue(ByVal lColorValue As Long) As Variant
    Dim Red, Green, Blue As Long
    Red = lColorValue Mod 256
    Green = Int(lColorValue / 256) Mod 256
    Blue = Int(lColorValue / 256 / 256)
    
    GetRGBValue = RGB(Red, Green, Blue)
    
End Function


' �w��͈͓��ŏ����ΏۃO���[�v�ɂ�����S���Җ������ׂČ������A�K���g�`���[�g������
Function SearchCharge(ByVal chargeName As String, ByVal chargeColor As Variant, ByVal chargeRank As Long, _
                        ByVal startPlanDate As Date, ByVal curRow As Long) As Date
    ' �V�[�g����S���Җ�������
    
    ' �Ώۂ̍s�̃O���[�v�ԍ�
    ' �O���[�v�P��
    Dim searchChargeRng, findChargeRng As Range
    Dim curRank As Long
    Set searchChargeRng = sh.Range(C_CHARGE_COL & (curRow - 1) & ":" & C_CHARGE_COL & maxRow)
    Set findChargeRng = searchChargeRng.Find(chargeName, LookAt:=xlWhole)
    If findChargeRng Is Nothing Then GoTo SearchChargeError
    
    ' �Y���s�̃O���[�v�ԍ�
    curRank = sh.Range(C_GROUP_COL & findChargeRng.Row).Value
    
    Do
        ' �Y���s�̃O���[�v�ԍ��Ə����Ώۂ̃O���[�v�ԍ����قȂ�ꍇ�͎����[�v��
        If curRank <> chargeRank Then GoTo NextDo
        
        ' �J�n�\���������������
        Dim startPlanBoldFlg As Boolean: startPlanBoldFlg = False
        Dim setTaskLineCompFlg As Boolean: setTaskLineCompFlg = False
        If sh.Range(C_STARTPLAN_COL & findChargeRng.Row).Font.FontStyle = C_FONT_BOLD Then
            startPlanBoldFlg = True
        ElseIf sh.Range(C_STARTPLAN_COL & findChargeRng.Row).Value <> "" Then
            ' �J�n�\������׎��Ŋ��ɋL�ڂ���Ă���ꍇ
            setTaskLineCompFlg = True
        End If
        
        ' �\��H���擾
        Dim manHourPlanVal As String
        manHourPlanVal = sh.Range(C_MANHOUR_COL & CStr(findChargeRng.Row)).Value
            
        If IsNumeric(manHourPlanVal) = True And setTaskLineCompFlg = False Then
            ' �\��H���𐔎��ɕϊ�
            Dim manHourPlanNum As Long
            manHourPlanNum = CLng(manHourPlanVal)
        
            ' �J�n���擾
            Dim startResVal As String
            Dim startResDate As Date: startResDate = 0
            startResVal = sh.Range(C_STARTRES_COL & CStr(findChargeRng.Row)).Value
            If IsDate(startResVal) Then
                startResDate = CDate(startResVal)
            End If
            ' �������擾
            Dim endResVal As String
            Dim endResDate As Date: endResDate = 0
            endResVal = sh.Range(C_ENDRES_COL & CStr(findChargeRng.Row)).Value
            If IsDate(endResVal) Then
                endResDate = CDate(endResVal)
            End If

            ' �K���g�`���[�g�������i�v��j
            Dim tmpStartPlanDate As Date
            tmpStartPlanDate = SetTaskLine(findChargeRng.Row, chargeName, chargeColor, startPlanDate, startPlanBoldFlg, manHourPlanNum)
            If tmpStartPlanDate = 0 Then
                ' �K���g�`���[�g�����������ʁA�����\����𒴉߂��Ă���ꍇ
                SearchCharge = 0
                Exit Function
            End If
            
            ' �K���g�`���[�g�������i���сj
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
            ' �O���[�v���擾
            curRank = CLng(sh.Range(C_GROUP_COL & findChargeRng.Row).Value)
        Else
            Exit Do
        End If
        
    Loop
    
SearchChargeError:
    SearchCharge = startPlanDate
    
End Function


' �K���g�`���[�g������
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

        ' �y���j���A��ғ��������������J�n�\������擾
        ' tmpStartDate = GetWorkDay(curRow, chargeName, startDate, startPlanBoldFlg)
        tmpStartDate = GetWorkDay(curRow, chargeName, startDate)
        tmpEndDate = tmpStartDate

        ' �H�������Z���Ă����A���������Z�o
        If manHourPlanNum <> 0 Then
            ' �v��̏ꍇ
            
            ' �\��H�����y30�̔{���z�̏ꍇ�A�����܂ł̓������v�Z
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
                ' 1�����Z���邲�Ƃɓy���j���E��ғ����`�F�b�N���s��
                ' tmpEndDate = GetWorkDay(curRow, chargeName, DateAdd("d", 1, tmpEndDate), startPlanBoldFlg)
                tmpEndDate = GetWorkDay(curRow, chargeName, DateAdd("d", 1, tmpEndDate))
                
                ' �����\����𒴂��Ă��Ȃ����`�F�b�N
                If tmpEndDate > maxDate Then
                    MsgBox "�����\����ȓ��ɑS�^�X�N�������ł��܂���B" & vbCrLf & _
                                "�����\����܂��̓^�X�N���������Ă��������B"
                    SetTaskLine = 0
                    Exit Function
                End If
            Next
        Else
            ' ���т̏ꍇ
            ' �J�n���擾
            Dim tmpStartResDateVal As String
            tmpStartResDateVal = sh.Range(C_STARTRES_COL & curRow).Value
            If IsDate(tmpStartResDateVal) Then
                tmpStartResDate = CDate(tmpStartResDateVal)
                ' �J�n���̃Z���ʒu���擾
                Set startDayResRng = GetDayRange(tmpStartResDate)
            Else
                Set startDayResRng = Nothing
            End If
        
            ' �������擾
            Dim tmpEndResDateVal As String
            tmpEndResDateVal = sh.Range(C_ENDRES_COL & curRow).Value
            If IsDate(tmpEndResDateVal) Then
                tmpEndResDate = CDate(tmpEndResDateVal)
                ' �������̃Z���ʒu���擾
                Set endDayResRng = GetDayRange(tmpEndResDate)
            Else
                Set endDayResRng = Nothing
            End If
            
        End If
        
        ' �J�n�\����̃Z���ʒu���擾
        Set startDayRng = GetDayRange(tmpStartDate)
        
        ' �����\����̃Z���ʒu���擾
        Set endDayRng = GetDayRange(tmpEndDate)
        
        If startDayRng Is Nothing Or endDayRng Is Nothing Then
            GoTo SetTaskLineError
        End If
        
        ' �Z�����F
        If manHourPlanNum <> 0 Then
            Dim j As Long
            For j = startDayRng.Column To endDayRng.Column
                Dim bkcolor As Variant
                bkcolor = GetRGBValue(sh.Cells(curRow, j).Interior.Color)
        
                ' �J�n�\����������̏ꍇ�A�܂��̓Z���w�i�F�����̏ꍇ�̂ݒ��F
                ' If startPlanBoldFlg = True Or bkcolor = C_WHITE_COLOR Or bkcolor = C_NOW_COLOR Then
                If bkcolor = C_WHITE_COLOR Or bkcolor = C_NOW_COLOR Then
                    ' �v��̏ꍇ
                    ' �J�n/�����\�������́i�t�H���g�����̏ꍇ�͎蓮���͂Ƃ��ėD��j
                    If sh.Cells(curRow, C_STARTPLAN_COL).Font.FontStyle <> C_FONT_BOLD Then
                        sh.Cells(curRow, C_STARTPLAN_COL).Value = tmpStartDate
                    End If
                    sh.Cells(curRow, C_ENDPLAN_COL).Value = tmpEndDate
                    sh.Cells(curRow, j).Interior.Color = chargeColor
                End If
            Next
        Else
            ' ���т̏ꍇ
            If Not startDayResRng Is Nothing And Not endDayResRng Is Nothing Then
                ' �������܂ő��g���ň͂�
                Dim k As Long
                For k = startDayResRng.Column To endDayResRng.Column
                    With sh.Cells(curRow, k)
                        Dim borderColor As Variant
                        Dim leftBorder, rightBorder As Border
                        
                        ' �J�n���������̏ꍇ�A�i�s���i�������j�Ƃ��Ęg�̐F��ς���
                        If sh.Range(C_STARTRES_COL & curRow).Font.FontStyle = C_FONT_BOLD Then
                            borderColor = C_BORDER_COLOR_BLUE
                        Else
                            borderColor = C_BORDER_COLOR_RED
                        End If
                        
                        Set leftBorder = .Borders(xlEdgeLeft)
                        Set rightBorder = .Borders(xlEdgeRight)
                        
                        ' �r��������
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
    
    ' ��������ԋp
    SetTaskLine = tmpEndDate

SetTaskLineError:
    
End Function


' �y���j��,��ғ��������������ғ�����ԋp
' Function GetWorkDay(ByVal curRow As Long, ByVal chargeName, ByVal tmpDate As Date, ByVal startPlanBoldFlg As Boolean) As Date
Function GetWorkDay(ByVal curRow As Long, ByVal chargeName, ByVal tmpDate As Date) As Date
    Dim tmpStSuDateRng, tmpHoliDateRng, tmpHoliDayRng, tmpNoWkDateRng, tmpMonthRng, tmpDayRng As Range
    Dim checkOkFlg As Boolean: checkOkFlg = False
        
    ' If startPlanBoldFlg = True Then
    '     GetWorkDay = tmpDate
    '     Exit Function
    ' End If
    
    Do
        ' �y���`�F�b�N
        Do
            Select Case Weekday(tmpDate)
                Case vbSaturday, vbSunday
                    ' tmpDate���y���ɊY������
                    ' �J�n�\����������̏ꍇ�͓��t���Z���Ȃ��i�ȉ����l�j
                    tmpDate = DateAdd("d", 1, tmpDate)
                Case Else
                    Exit Do
            End Select
        Loop
    
        ' �j���`�F�b�N
        Do
            ' tmpDate���j���ɊY�����邩�j���ꗗ���猟��
            Set tmpHoliDateRng = shConf.Range(C_HOLIDAY_COL & ":" & C_HOLIDAY_COL).Find(tmpDate, LookAt:=xlWhole)
            If Not tmpHoliDateRng Is Nothing Then
                ' tmpDate���j���ɊY������
                tmpDate = DateAdd("d", 1, tmpDate)
            Else
                ' tmpDate���j���ɊY�����Ȃ�
                Exit Do
            End If
        Loop While tmpHoliDateRng.Address <> ""
        
        ' ��ғ����`�F�b�N
        Do
            ' tmpDate����ғ����ɊY�����邩��ғ����ꗗ���猟��
            Set tmpNoWkDateRng = shConf.Range(C_NOWORKDAY_COL & ":" & C_NOWORKDAY_COL).Find(tmpDate, LookAt:=xlWhole)
            If Not tmpNoWkDateRng Is Nothing Then
                ' tmpDate����ғ����ɊY������
                tmpDate = DateAdd("d", 1, tmpDate)
            Else
                ' tmpDate����ғ����ɊY�����Ȃ�
                Exit Do
            End If
        Loop While tmpNoWkDateRng.Address <> ""
    
        ' �S���ҕʂ̔�ғ��j���A�����l��
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


' ���t�̗�ʒu�i3�s�ځj���擾
Function GetDayRange(ByVal tmpDate As Variant) As Range
    Dim tmpYear, tmpMonth, tmpDay As Long
    Dim tmpMonthColNum, tmpDayColNum, nextMonthColNum As Long
    tmpYear = Year(tmpDate)
    tmpMonth = Month(tmpDate)
    tmpDay = Day(tmpDate)
    
    Dim tmpYearRng, tmpCurYearRng, tmpNextYearRng, tmpMonthRng, tmpDayRng As Range
    Dim curYearCol, nextYearCol As Variant
    ' �Ώۓ��t�̔N���l��
    ' 2019/8/1��2019�N�̓��t�Ƃ��ĔF������B�i2020/8/1�Ƌ�ʁj
    Set tmpYearRng = sh.Range(C_YEAR_ROW & ":" & C_YEAR_ROW).Find(tmpYear, LookAt:=xlWhole)
    ' �K���g�`���[�g��ɑΏ۔N�x��������Ȃ��ꍇ
    If tmpYearRng Is Nothing Then
        GoTo GetDayRangeError
    End If
    
    Set tmpNextYearRng = sh.Range(C_YEAR_ROW & ":" & C_YEAR_ROW).Find(tmpYear + 1, LookAt:=xlWhole)
    If tmpNextYearRng Is Nothing Then
        ' �E�[�܂Ō���
        nextYearCol = "XFD"
    Else
        nextYearCol = ConvertNumAlp(tmpNextYearRng.Column)
    End If
    curYearCol = ConvertNumAlp(tmpYearRng.Column)
    Set tmpMonthRng = sh.Range(curYearCol & C_MONTH_ROW & ":" & nextYearCol & C_MONTH_ROW).Find(tmpMonth, LookAt:=xlWhole)
    
    If Not tmpMonthRng Is Nothing Then
        ' �����̃Z���ʒu�擾
        tmpMonthColNum = tmpMonthRng.Column
        ' �����̃Z���ʒu�擾
        ' ���������̃Z���ʒu�����[�̏ꍇ�́H
        nextMonthColNum = tmpMonthRng.End(xlToRight).Column - 1
        
        ' �Z�������i�������痂�������܂ł̊Ԃœ��t�����j
        Dim st, ed As Variant
        st = ConvertNumAlp(tmpMonthColNum)
        ed = ConvertNumAlp(nextMonthColNum)
        Set GetDayRange = sh.Range(st & C_DAY_ROW & ":" & ed & C_DAY_ROW).Find(tmpDay, LookAt:=xlWhole)
        
    End If
    
GetDayRangeError:

End Function


' �S���҂̃K���g�`���[�g���[�̓��t���擾
Function GetChargeEndDate(ByVal chargeName As String) As Date
    Dim tmpRng As Range
    Dim tmpRngAdr, tmpCellVal As String
    Dim tmpDate, tmpMaxDate As Date
    Set tmpRng = sh.Range(C_CHARGE_COL & C_HEADER_ROW & ":" & C_CHARGE_COL & maxRow).Find(chargeName, LookAt:=xlWhole)
    tmpRngAdr = tmpRng.Address
    
    Do
        If Not tmpRng Is Nothing Then
            tmpCellVal = sh.Range(C_ENDPLAN_COL & tmpRng.Row).Value
            ' �Z���̒l�����t�^�̏ꍇ�̂�
            If IsDate(tmpCellVal) = True Then
                tmpDate = CDate(tmpCellVal)
                If tmpMaxDate < tmpDate Then
                    ' ���t�̍ő�l��ێ�
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


' �S���҂̔�ғ��j���A����WBS�ɔ��f
' chargeName: �S���Җ�, tmpDate: �Ώۓ�, curCol: �Ώۓ��̗�ԍ�
' chargeName����: �K���g�`���[�g�ݒ莞�A�Ȃ��F�J�����_�[�쐬��
Function SetChargeNoWorkDay(ByVal chargeName As String, ByVal tmpDate As Date, ByVal curCol As Long) As Date
    
    If chargeName <> "" Then
        ' WBS�쐬��
        SetChargeNoWorkDay = GetChargeWorkDay(chargeName, tmpDate, 0)
    Else
        ' �J�����_�[�ݒ莞
        Dim confChargeColEnd As Long
        confChargeColEnd = shConf.Cells(C_CONFHEADER_ROW, Columns.Count).End(xlToLeft).Column
        
        Dim i As Long
        For i = ConvertNumAlp(C_CONF_CHARGENAME_COL) To confChargeColEnd
            ' �S���Җ����擾
            chargeName = shConf.Cells(C_CONFHEADER_ROW, i).Value
            
            Call GetChargeWorkDay(chargeName, tmpDate, curCol)
        Next
    
    End If
        
End Function


' �S���ҕʂ̔�ғ��j���A�����l�������ғ������Z�o
Function GetChargeWorkDay(ByVal chargeName As String, ByVal tmpDate As Date, ByVal curCol As Long) As Date
    ' chargeFromWbsRng�ɑ΂���FindNext�������Ȃ����߁AFor���[�v�őΉ�
    Dim maxRow As Long
    
     ' WBS����S���҂̍s���擾
    Dim chargeFromWbsRng, chargeFromWbsRngMax As Range
    Set chargeFromWbsRng = sh.Range(C_CHARGE_COL & ":" & C_CHARGE_COL).Find(chargeName, LookAt:=xlWhole)
    If chargeFromWbsRng Is Nothing Then
        ' �I�I�I�v�������I�I�I
        GetChargeWorkDay = tmpDate
        Exit Function
    Else
        Set chargeFromWbsRngMax = sh.Range(C_CHARGE_COL & ":" & C_CHARGE_COL).FindPrevious(chargeFromWbsRng)
        maxRow = chargeFromWbsRngMax.Row
    End If

    Dim i As Long
    For i = C_HEADER_ROW + 1 To maxRow
        If chargeName <> sh.Cells(i, ConvertNumAlp(C_CHARGE_COL)).Value Then GoTo NextLoop
    
        ' �S���҂̔�ғ��j���ɊY�����Ȃ����`�F�b�N

        ' config�V�[�g����S���Ҍ���
        Dim chargeNameRng As Range
        Set chargeNameRng = shConf.Range(C_CONFHEADER_ROW & ":" & C_CONFHEADER_ROW).Find(chargeName, LookAt:=xlWhole)
        
        If chargeNameRng Is Nothing Then
            ' �I�I�I�v�������I�I�I
            GetChargeWorkDay = tmpDate
            Exit Function
        End If

        ' �S���҂̔�ғ��j���`�F�b�N
        Dim chargeNoWorkWeekList As Variant
        Dim chargeNoWorkWeekStr As String
        chargeNoWorkWeekStr = shConf.Cells(C_CONF_CHARGENOWKWEEK_ROW, chargeNameRng.Column).Value
        chargeNoWorkWeekList = Split(chargeNoWorkWeekStr, ",")
        
        Dim s As Integer
        For s = LBound(chargeNoWorkWeekList) To UBound(chargeNoWorkWeekList)
            ' tmpDate�̗j�����擾
            Dim tmpDateWeekDay As String
            tmpDateWeekDay = ConvWeekDayJp(tmpDate)
            If tmpDateWeekDay = chargeNoWorkWeekList(s) Then
                ' �S���҂̔�ғ��j���ɊY������
                If curCol <> 0 Then
                    ' �Y���s�̃Z���̐F���O���[�ɂ���
                    sh.Cells(i, curCol).Interior.Color = C_NOWORKDAY_COLOR
                Else
                    ' ���t���Z
                    tmpDate = IIf(chargeName <> "", DateAdd("d", 1, tmpDate), tmpDate)
                End If
            End If
        Next s
    
        ' �S���҂̔�ғ����`�F�b�N
        Dim chargeNoWorkDayRng As Range
        Set chargeNoWorkDayRng = shConf.Range(ConvertNumAlp(chargeNameRng.Column) & ":" & _
                                        ConvertNumAlp(chargeNameRng.Column)).Find(tmpDate, LookAt:=xlWhole)
        Do
            ' �S���҂̔�ғ����ɊY�����Ȃ����`�F�b�N
            If chargeNoWorkDayRng Is Nothing Then Exit Do
            
            ' �S���҂̔�ғ����ɊY������
            If curCol <> 0 Then
                ' �Y���s�̃Z���̐F���O���[�ɂ���
                sh.Cells(i, curCol).Interior.Color = C_NOWORKDAY_COLOR
            Else
                ' ���t���Z
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
        
        ' WBS�쐬���A1���[�v�ŏI��
        If curCol = 0 Then Exit For
NextLoop:
        
    Next
    
    GetChargeWorkDay = tmpDate

End Function
