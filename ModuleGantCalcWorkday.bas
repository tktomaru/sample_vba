Attribute VB_Name = "ModuleGantCalcWorkday"


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
    Set WS = Worksheets("�H���\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("�H���\Config")
    Dim priMax As Integer
    
    Set nameC = CreateObject("Scripting.Dictionary")
    Set nameCSum = CreateObject("Scripting.Dictionary")
    
    taskDbl = WS.Range("K5")
    startDate = WS.Range("E2")
  
    row = 5
    priNString = "N"
    nameA = "P"
        
    ' �D�揇�ʂ̍ő�l
    priMax = WorksheetFunction.Max(Range("N5", "N500"))
      
    ' �S���̊���o��
    For ii = row To WS.Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(WS.Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
           nameCSum.Add tmp, 0
        End If
    Next ii
    
    ' �D��x���ɌJ��Ԃ��D�揇�ʂ����߂�
    For ii = 1 To priMax
    
       For jj = row To 500
    
          If ("" = WS.Cells(jj, "N")) Then
             GoTo LFEND
          End If
       
          ' �T�����̗D��x�ƈ�v���邩
          If (ii = Cells(jj, priNString)) Then
             ' �ݒ�V�[�g�̃J�����ԍ����擾
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
    
          ' ���O���猻�݂̍�Ɠ������擾
          Dim task As Double
          task = nameC.item(tmp) + WS.Cells(jj, "K")
          Cells(jj, "L") = task
           
           Cells(jj, "D") = startDate + nameCSum(tmp)
               
          ' �J�n������̓������Z�o
          Dim sumtask As Double
          sumtask = calcWorkday(startDate, task, _
                     holidayDate, _
                     youbiInt, _
                     personalDate)
                     
           'task = sumtask + Cells(jj, "K")
           nameC(tmp) = task
           nameCSum(tmp) = sumtask - WS.Cells(jj, "K")
           Cells(jj, "E") = startDate + sumtask - WS.Cells(jj, "K")
          End If
       Next jj
LFEND:
    Next ii

End Sub

' �S���̏��O���ȊO�̏j����Ԃ�
Function conbertRangeToDateWithout(inputRange As Range, name As String) As Date()
    Dim WS As Worksheet
    Set WS = Worksheets("�H���\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("�H���\Config")
    
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

Function convertYoubi(youbi As String) As Integer()
   Dim ret() As Integer
   Dim tmp As Variant
   Dim retNum As Integer
   
   retNum = 0
   tmp = Split(youbi, ",")

'�O�FvbUseSystemDayOfWeek(PC��OS�̃V�X�e������)
'�P�FvbSunday (���j��)
'�Q�FvbMonday (���j��)
'�R�FvbTuesday (�Ηj��)
'�S�FvbWednesday (���j��)
'�T�FvbThursday (�ؗj��)
'�U�FvbFriday (���j��)
'�V�FvbSaturday (�y�j��)

    For ii = LBound(tmp) To UBound(tmp)
       Select Case tmp(ii)
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 1
           retNum = retNum + 1
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 2
           retNum = retNum + 1
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 3
           retNum = retNum + 1
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 4
           retNum = retNum + 1
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 5
           retNum = retNum + 1
       Case "��"
           ReDim Preserve ret(retNum)
           ret(retNum) = 6
           retNum = retNum + 1
       Case "�y"
           ReDim Preserve ret(retNum)
           ret(retNum) = 7
           retNum = retNum + 1
       End Select
    Next ii
    
    convertYoubi = ret
End Function

Sub calcPriority()
    Dim WS As Worksheet
    Set WS = Worksheets("�H���\")
    
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
    
    ' �S���̊���o��
    For ii = row To Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
        End If
    Next ii
        
    ' �D��x�̍ő�l
    Set rng = Range("M5", "M500")
    priMax = WorksheetFunction.Max(rng)
      
    ' �D��x���ɌJ��Ԃ��D�揇�ʂ����߂�
    For ii = 1 To priMax
       For jj = row To 500
          ' �T�����̗D��x�ƈ�v���邩
          If (ii = Cells(jj, CNumAlp(priA))) Then
             ' ���O���猻�݂̍�Ɠ������擾
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


Sub YoubiColor()
    Dim holidayRange As Range
    Dim holidayDate() As Date
    
    Dim youbiInt() As Integer
    Dim youbiString As String
    Dim personalDate() As Date
    
    Dim WS As Worksheet
    Set WS = Worksheets("�H���\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("�H���\Config")
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
       
             ' �ݒ�V�[�g�̃J�����ԍ����擾
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
          
       ' �s�̐F����
       For kk = CNumAlp("R") To colMax
         
         keikakuDate = WS.Cells(3, kk)
      
         ' �F�w����N���A�i�����w��j
         WS.Cells(jj, kk).Interior.Color = RGB(255, 255, 255) ' �w�i�F
         
         ' �j��
         If (True = getHoliday(holidayDate, youbiInt, personalDate, keikakuDate)) Then
                WS.Cells(jj, kk).Interior.Color = RGB(255, 200, 200) ' �w�i�F���s���N�ɂ���
         End If
         
       Next kk
LOOPEND:
    Next jj
End Sub

Function getHoliday(holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date, keikakuDate As Date) As Boolean
    Dim WS As Worksheet
    Set WS = Worksheets("�H���\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("�H���\Config")
    Dim ii As Integer

         ' �j��
         For ii = LBound(holidayDate) To UBound(holidayDate)
           ' �j���ƈ�v���邩�H
            If ((keikakuDate) = holidayDate(ii)) Then
               ' ��v���Ă�����
               getHoliday = True
               GoTo LOOPEND
            End If
         Next ii
   
         ' �l�̋x��
         For ii = LBound(personalDate) To UBound(personalDate)
           ' �l�̋x���ƈ�v���邩�H
            If ((keikakuDate) = personalDate(ii)) Then
               ' ��v���Ă�����
               getHoliday = True
               GoTo LOOPEND
            End If
         Next ii
   
         ' �l�̗j��
         For ii = LBound(youbiInt) To UBound(youbiInt)
           ' �l�̗j���ƈ�v���邩�H
            If (Weekday(keikakuDate) = youbiInt(ii)) Then
               ' ��v���Ă�����
               getHoliday = True
               GoTo LOOPEND
            End If
         Next ii
LOOPEND:
End Function

' startDate    �J�n��
' taskDbl      ��Ɠ���
' holidayDate  �j��
' workDate     �ғ��j��
' youbiInt     ��ғ��j��
' personalDate ��ғ���
Function calcWorkday(startDate As Date, taskDbl As Double, _
                     holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date) As Integer

   Dim startDayInt As Integer
   Dim endDayInt As Integer
   '�؂�̂�
   startDayInt = taskDbl
   ' �l�̌ܓ�
   endDayInt = WorksheetFunction.RoundUp(taskDbl, 0)
   Dim tmpTask  As Integer
   
   
   tmpTask = 0
   jj = 0
   
   Do While tmpTask < endDayInt
      Dim ii As Integer
      
         ' �j��
         If (True = getHoliday(holidayDate, youbiInt, personalDate, startDate + tmpTask)) Then
            ' ��v���Ă�����ғ��ŏI������������
            endDayInt = endDayInt + 1
            GoTo LOOPEND
         End If
         
LOOPEND:
      tmpTask = tmpTask + 1
   Loop
   
   calcWorkday = endDayInt

End Function
