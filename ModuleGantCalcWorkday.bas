Attribute VB_Name = "ModuleGantCalcWorkday"

' ŠJn“ú‚Æì‹Æ“ú”‚Æ—Dæ‡ˆÊ‚©‚çŒv‰æ‚ğ—§‚Ä‚é
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
    Set WS = Worksheets("H’ö•\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("H’ö•\Config")
    Dim priMax As Integer
    
    Set nameC = CreateObject("Scripting.Dictionary")
    Set nameCSum = CreateObject("Scripting.Dictionary")
    
    taskDbl = WS.Range("K5")
    startDate = WS.Range("E2")
  
    row = 5
    priNString = "N"
    nameA = "P"
        
    ' —Dæ‡ˆÊ‚ÌÅ‘å’l
    priMax = WorksheetFunction.Max(Range("N5", "N500"))
      
    ' ’S“–‚ÌŠ„‚èo‚µ
    For ii = row To WS.Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(WS.Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
           nameCSum.Add tmp, 0
        End If
    Next ii
    
    ' —Dæ“x‡‚ÉŒJ‚è•Ô‚µ—Dæ‡ˆÊ‚ğ‹‚ß‚é
    For ii = 1 To priMax
    
       For jj = row To 500
    
          If ("" = WS.Cells(jj, "N")) Then
             GoTo LFEND
          End If
       
          ' ’Tõ’†‚Ì—Dæ“x‚Æˆê’v‚·‚é‚©
          If (ii = Cells(jj, priNString)) Then
             ' İ’èƒV[ƒg‚ÌƒJƒ‰ƒ€”Ô†‚ğæ“¾
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
    
          ' –¼‘O‚©‚çŒ»İ‚Ìì‹Æ“ú”‚ğæ“¾
          Dim task As Double
           
           
          ' ŠJn“ú‚©‚ç‚Ì“ú”‚ğZo
          Dim sumtask As Double
           
           task = nameC.item(tmp)
          ' ŠJn“ú‚©‚ç‚Ì“ú”‚ğZoiŠJn“új
          sumtask = calcWorkday(startDate, task, _
                     holidayDate, _
                     youbiInt, _
                     personalDate)
           Cells(jj, "D") = startDate + sumtask
                     
          task = nameC.item(tmp) + WS.Cells(jj, "K")
          Cells(jj, "L") = task
           nameC(tmp) = task
               
          ' ŠJn“ú‚©‚ç‚Ì“ú”‚ğZoiI—¹“új
          sumtask = calcWorkday(startDate, task, _
                     holidayDate, _
                     youbiInt, _
                     personalDate)
                     
           nameCSum(tmp) = sumtask
           ' ŠJn“úˆÈŠOA‚©‚ÂA ®”‚Ìê‡‚É‚Í‚¿‚å‚¤‚Ç‚»‚Ì“ú‚Éƒ^ƒXƒN‚ªû‚Ü‚é‚½‚ßA‰ÁZ‚ğ‚È‚­‚·
           If (0 <> sumtask And (sumtask = Int(sumtask))) Then
              sumtask = sumtask - 1
           End If
           Cells(jj, "E") = startDate + sumtask
          End If
       Next jj
LFEND:
    Next ii

End Sub

' ’S“–‚ÌœŠO“úˆÈŠO‚Ìj“ú‚ğ•Ô‚·
Function conbertRangeToDateWithout(inputRange As Range, name As String) As Date()
    Dim WS As Worksheet
    Set WS = Worksheets("H’ö•\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("H’ö•\Config")
    
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

' Range‚©‚çDate”z—ñ‚É•ÏŠ·‚·‚é
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

' —áF"Œ,‰Î"‚Æ‚¢‚¤•¶š—ñ‚ğ—^‚¦‚é‚ÆAret=[2,3]‚Æ‚¢‚¤Integer”z—ñ‚Å•Ô‚·
Function convertYoubi(youbi As String) As Integer()
   Dim ret() As Integer
   Dim tmp As Variant
   Dim retNum As Integer
   
   retNum = 0
   tmp = Split(youbi, ",")

'‚OFvbUseSystemDayOfWeek(PC‚ÌOS‚ÌƒVƒXƒeƒ€ŠÔ)
'‚PFvbSunday (“ú—j“ú)
'‚QFvbMonday (Œ—j“ú)
'‚RFvbTuesday (‰Î—j“ú)
'‚SFvbWednesday (…—j“ú)
'‚TFvbThursday (–Ø—j“ú)
'‚UFvbFriday (‹à—j“ú)
'‚VFvbSaturday (“y—j“ú)

    For ii = LBound(tmp) To UBound(tmp)
       Select Case tmp(ii)
       Case "“ú"
           ReDim Preserve ret(retNum)
           ret(retNum) = 1
           retNum = retNum + 1
       Case "Œ"
           ReDim Preserve ret(retNum)
           ret(retNum) = 2
           retNum = retNum + 1
       Case "‰Î"
           ReDim Preserve ret(retNum)
           ret(retNum) = 3
           retNum = retNum + 1
       Case "…"
           ReDim Preserve ret(retNum)
           ret(retNum) = 4
           retNum = retNum + 1
       Case "–Ø"
           ReDim Preserve ret(retNum)
           ret(retNum) = 5
           retNum = retNum + 1
       Case "‹à"
           ReDim Preserve ret(retNum)
           ret(retNum) = 6
           retNum = retNum + 1
       Case "“y"
           ReDim Preserve ret(retNum)
           ret(retNum) = 7
           retNum = retNum + 1
       End Select
    Next ii
    
    convertYoubi = ret
End Function

Sub calcPriority()
    Dim WS As Worksheet
    Set WS = Worksheets("H’ö•\")
    
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
    
    ' ’S“–‚ÌŠ„‚èo‚µ
    For ii = row To Cells(Rows.Count, nameA).End(xlUp).row
        tmp = CStr(Cells(ii, nameA))
        If (False = nameC.Exists(tmp)) Then
           nameC.Add tmp, 0
        End If
    Next ii
        
    ' —Dæ“x‚ÌÅ‘å’l
    Set rng = Range("M5", "M500")
    priMax = WorksheetFunction.Max(rng)
      
    ' —Dæ“x‡‚ÉŒJ‚è•Ô‚µ—Dæ‡ˆÊ‚ğ‹‚ß‚é
    For ii = 1 To priMax
       For jj = row To 500
          ' ’Tõ’†‚Ì—Dæ“x‚Æˆê’v‚·‚é‚©
          If (ii = Cells(jj, CNumAlp(priA))) Then
             ' –¼‘O‚©‚çŒ»İ‚Ìì‹Æ“ú”‚ğæ“¾
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

' "R"—ñ‚æ‚è‰E‚Ì‹x“ú‚ğƒsƒ“ƒNF‚É‚·‚é
Sub YoubiColor()
    Dim holidayRange As Range
    Dim holidayDate() As Date
    
    Dim youbiInt() As Integer
    Dim youbiString As String
    Dim personalDate() As Date
    
    Dim WS As Worksheet
    Set WS = Worksheets("H’ö•\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("H’ö•\Config")
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
       
             ' İ’èƒV[ƒg‚ÌƒJƒ‰ƒ€”Ô†‚ğæ“¾
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
          
       ' s‚ÌF•ª‚¯
       For kk = CNumAlp("R") To colMax
         
         keikakuDate = WS.Cells(3, kk)
      
         ' Fw’è‚ğƒNƒŠƒAi”’‚ğw’èj
         WS.Cells(jj, kk).Interior.color = RGB(255, 255, 255) ' ”wŒiF
         
         ' j“ú
         If (True = isHoliday(holidayDate, youbiInt, personalDate, keikakuDate)) Then
                WS.Cells(jj, kk).Interior.color = RGB(255, 200, 200) ' ”wŒiF‚ğƒsƒ“ƒN‚É‚·‚é
         End If
         
       Next kk
LOOPEND:
    Next jj
End Sub

' keikakuDate ‚ªj“ú‚©‚Ç‚¤‚©‚ğ”»’è‚·‚é
' @param holidayDate j“ú
' @param youbiInt    ŒÂl‚Ì‹x‚Ş—j“ú
' @param personalDate ŒÂl‚Ì‹x“ú
' @return True=j“ú@False=•½“ú
Function isHoliday(holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date, keikakuDate As Date) As Boolean
    Dim WS As Worksheet
    Set WS = Worksheets("H’ö•\")
    Dim WSConfig As Worksheet
    Set WSConfig = Worksheets("H’ö•\Config")
    Dim ii As Integer
    Dim ret As Boolean
    

         ' j“ú
         If (CalcArrayLength(holidayDate) >= 1) Then
         For ii = LBound(holidayDate) To UBound(holidayDate)
           ' j“ú‚Æˆê’v‚·‚é‚©H
            If ((keikakuDate) = holidayDate(ii)) Then
               ' ˆê’v‚µ‚Ä‚¢‚½‚ç
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
   
         ' ŒÂl‚Ì‹x“ú
         If (CalcArrayLength(personalDate) >= 1) Then
         For ii = LBound(personalDate) To UBound(personalDate)
           ' ŒÂl‚Ì‹x“ú‚Æˆê’v‚·‚é‚©H
            If ((keikakuDate) = personalDate(ii)) Then
               ' ˆê’v‚µ‚Ä‚¢‚½‚ç
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
   
         ' ŒÂl‚Ì—j“ú
         If (CalcArrayLength(youbiInt) >= 1) Then
         For ii = LBound(youbiInt) To UBound(youbiInt)
           ' ŒÂl‚Ì—j“ú‚Æˆê’v‚·‚é‚©H
            If (Weekday(keikakuDate) = youbiInt(ii)) Then
               ' ˆê’v‚µ‚Ä‚¢‚½‚ç
               ret = True
               GoTo LOOPEND
            End If
         Next ii
         End If
         ret = False
LOOPEND:
    isHoliday = ret
End Function

' @param startDate    ŠJn“ú
' @param taskDbl      ì‹Æ“ú”
' @param holidayDate  j“ú
' @param youbiInt     ”ñ‰Ò“­—j“ú
' @param personalDate ”ñ‰Ò“­“ú
Function calcWorkday(startDate As Date, _
                     taskDbl As Double, _
                     holidayDate() As Date, _
                     youbiInt() As Integer, _
                     personalDate() As Date) As Double

   ' Dim startDayInt As Integer
   ' Dim endDayInt As Integer
   ' Ø‚èÌ‚Ä
   ' startDayInt = taskDbl
   ' Ø‚èã‚°
   ' endDayInt = WorksheetFunction.RoundUp(taskDbl, 0)
   Dim tmpTask  As Integer
   
   tmpTask = 0
   Do
         ' j“ú
         If (True = isHoliday(holidayDate, youbiInt, personalDate, startDate + tmpTask)) Then
            ' ‹x“ú‚È‚ç‚Î‰Ò“­ÅI“ú‚ğ‰„’·‚·‚é
            taskDbl = taskDbl + 1
            
         End If
         ' Œ»İ“ú‚ğ1“úŒo‰ß‚³‚¹‚é
         tmpTask = tmpTask + 1
   Loop While tmpTask < taskDbl
LFEND:
   calcWorkday = taskDbl

End Function

