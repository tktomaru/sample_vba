Attribute VB_Name = "エクスポート"
Sub Export()
    Set wb = ThisWorkbook
    Dim sh As Worksheet
    Set sh = wb.Sheets(C_WBS_SHNM)
    
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename(InitialFileName:="WBS_" & Format(Now, "yyyymmdd") & ".xlsx", FileFilter:="Excelファイル (*.xlsx),*.xlsx*")
    If fileName = False Then
        MsgBox "保存に失敗しました。処理を終了します。"
        Exit Sub
    Else
        Dim tmpSh As Worksheet
        Dim tmpShName As String: tmpShName = "WBS_" & Format(Now, "yyyymmdd")
        
        ' WBSを別シートにコピー
        sh.Copy After:=wb.Sheets(C_WBS_SHNM)
        Set tmpSh = ActiveSheet
        tmpSh.Name = tmpShName
        
        ' マクロボタン削除
        Dim Btn As Object
        For Each Btn In tmpSh.Buttons
            Btn.Delete
        Next Btn
        
        ' 新ブックにシートコピー
        tmpSh.Copy
        ' シート削除
        Application.DisplayAlerts = False
        tmpSh.Delete
        Application.DisplayAlerts = True
        
        ' 指定したファイル名で保存
        ActiveWorkbook.SaveAs fileName
        ActiveWorkbook.Close
        
        ' 元のWBSにフォーカスを合わせる
        sh.Select
        
        MsgBox "エクスポートが完了しました。"
        
        Set tmpSh = Nothing
    End If
    
    Set sh = Nothing
    Set wb = Nothing

End Sub
