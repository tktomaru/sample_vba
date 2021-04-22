Attribute VB_Name = "行追加削除"
Sub AddRow()
    Dim actRow As Long
    actRow = Selection.Row
    
    Dim sh As Worksheet
    Set sh = Worksheets(C_WBS_SHNM)
    
    Dim maxRow As Long
    maxRow = sh.Range("A:A").End(xlUp).Row
    If maxRow < actRow And C_HEADER_ROW > actRow Then
        MsgBox "選択されている位置では行追加できません。処理を終了します。"
        GoTo Finally
    End If
    
    sh.Rows(actRow).Copy
    
    sh.Rows(actRow).Insert
    sh.Rows(actRow + 1).PasteSpecial
    Application.CutCopyMode = False ' コピーモードを解除

Finally:
    Set sh = Nothing

End Sub

Sub DeleteRow()
    Dim actRow As Long
    actRow = Selection.Row

    Dim sh As Worksheet
    Set sh = Worksheets(C_WBS_SHNM)

    Dim maxRow As Long
    maxRow = sh.Range("A:A").End(xlUp).Row
    If maxRow < actRow And C_HEADER_ROW >= actRow Then
        MsgBox "選択されている位置では行削除できません。処理を終了します。"
        GoTo Finally
    End If
    
    sh.Rows(actRow).Delete

Finally:
    Set sh = Nothing

End Sub

