Attribute VB_Name = "�s�ǉ��폜"
Sub AddRow()
    Dim actRow As Long
    actRow = Selection.Row
    
    Dim sh As Worksheet
    Set sh = Worksheets(C_WBS_SHNM)
    
    Dim maxRow As Long
    maxRow = sh.Range("A:A").End(xlUp).Row
    If maxRow < actRow And C_HEADER_ROW > actRow Then
        MsgBox "�I������Ă���ʒu�ł͍s�ǉ��ł��܂���B�������I�����܂��B"
        GoTo Finally
    End If
    
    sh.Rows(actRow).Copy
    
    sh.Rows(actRow).Insert
    sh.Rows(actRow + 1).PasteSpecial
    Application.CutCopyMode = False ' �R�s�[���[�h������

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
        MsgBox "�I������Ă���ʒu�ł͍s�폜�ł��܂���B�������I�����܂��B"
        GoTo Finally
    End If
    
    sh.Rows(actRow).Delete

Finally:
    Set sh = Nothing

End Sub

