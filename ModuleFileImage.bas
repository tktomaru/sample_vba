Attribute VB_Name = "ModuleFileImage"

Function ImportPicture(importFile As String, pasetRangeText As String, _
                       Optional width As Integer = 0, Optional height As Integer = 0) As Integer
    Dim myFileName As String
    Dim myShape As Shape
    Dim paste As range
    
    Set paste = range(pasetRangeText)
    
    importFile = GetAbsolutePathNameEx(ThisWorkbook.Path, importFile)
    myFileName = importFile
    
    '--(1) 選択位置に画像ファイルを挿入し、変数myShapeに格納
    Set myShape = ActiveSheet.Shapes.AddPicture( _
          fileName:=myFileName, _
          LinkToFile:=False, _
          SaveWithDocument:=True, _
          Left:=paste.Left, _
          Top:=paste.Top, _
          width:=width, _
          height:=height)
          
    '--(2) 挿入した画像に対して元画像と同じ高さ・幅にする
    If width = 0 Or height = 0 Then
       With myShape
        If height = 0 Then
           .ScaleHeight 1, msoTrue
        End If
        If width = 0 Then
           .ScaleWidth 1, msoTrue
        End If
       End With
    End If
    
    ImportPicture = myShape.height
End Function

Sub アクティブシートの画像をすべて削除する()
  Dim shp As Shape
  For Each shp In ActiveSheet.Shapes
    If shp.Type = msoPicture Then shp.Delete
  Next shp
End Sub
