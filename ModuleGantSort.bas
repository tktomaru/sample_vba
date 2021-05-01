Attribute VB_Name = "ModuleGantSort"
Sub InsertionSortType(ByRef data() As Point, ByVal low As Long, ByVal high As Long)

    Dim i As Variant
    Dim k As Variant
    Dim temp As Point

    For i = low + 1 To high
        temp = data(i)
        If ComparePoint(data(i - 1), temp) > 0 Then
            k = i

            Do While k > low
                If ComparePoint(data(k - 1), temp) <= 0 Then
                    Exit Do
                End If

                data(k) = data(k - 1)
                k = k - 1
            Loop

            data(k) = temp
        End If
    Next
    
End Sub

' 構造体の比較用関数
Function ComparePoint(ByRef data1 As Point, ByRef data2 As Point) As Variant
    ComparePoint = data1.date - data2.date
End Function

' 標準モジュールに定義
Public Type Point
    date As Date
End Type


