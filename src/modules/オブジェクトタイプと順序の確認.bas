Attribute VB_Name = "Module1"
Sub ListShapeTypes()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Integer
    
    Set ws = ActiveSheet
    i = 1
    
    For Each shp In ws.Shapes
        Debug.Print "オブジェクト #" & i & ": タイプ = " & GetShapeTypeName(shp.Type) & _
                    ", 名前 = " & shp.Name
        i = i + 1
    Next shp
End Sub

Function GetShapeTypeName(shapeType As MsoShapeType) As String
    Select Case shapeType
        Case msoPicture
            GetShapeTypeName = "画像"
        Case msoShape
            GetShapeTypeName = "図形"
        Case msoGroup
            GetShapeTypeName = "グループ"
        Case msoTextBox
            GetShapeTypeName = "テキストボックス"
        Case Else
            GetShapeTypeName = "その他 (" & shapeType & ")"
    End Select
End Function

