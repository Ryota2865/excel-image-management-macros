Attribute VB_Name = "Module1"
Sub ListShapeTypes()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Integer
    
    Set ws = ActiveSheet
    i = 1
    
    For Each shp In ws.Shapes
        Debug.Print "�I�u�W�F�N�g #" & i & ": �^�C�v = " & GetShapeTypeName(shp.Type) & _
                    ", ���O = " & shp.Name
        i = i + 1
    Next shp
End Sub

Function GetShapeTypeName(shapeType As MsoShapeType) As String
    Select Case shapeType
        Case msoPicture
            GetShapeTypeName = "�摜"
        Case msoShape
            GetShapeTypeName = "�}�`"
        Case msoGroup
            GetShapeTypeName = "�O���[�v"
        Case msoTextBox
            GetShapeTypeName = "�e�L�X�g�{�b�N�X"
        Case Else
            GetShapeTypeName = "���̑� (" & shapeType & ")"
    End Select
End Function

