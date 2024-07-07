Attribute VB_Name = "Module2"
Sub EnhancedRegroupShapes()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim groupInfo As Collection
    Dim groupItem As Dictionary
    Dim groupMembers As Collection
    Dim memberShape As Shape
    Dim groupedShape As Shape
    
    Set ws = ActiveSheet
    Set groupInfo = New Collection
    
    ' グループ情報を収集
    For Each shp In ws.shapes
        If shp.Type = msoGroup Then
            Set groupItem = New Dictionary
            groupItem.Add "Left", shp.Left
            groupItem.Add "Top", shp.Top
            groupItem.Add "Width", shp.width
            groupItem.Add "Height", shp.height
            
            Set groupMembers = New Collection
            For Each memberShape In shp.GroupItems
                groupMembers.Add Array(memberShape.Left - shp.Left, memberShape.Top - shp.Top, memberShape.width, memberShape.height, memberShape.Type)
            Next memberShape
            
            groupItem.Add "Members", groupMembers
            groupInfo.Add groupItem
        End If
    Next shp
    
    ' ユーザーに確認
    If MsgBox("画像の確認と調整が完了しましたか？再グループ化を開始しますか？", vbQuestion + vbYesNo) = vbNo Then
        MsgBox "再グループ化をキャンセルしました。画像の確認と調整を完了してから再度実行してください。", vbInformation
        Exit Sub
    End If
    
    ' 再グループ化
    Dim leftPos As Single, topPos As Single
    Dim width As Single, height As Single
    Dim memberInfo As Variant
    Dim shapesToGroup As Collection
    Dim groupCount As Integer
    
    groupCount = 0
    For Each groupItem In groupInfo
        Set shapesToGroup = New Collection
        leftPos = groupItem("Left")
        topPos = groupItem("Top")
        width = groupItem("Width")
        height = groupItem("Height")
        
        For Each memberInfo In groupItem("Members")
            For Each shp In ws.shapes
                If shp.Type = memberInfo(4) And _
                   IsShapeInRange(shp, leftPos + memberInfo(0), topPos + memberInfo(1), CSng(memberInfo(2)), CSng(memberInfo(3))) Then
                    shapesToGroup.Add shp
                    Exit For
                End If
            Next shp
        Next memberInfo
        
        If shapesToGroup.Count > 1 Then
            ' グループ化をプレビュー
            HighlightShapes shapesToGroup
            If MsgBox("これらの形状をグループ化しますか？", vbQuestion + vbYesNo) = vbYes Then
                Set groupedShape = ws.shapes.Range(ShapeNamesToArray(shapesToGroup)).Group
                groupCount = groupCount + 1
            End If
            UnhighlightShapes shapesToGroup
        End If
    Next groupItem
    
    MsgBox "再グループ化が完了しました。グループ化された形状の数: " & groupCount, vbInformation
End Sub

Function IsShapeInRange(shp As Shape, expectedLeft As Single, expectedTop As Single, expectedWidth As Single, expectedHeight As Single) As Boolean
    Const TOLERANCE_FACTOR As Single = 0.1 ' 10%の許容範囲
    
    Dim leftDiff As Single, topDiff As Single, widthDiff As Single, heightDiff As Single
    
    leftDiff = Abs(shp.Left - expectedLeft)
    topDiff = Abs(shp.Top - expectedTop)
    widthDiff = Abs(shp.width - expectedWidth)
    heightDiff = Abs(shp.height - expectedHeight)
    
    IsShapeInRange = (leftDiff <= expectedWidth * TOLERANCE_FACTOR) And _
                     (topDiff <= expectedHeight * TOLERANCE_FACTOR) And _
                     (widthDiff <= expectedWidth * TOLERANCE_FACTOR) And _
                     (heightDiff <= expectedHeight * TOLERANCE_FACTOR)
End Function

Sub HighlightShapes(shapes As Collection)
    Dim shp As Shape
    For Each shp In shapes
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGB(255, 0, 0) ' 赤色
        shp.Line.Weight = 2
    Next shp
End Sub

Sub UnhighlightShapes(shapes As Collection)
    Dim shp As Shape
    For Each shp In shapes
        shp.Line.Visible = msoFalse
    Next shp
End Sub

Function ShapeNamesToArray(shapes As Collection) As Variant
    Dim result() As String
    ReDim result(1 To shapes.Count)
    Dim i As Integer
    For i = 1 To shapes.Count
        result(i) = shapes(i).Name
    Next i
    ShapeNamesToArray = result
End Function

