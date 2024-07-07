Attribute VB_Name = "Module1"
Sub UngroupAllObjects()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim grp As GroupShapes
    Dim i As Long
    
    ' アクティブなワークシートを設定
    Set ws = ActiveSheet
    
    ' ワークシート内のすべての図形をループ
    For Each shp In ws.Shapes
        ' オブジェクトがグループの場合
        If shp.Type = msoGroup Then
            ' グループ内のオブジェクト数を取得
            Set grp = shp.GroupItems
            i = grp.Count
            
            ' グループを解除
            shp.Ungroup
            
            ' カウンターを調整（解除されたグループの分だけ戻す）
            i = i - 1
        End If
    Next shp
    
    MsgBox "すべてのオブジェクトのグループ化が解除されました。", vbInformation
End Sub
