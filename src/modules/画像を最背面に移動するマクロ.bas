Attribute VB_Name = "Module1"
Sub SendImagesToBack()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim imageCount As Integer
    
    ' アクティブなワークシートを設定
    Set ws = ActiveSheet
    
    imageCount = 0
    
    ' ワークシート内のすべての図形をループ
    For Each shp In ws.Shapes
        ' 図形が画像の場合
        If shp.Type = msoPicture Then
            ' 画像を最背面に移動
            shp.ZOrder msoSendToBack
            imageCount = imageCount + 1
        End If
    Next shp
    
    MsgBox imageCount & "個の画像を最背面に移動しました。", vbInformation
End Sub
