Attribute VB_Name = "Module1"
Sub AdvancedFlexibleReplaceImages()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim folderPath As String
    Dim newImagePath As String
    Dim imageIndex As Integer
    Dim unknownFormatCount As Integer
    Dim isPositionPriority As Boolean
    Dim imageInfoCollection As Collection
    Dim imageInfo As Object
    
    ' フォルダー選択ダイアログ
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "置換用画像フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダーが選択されませんでした。処理を中止します。"
            Exit Sub
        End If
    End With
    
    ' ユーザーに優先順位を確認
    isPositionPriority = (MsgBox("画像の置換順序を位置優先にしますか？" & vbNewLine & _
                                 "「はい」：エクセル上の配置順" & vbNewLine & _
                                 "「いいえ」：エクセルに配置したタイミング順", _
                                 vbYesNo) = vbYes)
    
    Set ws = ActiveSheet
    Set imageInfoCollection = New Collection
    imageIndex = 1
    unknownFormatCount = 0
    
    ' 画像情報を収集
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            Set imageInfo = CreateObject("Scripting.Dictionary")
            imageInfo.Add "Shape", shp
            imageInfo.Add "Left", shp.Left
            imageInfo.Add "Top", shp.Top
            imageInfo.Add "Width", shp.Width
            imageInfo.Add "Height", shp.Height
            imageInfoCollection.Add imageInfo
        End If
    Next shp
    
    ' 位置優先の場合、左上から右下の順にソート
    If isPositionPriority Then
        Set imageInfoCollection = SortImagesByPosition(imageInfoCollection)
    End If
    
    ' 画像置換処理
    For Each imageInfo In imageInfoCollection
        Set shp = imageInfo("Shape")
        
        ' 新しい画像を探す
        newImagePath = FindMatchingImageFile(folderPath, imageIndex)
        
        If newImagePath <> "" Then
            ' 画像を置換
            shp.Delete
            ws.Shapes.AddPicture fileName:=newImagePath, _
                LinkToFile:=False, SaveWithDocument:=True, _
                Left:=imageInfo("Left"), Top:=imageInfo("Top"), _
                Width:=imageInfo("Width"), Height:=imageInfo("Height")
            imageIndex = imageIndex + 1
        Else
            unknownFormatCount = unknownFormatCount + 1
        End If
    Next imageInfo
    
    ' 結果を表示
    Dim resultMessage As String
    resultMessage = "画像の置換が完了しました。" & vbNewLine & _
                    "置換された画像の数: " & (imageIndex - 1) & vbNewLine
    
    If unknownFormatCount > 0 Then
        resultMessage = resultMessage & "未対応の形式または見つからなかった画像の数: " & unknownFormatCount & vbNewLine & _
                        "これらの画像は置換されていません。"
    End If
    
    MsgBox resultMessage, vbInformation
End Sub

Function SortImagesByPosition(imageCollection As Collection) As Collection
    Dim sortedCollection As Collection
    Set sortedCollection = New Collection
    
    ' コレクションを配列に変換
    Dim arr() As Variant
    ReDim arr(1 To imageCollection.Count)
    Dim i As Long
    For i = 1 To imageCollection.Count
        Set arr(i) = imageCollection(i)
    Next i
    
    ' バブルソートで位置でソート
    Dim j As Long
    Dim temp As Object
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)("Top") > arr(j)("Top") Or _
               (arr(i)("Top") = arr(j)("Top") And arr(i)("Left") > arr(j)("Left")) Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i
    
    ' ソートされた配列をコレクションに戻す
    For i = LBound(arr) To UBound(arr)
        sortedCollection.Add arr(i)
    Next i
    
    Set SortImagesByPosition = sortedCollection
End Function

Function FindMatchingImageFile(folderPath As String, index As Integer) As String
    Dim fileName As String
    Dim filePath As String
    Dim supportedFormats As Variant
    Dim imgFormat As Variant
    
    supportedFormats = Array("png", "jpg", "jpeg", "tiff", "tif", "bmp", "gif")
    
    ' ファイル名パターンを定義（必要に応じて追加）
    Dim patterns As Variant
    patterns = Array("image" & index & ".*", _
                     "image_" & Format(index, "0000") & ".*", _
                     "*_画像_" & Format(index, "0000") & ".*", _
                     "*_" & index & ".*", _
                     index & ".*")
    
    Dim pattern As Variant
    For Each pattern In patterns
        For Each imgFormat In supportedFormats
            fileName = Dir(folderPath & pattern)
            Do While fileName <> ""
                If LCase(Right(fileName, Len(CStr(imgFormat)))) = LCase(imgFormat) Then
                    FindMatchingImageFile = folderPath & fileName
                    Exit Function
                End If
                fileName = Dir()
            Loop
        Next imgFormat
    Next pattern
    
    FindMatchingImageFile = ""
End Function

