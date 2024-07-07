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
    
    ' �t�H���_�[�I���_�C�A���O
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�u���p�摜�t�H���_��I�����Ă�������"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "�t�H���_�[���I������܂���ł����B�����𒆎~���܂��B"
            Exit Sub
        End If
    End With
    
    ' ���[�U�[�ɗD�揇�ʂ��m�F
    isPositionPriority = (MsgBox("�摜�̒u���������ʒu�D��ɂ��܂����H" & vbNewLine & _
                                 "�u�͂��v�F�G�N�Z����̔z�u��" & vbNewLine & _
                                 "�u�������v�F�G�N�Z���ɔz�u�����^�C�~���O��", _
                                 vbYesNo) = vbYes)
    
    Set ws = ActiveSheet
    Set imageInfoCollection = New Collection
    imageIndex = 1
    unknownFormatCount = 0
    
    ' �摜�������W
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
    
    ' �ʒu�D��̏ꍇ�A���ォ��E���̏��Ƀ\�[�g
    If isPositionPriority Then
        Set imageInfoCollection = SortImagesByPosition(imageInfoCollection)
    End If
    
    ' �摜�u������
    For Each imageInfo In imageInfoCollection
        Set shp = imageInfo("Shape")
        
        ' �V�����摜��T��
        newImagePath = FindMatchingImageFile(folderPath, imageIndex)
        
        If newImagePath <> "" Then
            ' �摜��u��
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
    
    ' ���ʂ�\��
    Dim resultMessage As String
    resultMessage = "�摜�̒u�����������܂����B" & vbNewLine & _
                    "�u�����ꂽ�摜�̐�: " & (imageIndex - 1) & vbNewLine
    
    If unknownFormatCount > 0 Then
        resultMessage = resultMessage & "���Ή��̌`���܂��͌�����Ȃ������摜�̐�: " & unknownFormatCount & vbNewLine & _
                        "�����̉摜�͒u������Ă��܂���B"
    End If
    
    MsgBox resultMessage, vbInformation
End Sub

Function SortImagesByPosition(imageCollection As Collection) As Collection
    Dim sortedCollection As Collection
    Set sortedCollection = New Collection
    
    ' �R���N�V������z��ɕϊ�
    Dim arr() As Variant
    ReDim arr(1 To imageCollection.Count)
    Dim i As Long
    For i = 1 To imageCollection.Count
        Set arr(i) = imageCollection(i)
    Next i
    
    ' �o�u���\�[�g�ňʒu�Ń\�[�g
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
    
    ' �\�[�g���ꂽ�z����R���N�V�����ɖ߂�
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
    
    ' �t�@�C�����p�^�[�����`�i�K�v�ɉ����Ēǉ��j
    Dim patterns As Variant
    patterns = Array("image" & index & ".*", _
                     "image_" & Format(index, "0000") & ".*", _
                     "*_�摜_" & Format(index, "0000") & ".*", _
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

