Attribute VB_Name = "Module1"
Sub SendImagesToBack()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim imageCount As Integer
    
    ' �A�N�e�B�u�ȃ��[�N�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    imageCount = 0
    
    ' ���[�N�V�[�g���̂��ׂĂ̐}�`�����[�v
    For Each shp In ws.Shapes
        ' �}�`���摜�̏ꍇ
        If shp.Type = msoPicture Then
            ' �摜���Ŕw�ʂɈړ�
            shp.ZOrder msoSendToBack
            imageCount = imageCount + 1
        End If
    Next shp
    
    MsgBox imageCount & "�̉摜���Ŕw�ʂɈړ����܂����B", vbInformation
End Sub
