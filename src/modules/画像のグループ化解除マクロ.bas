Attribute VB_Name = "Module1"
Sub UngroupAllObjects()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim grp As GroupShapes
    Dim i As Long
    
    ' �A�N�e�B�u�ȃ��[�N�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' ���[�N�V�[�g���̂��ׂĂ̐}�`�����[�v
    For Each shp In ws.Shapes
        ' �I�u�W�F�N�g���O���[�v�̏ꍇ
        If shp.Type = msoGroup Then
            ' �O���[�v���̃I�u�W�F�N�g�����擾
            Set grp = shp.GroupItems
            i = grp.Count
            
            ' �O���[�v������
            shp.Ungroup
            
            ' �J�E���^�[�𒲐��i�������ꂽ�O���[�v�̕������߂��j
            i = i - 1
        End If
    Next shp
    
    MsgBox "���ׂẴI�u�W�F�N�g�̃O���[�v������������܂����B", vbInformation
End Sub
