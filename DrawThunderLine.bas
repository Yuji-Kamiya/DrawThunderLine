Attribute VB_Name = "DrawThunderLine"
Public Const SHEET_WBS_NAME = "�V�[�g��"    ' WBS�V�[�g��
Public Const LINE_NAME = "ThunderLine"      ' �C�i�Y�}���̖���
Public Const LINE_LENGTH = 50               ' �C�i�Y�}���̒���
Public Const DATE_LINE = 5                  ' ���t�s
Public Const DATE_ROW = 15                  ' ���t��

Sub DrawThunderLine()
    Dim i, j
    Dim line As Shape                         ' �C�i�Y�}��
    Dim start_x As Single, start_y As Single  ' �J�n���W
    Dim end_x As Single, end_y As Single      ' �I�����W
    Dim top_x As Single, top_y As Single      ' ���_���W
    Dim delay_line                            ' �x���i�O�|���j�s
    Dim delay_days                            ' �x���i�O�|���j����
    Dim ffb As FreeformBuilder
    
    '��Ȑ����폜����
    For Each s In Sheets(SHEET_WBS_NAME).Shapes
        If InStr(s.Name, LINE_NAME) > 0 Then
            s.Delete
        End If
    Next
    
    
    ' ���W��ݒ�
    start_x = Sheets(SHEET_WBS_NAME).Cells(DATE_LINE, DATE_ROW).Left
    start_y = Sheets(SHEET_WBS_NAME).Cells(DATE_LINE, DATE_ROW).Top
    end_x = start_x
    end_y = Sheets(SHEET_WBS_NAME).Cells(LINE_LENGTH, DATE_ROW).Top

    '
    Set ffb = Sheets(SHEET_WBS_NAME).Shapes.BuildFreeform(msoEditingCorner, start_x, start_y)
    ffb.AddNodes msoSegmentLine, msoEditingCorner, end_x, end_y
    Set progress = ffb.ConvertToShape
    
    ' ����������
    With progress.line
        .DashStyle = msoLineSolid       ' �X�^�C��
        .Weight = 3.5                   ' ����
        .ForeColor.RGB = RGB(255, 0, 0) ' �ŐV�͐�
    End With
    
    ' ���ɖ��̕t��
    progress.Name = LINE_NAME
    
    delay_line = Array(7, 15, 30)
    delay_days = Array(2, 1, 3)
    
    For i = 0 To UBound(delay_line)
        ' �ϋȍ��W��ݒ�
        start_y = Sheets(SHEET_WBS_NAME).Cells(delay_line(i), DATE_ROW).Top
        end_y = Sheets(SHEET_WBS_NAME).Cells(delay_line(i) + 1, DATE_ROW).Top
    
        ' ���_���W��ݒ�
        top_x = Sheets(SHEET_WBS_NAME).Cells(delay_line(i), (DATE_ROW - delay_days(i))).Left
        
        ' �s�̒����ɒ��_��ݒ�
        top_y = (end_y - start_y) / 2 + start_y
    
        ' ���_��`��
        With progress
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, start_x, start_y
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, top_x, top_y
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, end_x, end_y
        End With
    Next i
End Sub

