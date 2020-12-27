Attribute VB_Name = "DrawThunderLine"
Public Const SHEET_WBS_NAME = "シート名"    ' WBSシート名
Public Const LINE_NAME = "ThunderLine"      ' イナズマ線の名称
Public Const LINE_LENGTH = 50               ' イナズマ線の長さ
Public Const DATE_LINE = 5                  ' 日付行
Public Const DATE_ROW = 15                  ' 日付列

Sub DrawThunderLine()
    Dim i, j
    Dim line As Shape                         ' イナズマ線
    Dim start_x As Single, start_y As Single  ' 開始座標
    Dim end_x As Single, end_y As Single      ' 終了座標
    Dim top_x As Single, top_y As Single      ' 頂点座標
    Dim delay_line                            ' 遅延（前倒し）行
    Dim delay_days                            ' 遅延（前倒し）日数
    Dim ffb As FreeformBuilder
    
    '稲妻線を削除する
    For Each s In Sheets(SHEET_WBS_NAME).Shapes
        If InStr(s.Name, LINE_NAME) > 0 Then
            s.Delete
        End If
    Next
    
    
    ' 座標を設定
    start_x = Sheets(SHEET_WBS_NAME).Cells(DATE_LINE, DATE_ROW).Left
    start_y = Sheets(SHEET_WBS_NAME).Cells(DATE_LINE, DATE_ROW).Top
    end_x = start_x
    end_y = Sheets(SHEET_WBS_NAME).Cells(LINE_LENGTH, DATE_ROW).Top

    '
    Set ffb = Sheets(SHEET_WBS_NAME).Shapes.BuildFreeform(msoEditingCorner, start_x, start_y)
    ffb.AddNodes msoSegmentLine, msoEditingCorner, end_x, end_y
    Set progress = ffb.ConvertToShape
    
    ' 直線を引く
    With progress.line
        .DashStyle = msoLineSolid       ' スタイル
        .Weight = 3.5                   ' 太さ
        .ForeColor.RGB = RGB(255, 0, 0) ' 最新は赤
    End With
    
    ' 線に名称付け
    progress.Name = LINE_NAME
    
    delay_line = Array(7, 15, 30)
    delay_days = Array(2, 1, 3)
    
    For i = 0 To UBound(delay_line)
        ' 変曲座標を設定
        start_y = Sheets(SHEET_WBS_NAME).Cells(delay_line(i), DATE_ROW).Top
        end_y = Sheets(SHEET_WBS_NAME).Cells(delay_line(i) + 1, DATE_ROW).Top
    
        ' 頂点座標を設定
        top_x = Sheets(SHEET_WBS_NAME).Cells(delay_line(i), (DATE_ROW - delay_days(i))).Left
        
        ' 行の中央に頂点を設定
        top_y = (end_y - start_y) / 2 + start_y
    
        ' 頂点を描画
        With progress
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, start_x, start_y
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, top_x, top_y
            .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, end_x, end_y
        End With
    Next i
End Sub

