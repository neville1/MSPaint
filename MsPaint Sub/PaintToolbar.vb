Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintToolBar
    '此類別定義關於工具列外觀及行為

    Inherits PaintWithRegistry

    Protected ToolBarCommand As Integer = 7             '使用者工具列命令選擇
    Protected LastCommand As Integer                    '上一個命令
    Protected pnToolbar As Panel                        '工具列
    Protected radbtnToolbar(15) As NoFocusRadbtn        '工具列功能
    Sub New()
        '工具列建構元
        Dim tx, coly, sy As Integer
        If bToolbar Then tx = 60 Else tx = 0
        If bColorbar Then coly = 50 Else coly = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnToolbar = New Panel
        With pnToolbar
            .Parent = Me
            .BorderStyle = BorderStyle.None
            .Location = New Point(0, 0)
            .Size = New Size(60, ClientSize.Height - coly - sy)
        End With

        Dim TotTool As New ToolTip
        Dim strTotTool() As String = {"選擇任意範圍", "選擇", "橡皮擦/彩色橡皮擦", "填入色彩", _
                                      "挑選顏色", "放大鏡", "鉛筆", "粉刷", "噴槍", "文字", _
                                      "直線", "曲線", "矩形", "多邊形", "橢圓形", "圓角矩形"}

        Dim i, cx As Integer
        For i = 0 To 15
            radbtnToolbar(i) = New NoFocusRadbtn
            With radbtnToolbar(i)
                .Name = (i + 1).ToString
                .Parent = pnToolbar
                .Appearance = Appearance.Button
                .Size = New Size(26, 26)
                .Image = New Bitmap(Me.GetType(), "toolbutton" & (i + 1).ToString & ".bmp")
                If (i + 1) Mod 2 <> 0 Then cx = 3 Else cx = 29
                .Location = New Point(cx, (i \ 2) * 26)
            End With
            TotTool.SetToolTip(radbtnToolbar(i), strTotTool(i))
            AddHandler radbtnToolbar(i).MouseEnter, AddressOf ToolbarHelp
            AddHandler radbtnToolbar(i).MouseLeave, AddressOf ToolbarHelpLeave
            AddHandler radbtnToolbar(i).Click, AddressOf ToolbarRadbtnSelect
        Next i
        radbtnToolbar(6).Checked = True
        pnWorkarea.Cursor = New Cursor(Me.GetType(), "Pen.cur")

    End Sub
    Protected Overrides Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '定義因視覺元件變更(工具列,色塊區)工具列座標配置
        MyBase.FormSizeChanged(obj, ea)

        Dim tx, cy, sy As Integer
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnToolbar.Size = New Size(60, ClientSize.Height - cy - sy)
    End Sub
    Protected Overridable Sub ToolbarHelp(ByVal obj As Object, ByVal ea As EventArgs)
        '當滑鼠移至工具列元件上時,顯示該元件在狀態列的資訊.由 PaintStatusbar 覆寫
    End Sub
    Protected Overridable Sub ToolbarHelpLeave(ByVal obj As Object, ByVal ea As EventArgs)
        '當滑鼠移出工具列元件時,顯示該元件在狀態列的資訊.由 PaintStatusbar 覆寫
    End Sub
    Protected Overridable Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '定義當使用者選擇工具列元件時,所產生的行為
        Dim radbtnSelect As NoFocusRadbtn = DirectCast(obj, NoFocusRadbtn)

        LastCommand = ToolBarCommand               '取得上一個功能

        ToolBarCommand = CInt(radbtnSelect.Name)   '取得使用者所選擇的功能
        Dim CursorName As String
        Select Case ToolBarCommand
            Case 1, 2, 10, 11, 12, 13, 14, 15, 16
                CursorName = "Cross.cur"
            Case 3
                CursorName = "Null.cur"
            Case 4
                CursorName = "FillColor.cur"
            Case 5
                CursorName = "Color.cur"
            Case 6
                CursorName = "ZoomSet.cur"
            Case 7
                CursorName = "Pen.cur"
            Case 8
                CursorName = "Brush.cur"
            Case 9
                CursorName = "Spray.cur"
        End Select
        pnWorkarea.Cursor = New Cursor(Me.GetType(), CursorName)
        '這到目前為止,僅定義了,對應於繪圖工作區的遊標顯示,主要的控制流程,由 Overrides覆寫
    End Sub
End Class
