Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintStatusbar
    '此類定義關於狀態列的外觀及行為
    Inherits PaintColorArea
    Protected sbar As StatusBar         '狀態列
    Protected sbpMenu, sbpLocation, sbpRectangle As StatusBarPanel   '指令說明,畫面位置,繪製範圍
    Sub New()
        '狀態列建構元

        sbar = New StatusBar
        sbar.Parent = Me
        sbar.ShowPanels = True

        sbpMenu = New StatusBarPanel
        sbpMenu.Text = "如需說明，請按一下[說明]功能表中的[說明主題]。"
        sbpMenu.AutoSize = StatusBarPanelAutoSize.Spring

        sbpLocation = New StatusBarPanel
        sbpRectangle = New StatusBarPanel

        sbar.Panels.AddRange(New StatusBarPanel() {sbpMenu, sbpLocation, sbpRectangle})
    End Sub
    Protected Overrides Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '定義改變繪圖工作區大小時,在狀態列所顯示的資訊
        MyBase.ControldotOnMove(obj, mea)

        If mea.Button <> MouseButtons.Left Then Return
        Dim sz As Point = pnWorkarea.Parent.PointToClient(MousePosition)
        sbpRectangle.Text = sz.X.ToString & "x" & sz.Y.ToString
    End Sub
    Protected Overrides Sub ToolbarHelp(ByVal obj As Object, ByVal ea As EventArgs)
        '覆寫 PaintToolbar.vb,定義當滑鼠移至工具列元件上方時,狀態列所要顯示的訊息.
        Dim radbtnHelp As NoFocusRadbtn = DirectCast(obj, NoFocusRadbtn)

        Select Case CInt(radbtnHelp.Name)
            Case 1
                sbpMenu.Text = "選圖片的任何部份，加以移動，複製或編輯。"
            Case 2
                sbpMenu.Text = "選圖片的矩形部份，加以移動，複製或編輯。"
            Case 3
                sbpMenu.Text = "使用所選的橡皮擦形狀，清除部份圖片。"
            Case 4
                sbpMenu.Text = "使用目前的繪圖色形，填入某個區域。"
            Case 5
                sbpMenu.Text = "挑選圖片中的一種顏色來繪圖。"
            Case 6
                sbpMenu.Text = "變更倍率。"
            Case 7
                sbpMenu.Text = "以一個寬度像素大小繪出任意形狀的線條。"
            Case 8
                sbpMenu.Text = "使用所選的刷子形狀和大小來繪圖。"
            Case 9
                sbpMenu.Text = "用所選的噴鎗大小來繪圖。"
            Case 10
                sbpMenu.Text = "將文字貼到圖片中。"
            Case 11
                sbpMenu.Text = "用所選的線條寬度繪製直線。"
            Case 12
                sbpMenu.Text = "用所選的線條寬度繪製曲線。"
            Case 13
                sbpMenu.Text = "用所選的填入樣式繪製矩形。"
            Case 14
                sbpMenu.Text = "用所選的填入樣式繪製多邊形。"
            Case 15
                sbpMenu.Text = "用所選的填入樣式繪製橢圓形。"
            Case 16
                sbpMenu.Text = "用所選的填入樣式繪製圓角矩形。"
        End Select
    End Sub
    Protected Overrides Sub ToolbarHelpLeave(ByVal obj As Object, ByVal ea As EventArgs)
        '離開工具列元件時,狀態列所顯示的訊息
        sbpMenu.Text = "如需說明，請按一下[說明]功能表中的[說明主題]。"
    End Sub
End Class
