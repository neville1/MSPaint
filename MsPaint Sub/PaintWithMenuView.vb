Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class PaintWithMenuView
    '本類別定義主功能表檢視項元件的外觀及實作
    Inherits PaintWithMenuEdit

    Private miViewToolbar, miViewColorArea, miViewStatusbar, miViewTextbox As MenuItemHelp
    Private miZoomStand, miZoomLarge As MenuItemHelp
    Private miZoomShowLine, miZoomReview As MenuItemHelp
    Private pathLine As GraphicsPath
    Private Mircocosm As BitmapMicrocosm                 '縮圖表單
    '此部份管理畫面元件配置,圖形縮放比
    Sub New()
        Dim mi As New MenuItem("檢視(&V)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuViewOnPopup

        miViewToolbar = New MenuItemHelp("工具箱(&T)")
        AddHandler miViewToolbar.Click, AddressOf MenuViewToolbarOnClick
        miViewToolbar.Shortcut = Shortcut.CtrlT
        miViewToolbar.HelpPanel = sbpMenu
        miViewToolbar.HelpText = "顯示或隱藏工具箱。"
        Menu.MenuItems(index).MenuItems.Add(miViewToolbar)

        miViewColorArea = New MenuItemHelp("色塊(&C)")
        AddHandler miViewColorArea.Click, AddressOf MenuViewColorAreaOnClick
        miViewColorArea.Shortcut = Shortcut.CtrlL
        miViewColorArea.HelpPanel = sbpMenu
        miViewColorArea.Checked = True
        miViewColorArea.HelpText = "顯示或隱藏色塊。"
        Menu.MenuItems(index).MenuItems.Add(miViewColorArea)

        miViewStatusbar = New MenuItemHelp("狀態列(&S)")
        AddHandler miViewStatusbar.Click, AddressOf MenuViewStatusbarOnClick
        miViewStatusbar.HelpPanel = sbpMenu
        miViewStatusbar.HelpText = "顯示或隱藏狀態列。"
        Menu.MenuItems(index).MenuItems.Add(miViewStatusbar)

        miViewTextbox = New MenuItemHelp("文字工具列(&E)")
        AddHandler miViewTextbox.Click, AddressOf MenuViewTextboxOnClick
        miViewTextbox.HelpPanel = sbpMenu
        miViewTextbox.HelpText = "顯示或隱藏文字工具列。"
        miViewTextbox.Enabled = False
        Menu.MenuItems(index).MenuItems.Add(miViewTextbox)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim miZoom As New MenuItem("縮放(&Z)")
        Menu.MenuItems(index).MenuItems.Add(miZoom)

        miZoomStand = New MenuItemHelp("標準大小(&N)")
        miZoomStand.HelpPanel = sbpMenu
        miZoomStand.HelpText = "將圖片放大 100%。"
        miZoom.MenuItems.Add(miZoomStand)
        AddHandler miZoomStand.Click, AddressOf ZoomStandOnClick

        miZoomLarge = New MenuItemHelp("放大(&L)")
        miZoomLarge.HelpPanel = sbpMenu
        miZoomLarge.HelpText = "將圖片放大 400%。"
        miZoom.MenuItems.Add(miZoomLarge)
        AddHandler miZoomLarge.Click, AddressOf ZoomLargeOnClick

        Dim miZoomSet As New MenuItemHelp("自訂(&U)...")
        miZoomSet.HelpPanel = sbpMenu
        miZoomSet.HelpText = "縮放圖片。"
        miZoom.MenuItems.Add(miZoomSet)
        AddHandler miZoomSet.Click, AddressOf ZoomSetOnClick

        miZoom.MenuItems.Add("-")

        miZoomShowLine = New MenuItemHelp("顯示格線(&G)")
        miZoomShowLine.Shortcut = Shortcut.CtrlG
        miZoomShowLine.HelpPanel = sbpMenu
        miZoomShowLine.HelpText = "顯示或隱藏格線。"
        miZoomShowLine.Enabled = False
        miZoomShowLine.Checked = False
        miZoom.MenuItems.Add(miZoomShowLine)
        AddHandler miZoomShowLine.Click, AddressOf ZoomShowLineOnClick

        miZoomReview = New MenuItemHelp("顯示縮圖(&H)")
        miZoomReview.HelpPanel = sbpMenu
        miZoomReview.HelpText = "顯示或隱藏縮圖。"
        miZoomReview.Enabled = False
        miZoom.MenuItems.Add(miZoomReview)
        AddHandler miZoomReview.Click, AddressOf ZoomReviewOnClick

        Dim miViewShowBitmap As New MenuItemHelp("檢視點陣圖(&V)")
        miViewShowBitmap.Shortcut = Shortcut.CtrlF
        miViewShowBitmap.HelpPanel = sbpmenu
        miViewShowBitmap.HelpText = "顯示整個圖片。"
        Menu.MenuItems(index).MenuItems.Add(miViewShowBitmap)
        AddHandler miViewShowBitmap.Click, AddressOf ShowFullScreenBitmap
    End Sub
    Protected Overrides Sub OnLoad(ByVal ea As EventArgs)
        '載入基本元件的起始狀態
        MyBase.OnLoad(ea)

        miViewToolbar.Checked = bToolbar
        miViewColorArea.Checked = bColorbar
        miViewStatusbar.Checked = bStatusbar
        miViewTextbox.Checked = bTextFrom
    End Sub
    Private Sub MenuViewOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        miViewTextbox.Enabled = bText

        If ZoomBase = 1 Then
            miZoomStand.Enabled = False
            miZoomLarge.Enabled = True
        Else
            miZoomStand.Enabled = True
        End If
        If ZoomBase >= 2 Then
            miZoomReview.Enabled = True
        Else
            miZoomReview.Enabled = False
        End If
        If ZoomBase >= 4 Then
            miZoomShowLine.Enabled = True
        Else
            miZoomShowLine.Enabled = False
        End If
    End Sub
    Private Sub MenuViewToolbarOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '工具列顯示/隱藏
        bToolbar = Not bToolbar

        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bToolbar
        pntoolbar.Visible = bToolbar

        MyBase.FormSizeChanged(Me.GetType(), ea)
    End Sub
    Private Sub MenuViewColorAreaOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '色塊顯示/隱藏
        bColorbar = Not bColorbar
        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bColorbar
        pnColorArea.Visible = bColorbar

        MyBase.FormSizeChanged(Me.GetType(), ea)
    End Sub
    Private Sub MenuViewStatusbarOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '狀態列顯示/隱藏
        bStatusbar = Not bStatusbar
        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bStatusbar
        sbar.Visible = bStatusbar

        MyBase.FormSizeChanged(Me.GetType(), ea)        
    End Sub
    Private Sub MenuViewTextboxOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '文字工具列顯示/隱藏
        bTextFrom = Not bTextFrom
        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bTextFrom

        If bTextFrom Then
            frmFontSet.Show()
        Else
            frmFontSet.Hide()
        End If
    End Sub

    Private Sub ZoomStandOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '標準大小
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()
        miZoomLarge.Enabled = True

        ZoomBase = 1
        ZoomBaseChange()
    End Sub
    Private Sub ZoomLargeOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '放大400%
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()
        ZoomBase = 4
        ZoomBaseChange()
    End Sub
    Private Sub ZoomSetOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '自訂放大倍率
        Dim dlg As New BitmapZoomSet(ZoomBase)

        If dlg.ShowDialog = DialogResult.OK AndAlso ZoomBase <> dlg.ZoomBase Then
            ZoomBase = dlg.ZoomBase
            ZoomBaseChange()
        End If
    End Sub
    Private Sub ZoomShowLineOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '顯示格線顯示/隱藏
        miZoomShowLine.Checked = Not miZoomShowLine.Checked
        pnWorkarea.Refresh()
    End Sub
    Private Sub ZoomReviewOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '縮圖顯示/隱藏
        miZoomReview.Checked = Not miZoomReview.Checked

        If miZoomReview.Checked Then
            Mircocosm = New BitmapMicrocosm(WorkareaChangeBmp())
            AddHandler Mircocosm.Closing, AddressOf MircocosmOnClosed
            Me.AddOwnedForm(Mircocosm)
            Mircocosm.Show()
        Else
            Me.RemoveOwnedForm(Mircocosm)
            Mircocosm.Dispose()
        End If
        '顯示縮圖,附帶提個與程式無關的事情,小畫家的顯示縮圖功能,在變更畫面大小時,
        '螢幕的閃爍非常嚴重,原因不知為何.
    End Sub
    Protected Overrides Sub WorkareaOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '顯示格式實作部份
        MyBase.WorkareaOnPaint(obj, pea)

        If miZoomShowLine.Checked AndAlso ZoomBase >= 4 Then

            pathLine = New GraphicsPath
            Dim x, y As Integer
            For x = 0 To pnWorkarea.Width Step ZoomBase
                pathLine.AddLine(New Point(x, 0), New Point(x, pnWorkarea.Height))
                pathLine.StartFigure()
            Next x
            For y = 0 To pnWorkarea.Height Step ZoomBase
                pathLine.AddLine(New Point(0, y), New Point(pnWorkarea.Width, y))
                pathLine.StartFigure()
            Next y

            Dim pn As New Pen(Color.Gray)
            pea.Graphics.DrawPath(pn, pathLine)
        End If
    End Sub

    Private Sub MircocosmOnClosed(ByVal obj As Object, ByVal cea As CancelEventArgs)
        miZoomReview.Checked = False

        Dim frm As DoubleBufferForm = DirectCast(obj, DoubleBufferForm)
        frm.Dispose()
    End Sub
    Private Sub ShowFullScreenBitmap(ByVal obj As Object, ByVal ea As EventArgs)
        '檢視點陣圖
        Dim ShowImage As New FullScreenShowImage(WorkareaChangeBmp())

        ShowImage.Show()
        '此作法與小畫家不同,小畫家是以不變更圖形大小為原則,將目前工作區的所見範圍
        '以全螢幕的方式表現.
        '個人覺得這種作法,在實用價值上不好,是採取以不變更圖形長寬比為原則,顯示出
        '最大可能的全圖形.
    End Sub
End Class
