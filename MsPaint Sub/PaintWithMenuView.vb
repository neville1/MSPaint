Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class PaintWithMenuView
    '�����O�w�q�D�\����˵������󪺥~�[�ι�@
    Inherits PaintWithMenuEdit

    Private miViewToolbar, miViewColorArea, miViewStatusbar, miViewTextbox As MenuItemHelp
    Private miZoomStand, miZoomLarge As MenuItemHelp
    Private miZoomShowLine, miZoomReview As MenuItemHelp
    Private pathLine As GraphicsPath
    Private Mircocosm As BitmapMicrocosm                 '�Y�Ϫ��
    '�������޲z�e������t�m,�ϧ��Y���
    Sub New()
        Dim mi As New MenuItem("�˵�(&V)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuViewOnPopup

        miViewToolbar = New MenuItemHelp("�u��c(&T)")
        AddHandler miViewToolbar.Click, AddressOf MenuViewToolbarOnClick
        miViewToolbar.Shortcut = Shortcut.CtrlT
        miViewToolbar.HelpPanel = sbpMenu
        miViewToolbar.HelpText = "��ܩ����äu��c�C"
        Menu.MenuItems(index).MenuItems.Add(miViewToolbar)

        miViewColorArea = New MenuItemHelp("���(&C)")
        AddHandler miViewColorArea.Click, AddressOf MenuViewColorAreaOnClick
        miViewColorArea.Shortcut = Shortcut.CtrlL
        miViewColorArea.HelpPanel = sbpMenu
        miViewColorArea.Checked = True
        miViewColorArea.HelpText = "��ܩ����æ���C"
        Menu.MenuItems(index).MenuItems.Add(miViewColorArea)

        miViewStatusbar = New MenuItemHelp("���A�C(&S)")
        AddHandler miViewStatusbar.Click, AddressOf MenuViewStatusbarOnClick
        miViewStatusbar.HelpPanel = sbpMenu
        miViewStatusbar.HelpText = "��ܩ����ê��A�C�C"
        Menu.MenuItems(index).MenuItems.Add(miViewStatusbar)

        miViewTextbox = New MenuItemHelp("��r�u��C(&E)")
        AddHandler miViewTextbox.Click, AddressOf MenuViewTextboxOnClick
        miViewTextbox.HelpPanel = sbpMenu
        miViewTextbox.HelpText = "��ܩ����ä�r�u��C�C"
        miViewTextbox.Enabled = False
        Menu.MenuItems(index).MenuItems.Add(miViewTextbox)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim miZoom As New MenuItem("�Y��(&Z)")
        Menu.MenuItems(index).MenuItems.Add(miZoom)

        miZoomStand = New MenuItemHelp("�зǤj�p(&N)")
        miZoomStand.HelpPanel = sbpMenu
        miZoomStand.HelpText = "�N�Ϥ���j 100%�C"
        miZoom.MenuItems.Add(miZoomStand)
        AddHandler miZoomStand.Click, AddressOf ZoomStandOnClick

        miZoomLarge = New MenuItemHelp("��j(&L)")
        miZoomLarge.HelpPanel = sbpMenu
        miZoomLarge.HelpText = "�N�Ϥ���j 400%�C"
        miZoom.MenuItems.Add(miZoomLarge)
        AddHandler miZoomLarge.Click, AddressOf ZoomLargeOnClick

        Dim miZoomSet As New MenuItemHelp("�ۭq(&U)...")
        miZoomSet.HelpPanel = sbpMenu
        miZoomSet.HelpText = "�Y��Ϥ��C"
        miZoom.MenuItems.Add(miZoomSet)
        AddHandler miZoomSet.Click, AddressOf ZoomSetOnClick

        miZoom.MenuItems.Add("-")

        miZoomShowLine = New MenuItemHelp("��ܮ�u(&G)")
        miZoomShowLine.Shortcut = Shortcut.CtrlG
        miZoomShowLine.HelpPanel = sbpMenu
        miZoomShowLine.HelpText = "��ܩ����î�u�C"
        miZoomShowLine.Enabled = False
        miZoomShowLine.Checked = False
        miZoom.MenuItems.Add(miZoomShowLine)
        AddHandler miZoomShowLine.Click, AddressOf ZoomShowLineOnClick

        miZoomReview = New MenuItemHelp("����Y��(&H)")
        miZoomReview.HelpPanel = sbpMenu
        miZoomReview.HelpText = "��ܩ������Y�ϡC"
        miZoomReview.Enabled = False
        miZoom.MenuItems.Add(miZoomReview)
        AddHandler miZoomReview.Click, AddressOf ZoomReviewOnClick

        Dim miViewShowBitmap As New MenuItemHelp("�˵��I�}��(&V)")
        miViewShowBitmap.Shortcut = Shortcut.CtrlF
        miViewShowBitmap.HelpPanel = sbpmenu
        miViewShowBitmap.HelpText = "��ܾ�ӹϤ��C"
        Menu.MenuItems(index).MenuItems.Add(miViewShowBitmap)
        AddHandler miViewShowBitmap.Click, AddressOf ShowFullScreenBitmap
    End Sub
    Protected Overrides Sub OnLoad(ByVal ea As EventArgs)
        '���J�򥻤��󪺰_�l���A
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
        '�u��C���/����
        bToolbar = Not bToolbar

        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bToolbar
        pntoolbar.Visible = bToolbar

        MyBase.FormSizeChanged(Me.GetType(), ea)
    End Sub
    Private Sub MenuViewColorAreaOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '������/����
        bColorbar = Not bColorbar
        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bColorbar
        pnColorArea.Visible = bColorbar

        MyBase.FormSizeChanged(Me.GetType(), ea)
    End Sub
    Private Sub MenuViewStatusbarOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '���A�C���/����
        bStatusbar = Not bStatusbar
        Dim ModiMenuItemHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        ModiMenuItemHelp.Checked = bStatusbar
        sbar.Visible = bStatusbar

        MyBase.FormSizeChanged(Me.GetType(), ea)        
    End Sub
    Private Sub MenuViewTextboxOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '��r�u��C���/����
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
        '�зǤj�p
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()
        miZoomLarge.Enabled = True

        ZoomBase = 1
        ZoomBaseChange()
    End Sub
    Private Sub ZoomLargeOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '��j400%
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()
        ZoomBase = 4
        ZoomBaseChange()
    End Sub
    Private Sub ZoomSetOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�ۭq��j���v
        Dim dlg As New BitmapZoomSet(ZoomBase)

        If dlg.ShowDialog = DialogResult.OK AndAlso ZoomBase <> dlg.ZoomBase Then
            ZoomBase = dlg.ZoomBase
            ZoomBaseChange()
        End If
    End Sub
    Private Sub ZoomShowLineOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '��ܮ�u���/����
        miZoomShowLine.Checked = Not miZoomShowLine.Checked
        pnWorkarea.Refresh()
    End Sub
    Private Sub ZoomReviewOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�Y�����/����
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
        '����Y��,���a���ӻP�{���L�����Ʊ�,�p�e�a������Y�ϥ\��,�b�ܧ�e���j�p��,
        '�ù����{�{�D�`�Y��,��]��������.
    End Sub
    Protected Overrides Sub WorkareaOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '��ܮ榡��@����
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
        '�˵��I�}��
        Dim ShowImage As New FullScreenShowImage(WorkareaChangeBmp())

        ShowImage.Show()
        '���@�k�P�p�e�a���P,�p�e�a�O�H���ܧ�ϧΤj�p����h,�N�ثe�u�@�Ϫ��Ҩ��d��
        '�H���ù����覡��{.
        '�ӤHı�o�o�ا@�k,�b��λ��ȤW���n,�O�Ĩ��H���ܧ�ϧΪ��e�񬰭�h,��ܥX
        '�̤j�i�઺���ϧ�.
    End Sub
End Class
