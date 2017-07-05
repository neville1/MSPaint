Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Public Class PaintWithMenuImage
    '�����O�w�q�D�\���v�����γ������e�\���~�[�ι�@
    Inherits PaintWithMenuView

    Private miImageTransParent, miImageClear As MenuItemHelp
    Sub New()
        '�D�\���v�����غc��
        Dim mi As New MenuItem("�v��(&I)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuImageOnPopup

        Dim miImageTurn As New MenuItemHelp("½��/����(&F)...")
        miImageTurn.Shortcut = Shortcut.CtrlR
        miImageTurn.HelpPanel = sbpMenu
        miImageTurn.HelpText = "½��α���Ϥ��ο�ܪ������C"
        Menu.MenuItems(index).MenuItems.Add(miImageTurn)
        AddHandler miImageTurn.Click, AddressOf MenuImageTurnOnClick

        Dim miImageExtend As New MenuItemHelp("����/�ᦱ(&S)...")
        AddHandler miImageExtend.Click, AddressOf MenuImageExtendOnClick
        miImageExtend.Shortcut = Shortcut.CtrlW
        miImageExtend.HelpPanel = sbpMenu
        miImageExtend.HelpText = "�Ԫ��Χᦱ�Ϥ��ο�ܪ������C"
        Menu.MenuItems(index).MenuItems.Add(miImageExtend)

        Dim miImageColorXor As New MenuItemHelp("��m�ﴫ(&I)")
        AddHandler miImageColorXor.Click, AddressOf MenuImageColorNotOnClick
        miImageColorXor.Shortcut = Shortcut.CtrlI
        miImageColorXor.HelpPanel = sbpMenu
        miImageColorXor.HelpText = "�N�Ϥ��ο�ܪ������C���աC"
        Menu.MenuItems(index).MenuItems.Add(miImageColorXor)

        Dim miImageProperty As New MenuItemHelp("�ݩ�(&A)...")
        AddHandler miImageProperty.Click, AddressOf MenuImagePropertyOnClick
        miImageProperty.Shortcut = Shortcut.CtrlE
        miImageProperty.HelpPanel = sbpMenu
        miImageProperty.HelpText = "�ܧ�Ϥ����ݩʡC"
        Menu.MenuItems(index).MenuItems.Add(miImageProperty)

        miImageClear = New MenuItemHelp("�M���v��(&C)")
        AddHandler miImageClear.Click, AddressOf MenuImageClearOnClick
        miImageClear.Shortcut = Shortcut.CtrlShiftN
        miImageClear.HelpPanel = sbpMenu
        miImageClear.HelpText = "�M���Ϥ��Ωҿ諸�����C"
        Menu.MenuItems(index).MenuItems.Add(miImageClear)

        miImageTransParent = New MenuItemHelp("���z���B�z(&D)")
        AddHandler miImageTransParent.Click, AddressOf MenuImageTransParentOnClick
        With miImageTransParent
            .HelpPanel = sbpMenu
            .HelpText = "�N�ثe��ܪ������ܦ��z���Τ��z���C"
        End With
        Menu.MenuItems(index).MenuItems.Add(miImageTransParent)

        CreateContextMenu()             '�غc���e�\���

    End Sub
    Private Sub CreateContextMenu()
        '���e�\���غc��
        Dim ConImageTurn As New MenuItemHelp("½��/����(&F)...", sbpMenu)
        ConImageTurn.HelpText = "½��α���Ϥ��ο�ܪ������C"
        ConMenuEdit.MenuItems.Add(ConImageTurn)
        AddHandler ConImageTurn.Click, AddressOf MenuImageTurnOnClick

        Dim ConImageExtend As New MenuItemHelp("����/�ᦱ(&S)...", sbpMenu)
        ConImageExtend.HelpText = "�Ԫ��Χᦱ�Ϥ��ο�ܪ������C"
        ConMenuEdit.MenuItems.Add(ConImageExtend)
        AddHandler ConImageExtend.Click, AddressOf MenuImageExtendOnClick

        Dim ConImageColorXor As New MenuItemHelp("��m�ﴫ(&I)", sbpMenu)
        ConImageColorXor.HelpText = "�N�Ϥ��ο�ܪ������C���աC"
        ConMenuEdit.MenuItems.Add(ConImageColorXor)
        AddHandler ConImageColorXor.Click, AddressOf MenuImageColorNotOnClick
    End Sub
    Private Sub MenuImageOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        miImageTransParent.Checked = Not bTransParent

        miImageClear.Enabled = Not bEdit
    End Sub
    Private Sub MenuImageTurnOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '½��
        If bText Then TextAreaComplete()
        Dim dlg As New BitmapTurnDialog

        If dlg.ShowDialog() <> DialogResult.OK Then Return

        Dim bmp As Bitmap
        If bEdit Then
            bmp = pnEditarea.BmpData
        Else
            If bText Then TextAreaComplete()
            bmp = WorkareaChangeBmp()
        End If

        Select Case dlg.SelectCommand
            Case 1
                bmp.RotateFlip(RotateFlipType.RotateNoneFlipX)
            Case 2
                bmp.RotateFlip(RotateFlipType.RotateNoneFlipY)
            Case 3, 4
                bmp.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case 5
                bmp.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case 6
                bmp.RotateFlip(RotateFlipType.Rotate270FlipNone)
        End Select

        If bEdit Then
            Dim ptSt As Point = pnEditarea.Location
            Dim bTrans As Boolean = pnEditarea.FormTransParent
            pnEditarea.Dispose()
            CreateEditPanelWithBmp(bmp, bTrans)
            pnEditarea.Location = ptSt
        Else            
            WorkareaAddGraphicsItem(bmp)
        End If
        dlg.Dispose()
        '����½��,�����Φ�m�ഫ�����@�ӫܤj�����D,�z�Q�����p�U,�{�����ӭn����s��ϰ��
        '�u�@�Ϭ��ۦP���B�z�覡.�i���ڦb�]�p�z���W,�٦��\�h�ݭn���D������.
    End Sub
    Private Sub MenuImageExtendOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�����ᦱ
        If bText Then TextAreaComplete()

        Dim dlg As New BitmapExtendDialog

        If dlg.ShowDialog <> DialogResult.OK Then Return

        If dlg.IsExtend Then
            If bEdit Then
                pnEditarea.Size = New Size(pnEditarea.Width * dlg.ExtendX \ 100, _
                                           pnEditarea.Height * dlg.ExtendY \ 100)
            Else
                If bText Then TextAreaComplete()
                Dim nsz As New Size(pnWorkarea.Width * dlg.ExtendX \ 100, _
                                    pnWorkarea.Height * dlg.ExtendY \ 100)
                Dim bmpExtEnd As New Bitmap(nsz.Width, nsz.Height)
                Dim grfxExtEnd As Graphics = Graphics.FromImage(bmpExtEnd)
                Dim rectSrc As New Rectangle(0, 0, szBmp.Width, szBmp.Height)
                Dim rectDst As New Rectangle(0, 0, nsz.Width, nsz.Height)
                grfxExtEnd.DrawImage(WorkareaChangeBmp(), rectDst, rectSrc, GraphicsUnit.Pixel)
                grfxExtEnd.Dispose()

                WorkareaAddGraphicsItem(bmpExtEnd)
            End If
        End If

        If dlg.IsDistort Then
            If bEdit Then
                Dim bmpback = BitmapDistort(pnEditarea.BmpData, New Point(dlg.DistortX, dlg.DistortY))
                Dim ptSt As Point = pnEditarea.Location
                Dim bTrans As Boolean = pnEditarea.FormTransParent
                pnEditarea.Dispose()

                CreateEditPanelWithBmp(bmpback, bTrans)
                pnEditarea.Location = ptSt
            Else
                Dim bmpback = BitmapDistort(WorkareaChangeBmp(), New Point(dlg.DistortX, dlg.DistortY))

                WorkareaAddGraphicsItem(bmpback)
            End If
        End If
    End Sub
    Private Function BitmapDistort(ByVal bmp As Bitmap, ByVal Distort As Point) As Bitmap
        Dim WorkBmp As Bitmap = bmp
        If Distort.X <> 0 Then
            Dim rDistortX As Double = (Distort.X) * (Math.PI / 180)
            Dim ExtX As Double = WorkBmp.Height * Math.Abs(Math.Tan(rDistortX))
            Dim nsz As New Size(WorkBmp.Width + CInt(ExtX), WorkBmp.Height)

            Dim bmpBack As New Bitmap(nsz.Width, nsz.Height)
            Dim grfxback As Graphics = Graphics.FromImage(bmpback)
            grfxback.Clear(clrBack)
            Dim Marx As New Matrix(1, 0, -(Math.Tan(rDistortX)), 1, 0, 0)
            grfxback.Transform = Marx

            Dim startX As Integer
            If Distort.X > 0 Then startX = CInt(ExtX) Else startX = 0

            grfxback.DrawImage(WorkBmp, startX, 0, WorkBmp.Width, WorkBmp.Height)
            grfxback.Dispose()
            WorkBmp = bmpBack
        End If

        If Distort.Y <> 0 Then
            Dim rDistortY As Double = (Distort.Y) * (Math.PI / 180)
            Dim ExtY As Double = WorkBmp.Width * Math.Abs(Math.Tan(rDistortY))
            Dim nsz As New Size(WorkBmp.Width, WorkBmp.Height + CInt(ExtY))

            Dim bmpBack As New Bitmap(nsz.Width, nsz.Height)
            Dim grfxback As Graphics = Graphics.FromImage(bmpback)
            grfxback.Clear(clrBack)
            Dim Marx As New Matrix(1, -(Math.Tan(rDistortY)), 0, 1, 0, 0)
            grfxback.Transform = Marx

            Dim startY As Integer
            If Distort.Y > 0 Then startY = CInt(ExtY) Else startY = 0
            grfxback.DrawImage(WorkBmp, 0, startY, WorkBmp.Width, WorkBmp.Height)
            grfxback.Dispose()
            WorkBmp = bmpBack
        End If
        '�o�q�ϧΧᦱ���{���X,�̤j�����D�b��ڤ��F�ѷ�ϧ�X,Y�b�P�ɧᦱ��,�������B���¦
        '����? ���X�o�q�į�C�����{���X.

        Return WorkBmp
    End Function
    Protected Overrides Sub WorkareaAddGraphicsItem(ByVal bmp As Bitmap)
        '�N�ϧ��ܧ󵲪G���@ø�Ϧ欰�[�J
        ClearCommandRecord()

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "B"
            .Image = bmp
            .ImageLocation = New Point(0, 0)
            .ImageSize = New Size(bmp.Width, bmp.Height)
            .IsSizeChange = True
            .SourceSize = New Size(pnWorkarea.Width \ ZoomBase, pnWorkarea.Height \ ZoomBase)

        End With
        szBmp = New Size(bmp.Width, bmp.Height)
        pnWorkarea.Size = New Size(szbmp.Width * ZoomBase, szbmp.Height * ZoomBase)
        GraphicsData.Add(GrapMethod)
        pnWorkarea.Refresh()
        '���F�إߧ��@�ʪ� GrapHicsData.�δ_�쪺�Ҷq,�]�ӷl���F�j�q���O����.
        '���ӷ|����n����k.
    End Sub
    Private Sub MenuImageColorNotOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '��m�ﴫ
        If bText Then TextAreaComplete()

        Dim bmpback As Bitmap
        If bEdit Then
            bmpback = pnEditarea.BmpData
        Else
            bmpback = WorkareaChangeBmp()
        End If

        Dim clrTmp As Color
        Dim x, y As Integer
        Dim a, r, b, g As Byte
        For y = 0 To bmpback.Height - 1
            For x = 0 To bmpback.Width - 1
                clrTmp = bmpback.GetPixel(x, y)
                r = Not clrTmp.R
                b = Not clrTmp.B
                g = Not clrTmp.G
                bmpback.SetPixel(x, y, clrTmp.FromArgb(clrTmp.A, r, b, g))
            Next x
        Next y

        If bEdit Then
            Dim ptSt As Point = pnEditarea.Location
            Dim bTrans As Boolean = pnEditarea.FormTransParent
            pnEditarea.Dispose()


            CreateEditPanelWithBmp(bmpback, bTrans)
            pnEditarea.Location = ptSt
        Else
            WorkareaAddGraphicsItem(bmpback)
        End If

        '�o�q�{���X�۷�²��,�o�s�b�۷�j�����D
        '�HPixel �����,�@�Ӥ@�Ӫ����,�Ĳv�w�g�t�줣�౵�����a�B.
    End Sub
    Private Sub MenuImagePropertyOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�ϧ��ݩ�
        Dim dlg As New PropertyDialog(strFileName, SzBmp)
        
        If dlg.ShowDialog <> DialogResult.OK Then Return
        If Not dlg.BmpSize.Equals(szBmp) Then
            szBmp = dlg.BmpSize
            ZoomBase = 1
            Dim mea As MouseEventArgs
            WorkAreaSizeChanged(obj, mea)
        End If
    End Sub
    Private Sub MenuImageClearOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�M���ϧ�
        Dim GrapMethod As New GraphicsItem
        Dim Rect As New Rectangle(0, 0, pnWorkarea.Width, pnWorkarea.Height)

        With GrapMethod
            .DataType = "F"
            GrapMethod.FillColor = clrBack
            GrapMethod.Data.AddRectangle(Rect)
        End With
        GraphicsData.Add(GrapMethod)

        pnWorkarea.Refresh()
        '�o�ئ��ө_�Ǫ��{�H,�b�p�e�a��ۤ�,�u���b�S���s��Ϫ����p�U,�o�ӫ��O�~��
        '�u�@.�����D��۬���p���]�p,�N�p���a.
    End Sub
    Private Sub MenuImageTransParentOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�ϥγz����
        miImageTransParent.Checked = Not miImageTransParent.Checked
        bTransParent = Not bTransParent
        pnToolbar.Refresh()

        If bEdit Then pnEditarea.TransParent = bTransParent
    End Sub
    Protected Overrides Sub WorkareaSizeChanged(ByVal obj As Object, ByVal mea As MouseEventArgs)
        'ø�Ϥu�@�Ϥj�p�ܧ�
        If bText Then TextAreaComplete()

        Dim BmpTmp As New Bitmap(szBmp.Width, szBmp.Height)
        Dim grfxTmpBmp As Graphics = Graphics.FromImage(BmpTmp)
        grfxTmpBmp.Clear(clrBack)

        Dim Bmp As Bitmap = WorkareaChangeBmp()
        grfxTmpBmp.DrawImage(Bmp, 0, 0, Bmp.Width, Bmp.Height)
        grfxTmpBmp.Dispose()

        WorkareaAddGraphicsItem(BmpTmp)
        sbpRectangle.Text = ""
        pnWorkarea.Parent.Refresh()
        '�o�ػP�������P,ø�Ϥu�@�Ϥj�p�ܧ󪺦欰,�Q�����@�ص����Υ[�j.
    End Sub
End Class
