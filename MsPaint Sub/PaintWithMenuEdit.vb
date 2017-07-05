Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Public Class PaintWithMenuEdit
    '�����O�w�q�D�\���s�譶,�γ������e�\����󪺥~�[�ι�@
    Inherits PaintWithMenuPrint

    Private miEditUndo, miEditRepeat As MenuItemHelp
    Private miEditCut, miEditCopy, miEditPaste As MenuItemHelp
    Private miEditDel, miEditSelectAll, miEditCopyTo, miEditPasteTo As MenuItemHelp
    '�D�\�����
    Private ConEditCut, ConEditCopy, ConEditPaste As MenuItemHelp
    Private ConEditDel, ConEDitSelectAll, ConEditCopyTo, ConEditPasteTo As MenuItemHelp
    '���e�\�����

    Const strOpenFilter As String = _
  "�I�}���ɮ�(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
  "GIF �ϧΥ洫�榡(*.GIF)|*.gif|" & _
  "JPEG �ϧΥ洫�榡(*.JPG;,*.JPEG)|*.jpg;*.jpeg|" & _
  "�Ҧ��Ϥ��ɮ�|*.bmp;*.dib;*.ico;*.gif;*.jpg;*.jpeg|" & _
  "�Ҧ��ɮ�|*.*"
    Const strSaveFilter As String = _
    "����I�}��(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "16 ���I�}��(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "256���I�}��(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "24 �줸�I�}��(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "GIF �ϧΥ洫�榡(*.GIF)|*.gif|" & _
    "JPEG �ϧΥ洫�榡(*.JPG;,*.JPEG)|*.jpg;*.jpeg"

    Sub New()
        '�D�\���s�趵�غc��
        Dim mi As New MenuItem("�s��(&E)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuEditOnPopup

        miEditUndo = New MenuItemHelp("�_��(&U)", sbpMenu)
        miEditUndo.Shortcut = Shortcut.CtrlZ
        miEditUndo.HelpText = "�����e�����ʧ@�C"
        Menu.MenuItems(index).MenuItems.Add(miEditUndo)
        AddHandler miEditUndo.Click, AddressOf MenuEditUndoandRepeatOnClick

        miEditRepeat = New MenuItemHelp("����(&R)", sbpMenu)
        miEditRepeat.Shortcut = Shortcut.CtrlY
        miEditRepeat.HelpText = "�����e���������ʧ@�C"
        Menu.MenuItems(index).MenuItems.Add(miEditRepeat)
        AddHandler miEditRepeat.Click, AddressOf MenuEditUndoandRepeatOnClick

        Menu.MenuItems(index).MenuItems.Add("-")

        miEditCut = New MenuItemHelp("�ŤU(&T)", sbpMenu)
        miEditCut.Shortcut = Shortcut.CtrlX
        miEditCut.HelpText = "�ŤU��ܪ����ءA�é��ŶKï���C"
        Menu.MenuItems(index).MenuItems.Add(miEditCut)
        AddHandler miEditCut.Click, AddressOf MenuEditCutOnClick

        miEditCopy = New MenuItemHelp("�ƻs(&C)", sbpMenu)
        miEditCopy.Shortcut = Shortcut.CtrlC
        miEditCopy.HelpText = "�ƻs��ܪ����ءA�é��ŶKï���C"
        Menu.MenuItems(index).MenuItems.Add(miEditCopy)
        AddHandler miEditCopy.Click, AddressOf MenuEditCopyOnClick

        miEditPaste = New MenuItemHelp("�K�W(&P)", sbpMenu)
        miEditPaste.Shortcut = Shortcut.CtrlV
        miEditPaste.HelpText = "���J�ŶKï�����e�C"
        Menu.MenuItems(index).MenuItems.Add(miEditPaste)
        AddHandler miEditPaste.Click, AddressOf MenuEditPasteOnClick

        miEditDel = New MenuItemHelp("�M����ܪ�����(&L)", sbpMenu)
        miEditDel.Shortcut = Shortcut.Del
        miEditDel.HelpText = "�R����ܪ�����"
        Menu.MenuItems(index).MenuItems.Add(miEditDel)
        AddHandler miEditDel.Click, AddressOf MenuEditDelOnClick

        miEditSelectAll = New MenuItemHelp("����(&A)", sbpMenu)
        miEditSelectAll.Shortcut = Shortcut.CtrlA
        miEditSelectAll.HelpText = "��ܥ���"
        Menu.MenuItems(index).MenuItems.Add(miEditSelectAll)
        AddHandler miEditSelectAll.Click, AddressOf MenuEditSelectAllOnClick

        Menu.MenuItems(index).MenuItems.Add("-")

        miEditCopyTo = New MenuItemHelp("�ƻs��(&O)...", sbpMenu)
        miEditCopyTo.HelpText = "�N�ҿ諸�����ƻs���ɮ�"
        Menu.MenuItems(index).MenuItems.Add(miEditCopyTo)
        AddHandler miEditCopyTo.Click, AddressOf MenuEditCopyToOnClick

        miEditPasteTo = New MenuItemHelp("�K�W�ӷ�(&F)...", sbpMenu)
        miEditPasteTo.HelpText = "�N�ɮ׶K��ҿ諸����"
        Menu.MenuItems(index).MenuItems.Add(miEditPasteTo)
        AddHandler miEditPasteTo.Click, AddressOf MenuEditPasteToOnClick

        CreateContextMenu()

    End Sub
    Private Sub CreateContextMenu()
        '���e�\���غc��

        AddHandler ConMenuEdit.Popup, AddressOf ConMenuOnPopup

        ConEditCut = New MenuItemHelp("�ŤU(&T)", sbpMenu)
        ConEditCut.HelpText = "�ŤU��ܪ����ءA�é��ŶKï���C"
        ConMenuEdit.MenuItems.Add(ConEditCut)
        AddHandler ConEditCut.Click, AddressOf MenuEditCutOnClick

        ConEditCopy = New MenuItemHelp("�ƻs(&C)", sbpMenu)
        ConEditCopy.HelpText = "�ƻs��ܪ����ءA�é��ŶKï���C"
        ConMenuEdit.MenuItems.Add(ConEditCopy)
        AddHandler ConEditCopy.Click, AddressOf MenuEditCopyOnClick

        ConEditPaste = New MenuItemHelp("�K�W(&P)", sbpMenu)
        ConEditPaste.HelpText = "���J�ŶKï�����e�C"
        ConMenuEdit.MenuItems.Add(ConEditPaste)
        AddHandler ConEditPaste.Click, AddressOf MenuEditPasteOnClick

        ConEditDel = New MenuItemHelp("�M����ܳ���(&L)", sbpMenu)
        ConEditDel.HelpText = "�R����ܪ�����"
        ConMenuEdit.MenuItems.Add(ConEditDel)
        AddHandler ConEditDel.Click, AddressOf MenuEditDelOnClick

        ConEDitSelectAll = New MenuItemHelp("����(&A)", sbpMenu)
        ConEDitSelectAll.HelpText = "��ܥ���"
        ConMenuEdit.MenuItems.Add(ConEDitSelectAll)
        AddHandler ConEDitSelectAll.Click, AddressOf MenuEditSelectAllOnClick

        ConMenuEdit.MenuItems.Add("-")

        ConEditCopyTo = New MenuItemHelp("�ƻs��(&O)...", sbpMenu)
        ConEditCopyTo.HelpText = "�N�ҿ諸�����ƻs���ɮ�"
        ConMenuEdit.MenuItems.Add(ConEditCopyTo)
        AddHandler ConEditCopyTo.Click, AddressOf MenuEditCopyToOnClick

        ConEditPasteTo = New MenuItemHelp("�K�W�ӷ�(&F)...", sbpMenu)
        ConEditPasteTo.HelpText = "�N�ɮ׶K��ҿ諸����"
        ConMenuEdit.MenuItems.Add(ConEditPasteTo)
        AddHandler ConEditPasteTo.Click, AddressOf MenuEditPasteToOnClick

        ConMenuEdit.MenuItems.Add("-")
    End Sub
    Private Sub MenuEditOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        '�D�\���s�趵,���ا@�Ϊ��A
        If GraphicsData.Count = 0 OrElse GraphicsData.Count - UndoCount = 0 Then
            miEditUndo.Enabled = False
        Else
            miEditUndo.Enabled = True
        End If

        If UndoCount = 0 Then
            miEditRepeat.Enabled = False
        Else
            miEditRepeat.Enabled = True
        End If
        '���o�إi�H�g miEditRepeat.Enabled = CBool(UndoCount)
        '���O�ӤH�٬O�{���o�˪��g�k,�ݰ_�Ӥ���M������

        miEditCut.Enabled = bEdit
        miEditCopy.Enabled = bEdit
        miEditDel.Enabled = bEdit
        miEditCopyTo.Enabled = bEdit

        Dim data As IDataObject = Clipboard.GetDataObject()
        miEditPaste.Enabled = data.GetDataPresent(GetType(Bitmap))
    End Sub
    Private Sub ConMenuOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        '���e�\���,���ا@�Ϊ��A
        ConEditCut.Enabled = bEdit
        ConEditCopy.Enabled = bEdit
        ConEditDel.Enabled = bEdit
        ConEditCopyTo.Enabled = bEdit

        Dim data As IDataObject = Clipboard.GetDataObject()
        ConEditPaste.Enabled = data.GetDataPresent(GetType(Bitmap))
    End Sub

    Private Sub MenuEditUndoandRepeatOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�_��/���Ц欰
        Dim MenuHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        Dim EumShift() As Integer = {1, -1}
        Dim Shift As Integer = EumShift(MenuHelp.Index())

        Dim CurrentCount As Integer
        If Shift > 0 Then
            CurrentCount = Math.Max(GraphicsData.Count - UndoCount - 1, 0)
        Else
            CurrentCount = Math.Min(GraphicsData.Count - UndoCount + 1, GraphicsData.Count - 1)
        End If
        Dim GrapItem As GraphicsItem = DirectCast(GraphicsData(CurrentCount), GraphicsItem)
        If GrapItem.DataType = "B" AndAlso Not GrapItem.IsSizeChange AndAlso Not GrapItem.IsFont Then
            UndoCount += (Shift * 2)
        Else
            UndoCount += Shift
        End If
        
        If GrapItem.IsSizeChange Then
            pnWorkarea.Size = GrapItem.SourceSize
            szBmp = New Size(GrapItem.SourceSize.Width / ZoomBase, GrapItem.SourceSize.Height / ZoomBase)
        End If

        pnWorkarea.Refresh()
        '�o�بϥΤF�@�� UndoCount +=(Shift * 2),��M�O�@�ӫܮt���@�k
        '���L��ۤp�e�a���h�O,��_��@�ӥ� EditPanel �Ҳ��ͪ���ʮ�        
    End Sub

    Private Sub MenuEditCutOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Clipboard.SetDataObject(pnEditarea.BmpData, True)

        MenuEditDelOnClick(obj, ea)
    End Sub
    Private Sub MenuEditCopyOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Clipboard.SetDataObject(pnEditarea.BmpData, True)
    End Sub
    Private Sub MenuEditPasteOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        radbtnToolbar(1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(1), ea)         '���ʤu��C�ﶵ

        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        Dim data As IDataObject = Clipboard.GetDataObject()
        Dim bmp As Bitmap = DirectCast(data.GetData(GetType(Bitmap)), Bitmap)

        CheckDataSize(bmp)
    End Sub
    Private Sub MenuEditDelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        pnEditarea.Dispose()
        bEdit = False
    End Sub
    Private Sub MenuEditSelectAllOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        radbtnToolbar(1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(1), ea)          '���ʤu��C�ﶵ
        Dim rect As New Rectangle(0, 0, pnWorkarea.Width, pnWorkarea.Height)
        Dim path As New GraphicsPath
        path.AddRectangle(rect)

        CreateEditPanel(rect, path, False)
    End Sub
    Private Sub MenuEditCopyToOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim sfd As New SaveFileDialog
        sfd.Filter = strSaveFilter

        sfd.FilterIndex = 4
        If sfd.ShowDialog() = DialogResult.OK Then
            Dim FileName As String = sfd.FileName
            Select Case sfd.FilterIndex
                Case 0
                    pnEditarea.BmpData.Save(FileName)
                Case 1, 2, 3, 4
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Bmp)
                Case 5
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Gif)
                Case 6
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Jpeg)
            End Select
        End If
    End Sub
    Private Sub MenuEditPasteToOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        Dim ofd As New OpenFileDialog
        ofd.Filter = strOpenFilter
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim bmp As New Bitmap(ofd.FileName)
            CheckDataSize(bmp)
        End If
    End Sub
    Private Sub CheckDataSize(ByVal bmp As Bitmap)
        '�� bitmap �j�p�����B�z�覡
        Dim Rect As New Rectangle(0, 0, bmp.Width * ZoomBase, bmp.Height * ZoomBase)

        If Rect.Width >= pnWorkarea.Width OrElse Rect.Height >= pnWorkarea.Height Then
            Dim dr As DialogResult = _
            MessageBox.Show("[�ŶKï]���v���j���I�}�ϡC" & vbLf & "�A�n��j�I�}�϶�?", _
                            strProgName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case dr
                Case DialogResult.Yes
                    szBmp = bmp.Size
                    Dim tmp As New ControlDot
                    Dim mea As MouseEventArgs
                    WorkareaSizeChanged(tmp, mea)
                    tmp.Dispose()

                    CreateEditPanelWithBmp(bmp, False)

                Case DialogResult.No
                    Dim TmpBmp As New Bitmap(szBmp.Width, szBmp.Height)
                    Dim grfx As Graphics = Graphics.FromImage(TmpBmp)
                    grfx.DrawImage(bmp, Rect, Rect, GraphicsUnit.Pixel)
                    grfx.Dispose()

                    CreateEditPanelWithBmp(bmp, False)
                Case DialogResult.Cancel

            End Select
        Else
            CreateEditPanelWithBmp(bmp, False)
        End If
    End Sub
End Class
