Imports Microsoft.Win32
Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

Public Class PaintWithMenuFile
    Inherits PaintCommandWithOther
    '�����O�w�q����D�\��������ɮ׷s�W,�}��,�s�ɤΪ��ݥ\��~�[�غc�Τ�k�B�z

    Protected strFileName As String = "���R�W"

    Private fileColorByte As Integer
    Private strTempFileName(3) As String
    Private miFileTmpName(3) As MenuItemHelp
    Private TempFileIndex As Integer = 0
    Private miFileTmpline As MenuItem    
    Const strFilePath = "TempFilePath"
    Const strTempFileIndex = "TempFileIndex"
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
        '�D�\����ɮ׶��غc��
        Text = strProgName
        Menu = New MainMenu

        Dim mi As New MenuItem("�ɮ�(&F)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miFileNew As New MenuItemHelp("�}�s�ɮ�(&N)")
        AddHandler miFileNew.Click, AddressOf MenuFileNewOnClick
        miFileNew.Shortcut = Shortcut.CtrlN
        miFileNew.HelpPanel = sbpMenu
        miFileNew.HelpText = "�إ߷s�����C"
        Menu.MenuItems(index).MenuItems.Add(miFileNew)

        Dim miFileOpen As New MenuItemHelp("�}������(&O)...")
        AddHandler miFileOpen.Click, AddressOf MenuFileOpenOnClick
        miFileOpen.Shortcut = Shortcut.CtrlO
        miFileOpen.HelpPanel = sbpMenu
        miFileOpen.HelpText = "�}�Ҳ{�������C"
        Menu.MenuItems(index).MenuItems.Add(miFileOpen)

        Dim miFileSave As New MenuItemHelp("�s��(&S)")
        AddHandler miFileSave.Click, AddressOf MenuFileSaveOnClick
        miFileSave.Shortcut = Shortcut.CtrlS
        miFileSave.HelpPanel = sbpMenu
        miFileSave.HelpText = "�x�s�ϥΤ������C"
        Menu.MenuItems(index).MenuItems.Add(miFileSave)

        Dim miFileSaveAs As New MenuItemHelp("�t�s�s��(&A)...")
        AddHandler miFileSaveAs.Click, AddressOf MenuFileSaveAsOnClick
        miFileSaveAs.HelpPanel = sbpMenu
        miFileSaveAs.HelpText = "�H�s�W���x�s�ϥΤ������C"
        Menu.MenuItems(index).MenuItems.Add(miFileSaveAs)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim miFileReviewPrint As New MenuItemHelp("�w���C�L(&V)")
        AddHandler miFileReviewPrint.Click, AddressOf MenuFileReviewPrintOnClick
        miFileReviewPrint.HelpPanel = sbpmenu
        miFileReviewPrint.HelpText = "��ܥ����C"
        Menu.MenuItems(index).MenuItems.Add(miFileReviewPrint)

        Dim miFileSetPrint As New MenuItemHelp("�]�w�C�L�榡(&U)...")
        AddHandler miFileSetPrint.Click, AddressOf MenuFileSetPrintOnClick
        miFileSetPrint.HelpPanel = sbpmenu
        miFileSetPrint.HelpText = "�ܧ󪩭��t�m�C"
        Menu.MenuItems(index).MenuItems.Add(miFileSetPrint)

        Dim miFilePrint As New MenuItemHelp("�C�L(&P)...")
        AddHandler miFilePrint.Click, AddressOf MenuFilePrintOnClick
        miFilePrint.Shortcut = Shortcut.CtrlP
        miFilePrint.HelpPanel = sbpmenu
        miFilePrint.HelpText = "�C�L�ϥΤ������ó]�w�C�L�ﶵ�C"
        Menu.MenuItems(index).MenuItems.Add(miFilePrint)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim i As Integer = 3
        For i = 0 To 3
            miFileTmpName(i) = New MenuItemHelp
            miFileTmpName(i).HelpPanel = sbpmenu
            miFileTmpName(i).HelpText = "�}�ҳo����"
            miFileTmpName(i).Visible = False
            Menu.MenuItems(index).MenuItems.Add(miFileTmpName(i))
            AddHandler miFileTmpName(i).Click, AddressOf MenuFileOpenTempOnClick
        Next i

        miFileTmpline = New MenuItem("-")
        miFileTmpline.Visible = False
        Menu.MenuItems(index).MenuItems.Add(miFileTmpline)

        Dim miFileExit As New MenuItemHelp("����(&X)")
        AddHandler miFileExit.Click, AddressOf MenuFileExitOnClick
        miFileExit.Shortcut = Shortcut.AltF4
        miFileExit.HelpPanel = sbpmenu
        miFileExit.HelpText = "�����p�e�a���s���C"
        Menu.MenuItems(index).MenuItems.Add(miFileExit)

        MakeCaption()
        AddHandler Me.MenuComplete, AddressOf LoseFocusMenu
    End Sub
    Private Sub MenuFileNewOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�}�s�ɮצ欰��@
        If Not OktoTrash() Then Return

        ClearGraphicsData()

        bmpPic = New Bitmap(szBmp.Width, szBmp.Height)
        Dim grfxPic As Graphics = Graphics.FromImage(bmpPic)
        grfxPic.Clear(ClrBack)
        grfxPic.Dispose()

        pnWorkarea.Size = bmpPic.Size       
        strFileName = "���R�W"
        MakeCaption()
        pnWorkarea.Refresh()
        '��m�w�q
    End Sub
    Private Sub MenuFileOpenOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If Not OktoTrash() Then Return

        Dim ofd As New OpenFileDialog
        ofd.Filter = strOpenFilter
        If ofd.ShowDialog() = DialogResult.OK Then LoadFile(ofd.FileName)
    End Sub
    Private Sub MenuFileSaveOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        If strFileName = "���R�W" OrElse strFileName.Length = 0 Then
            SaveFileDlg()
        Else
            SaveFile(0)
        End If
    End Sub
    Private Sub MenuFileSaveAsOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        SaveFileDlg()
    End Sub
    Protected Overridable Sub MenuFileReviewPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�� PaintWithMenuPrint.vb �мg
    End Sub
    Protected Overridable Sub MenuFileSetPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
    End Sub
    Protected Overridable Sub MenuFilePrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
    End Sub
    Private Sub MenuFileExitOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Me.Close()
    End Sub
    Private Sub MenuFileOpenTempOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�}�ҳ̪�T���ɮ�
        If Not OktoTrash() Then Return

        Dim MenuFileIndex As MenuItem = DirectCast(obj, MenuItem)
        Dim FileIndex As Integer
        If MenuFileIndex.Text = Nothing Then Return
        FileIndex = CInt(MenuFileIndex.Text.Substring(1, 1)) - 1
        LoadFile(strTempFileName(FileIndex))
    End Sub
    Protected Overrides Sub Onload(ByVal ea As EventArgs)
        '�H�R�O�C�覡Ū�J�ɮ�
        MyBase.OnLoad(ea)

        Dim astrArgs As String() = Environment.GetCommandLineArgs()
        If astrArgs.Length > 1 Then
            If File.Exists(astrArgs(1)) Then
                LoadFile(astrArgs(1))
            Else
                Dim dr As DialogResult = _
                MessageBox.Show("�䤣���ɮ�: " & Path.GetFileName(astrArgs(1)) & vbLf _
                & vbLf & "�n�إ߷s�ɮ�?", Text, MessageBoxButtons.YesNoCancel)
                Select Case dr
                    Case DialogResult.Yes
                        strFileName = astrArgs(1)
                        File.Create(strFileName).Close()
                        MakeCaption()
                    Case DialogResult.No

                    Case DialogResult.Cancel
                        Close()
                End Select
            End If
        Else

        End If
    End Sub
    Protected Sub LoadFile(ByVal strFileName As String)
        '���ɮ�Ū���I�}��
        Dim img As Image
        Try
            img = Image.FromFile(strFileName)
        Catch ex As OutOfMemoryException
            MessageBox.Show(strFileName & vbLf & "[�p�e�a���s��]�L�kŪ���o���ɮסC" & vbLf & _
            "�o���O���T���I�}���ɮסA�Τ��w�g���A�䴩�o�خ榡�C", Text, MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
            Return
        End Try

        ClearGraphicsData()

        bmpPic = New Bitmap(img)
        SzBmp = bmpPic.Size
        pnWorkarea.Size = SzBmp   
        img.Dispose()
        Me.strFileName = strFileName
        ShowTempFilesMenu()
        MakeCaption()

        pnWorkarea.Refresh()
    End Sub
    Private Sub SaveFile(ByVal SaveFormat As Integer)
        '�g�J�I�}���ɮ�
        bmpPic = New Bitmap(WorkareaChangeBmp())

        Select Case SaveFormat
            Case 0
                bmpPic.Save(strFileName)
            Case 1, 2, 3, 4
                bmpPic.Save(strFileName, ImageFormat.Bmp)
            Case 5
                bmpPic.Save(strFileName, ImageFormat.Gif)
            Case 6
                bmpPic.Save(strFileName, ImageFormat.Jpeg)
        End Select

        GraphicsData.Clear()
        UndoCount = 0
        pnWorkarea.Refresh()
    End Sub
    Private Function SaveFileDlg() As Boolean
        'SaveFile ��ܤ���B�z
        Dim sfd As New SaveFileDialog
        If strFileName.Length > 1 Then
            sfd.FileName = Path.GetFileNameWithoutExtension(strFileName)
        Else
            sfd.FileName = ""
        End If

        sfd.Filter = strSaveFilter
        Select Case Path.GetExtension(strFileName).ToLower()
            Case ".bmp", ".dib"
                Select Case fileColorByte
                    Case 1
                        sfd.FilterIndex = 1
                    Case 4
                        sfd.FilterIndex = 2
                    Case 8
                        sfd.FilterIndex = 3
                    Case 24
                        sfd.FilterIndex = 4
                End Select
            Case ".gif"
                sfd.FilterIndex = 5
            Case ".jpeg", ".jpg"
                sfd.FilterIndex = 6
            Case Else
                sfd.FilterIndex = 4
        End Select

        If sfd.ShowDialog() = DialogResult.OK Then
            strFileName = sfd.FileName
            SaveFile(sfd.FilterIndex)
            MakeCaption()
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub MakeCaption()
        '�]�w Form ���D
        Text = FileTitle() & " - " & strProgName
    End Sub
    Protected Function FileTitle() As String
        Return Path.GetFileName(strFileName)
    End Function
    Private Sub ShowTempFilesMenu()
        Dim i As Integer
        If TempFileIndex > 0 Then
            For i = Math.Min(TempFileIndex, 3) To 1 Step -1
                strTempFileName(i) = strTempFileName(i - 1)
            Next i
        End If
        TempFileIndex = Math.Min(TempFileIndex + 1, 4)
        strTempFileName(0) = strFileName
        miFileTmpline.Visible = True
        For i = 0 To TempFileIndex - 1
            miFileTmpName(i).Text = "&" & (i + 1).ToString & " " & Path.GetFileName(strTempFileName(i))
            miFileTmpName(i).Visible = True
        Next i
        '�o�ؼ������,�̪�T�Ӷ}�Ҫ��ɮ�,�o���ӻP�зǧ@�k���P
    End Sub
    Private Function OktoTrash() As Boolean
        '�P�_�ϧάO�_�ק�
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        If GraphicsData.Count = 0 Then Return True

        Dim dr As DialogResult = MessageBox.Show("�x�s�ק�� " & FileTitle() & "?" _
        , strProgName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)

        Select Case dr
            Case DialogResult.Yes : Return SaveFileDlg()
            Case DialogResult.No : Return True
            Case DialogResult.Cancel : Return False
        End Select
        Return False
    End Function
    Private Sub ClearGraphicsData()
        '�M��ø�ϰʧ@,�]�w���_�l���A
        GraphicsData.Clear()
        UndoCount = 0
        ZoomBase = 1

        clrFore = Color.Black
        clrBack = Color.White
        btnColorFore.BackColor = clrFore
        btnColorBack.BackColor = clrBack
    End Sub
    Protected Overrides Sub SaveRegistryInfo(ByVal regkey As RegistryKey)
        '�N�̪�}�Ҫ��ɮצ�m,�g�J Registry
        MyBase.SaveRegistryInfo(regkey)

        Dim i As Integer
        regkey.SetValue(strTempFileIndex, TempFileIndex)

        For i = 1 To TempFileIndex
            regkey.SetValue(strFilePath & i.ToString, strTempFileName(i - 1))
        Next i
    End Sub
    Protected Overrides Sub LoadRegistryInfo(ByVal regkey As RegistryKey)
        '�� Registry Ū���̪�}�Ҫ��ɮ׸��
        MyBase.LoadRegistryInfo(regkey)

        TempFileIndex = DirectCast(regkey.GetValue(strTempFileIndex, 0), Integer)

        Dim i As Integer
        For i = 1 To TempFileIndex
            strTempFileName(i - 1) = DirectCast(regkey.GetValue(strFilePath & i.ToString, _
            Nothing), String)
        Next i

        If TempFileIndex > 0 Then
            For i = 0 To TempFileIndex - 1
                miFileTmpName(i).Text = "&" & (i + 1).ToString & " " & Path.GetFileName(strTempFileName(i))
                miFileTmpName(i).Visible = True
            Next i
            miFileTmpline.Visible = True
        End If
    End Sub
    Private Sub LoseFocusMenu(ByVal obj As Object, ByVal ea As EventArgs)
        '��ƹ����}�D�\����
        sbpMenu.Text = "�p�ݻ����A�Ы��@�U[����]�\�����[�����D�D]�C"
    End Sub
End Class
