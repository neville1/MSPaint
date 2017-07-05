Imports System
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms
Public Class PaintWithMenuPrint
    '�����O�B�z�\���,��@����C�L���欰.
    Inherits PaintWithMenuFile

    Private prndoc As New PrintDocument
    Private setdlg As New PageSetupDialog
    Private prndlg As New PrintDialog
    Private predlg As New PrintPreviewDialog
    Sub New()
        '�C�L��������غc     
        prndoc.DocumentName = "�p�e�a���s�����"
        AddHandler prndoc.PrintPage, AddressOf OnPrintPage
        setdlg.Document = prndoc
        predlg.Document = prndoc
        prndlg.Document = prndoc

        prndlg.AllowPrintToFile = False
    End Sub
    '�o�ت��C�L�欰,����P�p�e�a���P.�@�譱�o�O�]���p�e�a�w�q���C�L�欰
    '�H�������D,�̧Ƕ�J�U��.���o�ؤ覡���h��.
    '�t�@�譱�h�O�]���ۤv�⥦�w�쬰�@�Ӿǲߤu��,���Q�b�C�L���譱�h�[��z
    '�o�i�H���O�b .Net��,�̳�ª��C�L�覡�F.
    Protected Overrides Sub MenuFileReviewPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�w���C�L
        predlg.ShowDialog()
    End Sub
    Protected Overrides Sub MenuFileSetPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�]�w�C�L�榡
        setdlg.ShowDialog()
    End Sub
    Protected Overrides Sub MenuFilePrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�C�L
        If prndlg.ShowDialog() = DialogResult.OK Then
            prndoc.Print()
        End If
    End Sub
    Private Sub OnPrintPage(ByVal obj As Object, ByVal ppea As PrintPageEventArgs)
        Dim bmpBack As Bitmap = WorkareaChangeBmp()
        Dim RectF = New RectangleF(ppea.MarginBounds.X, ppea.MarginBounds.Y, _
                               ppea.MarginBounds.Width, ppea.MarginBounds.Height)

        Dim szf As New SizeF(bmpBack.Width / bmpBack.HorizontalResolution, _
                             bmpBack.Height / bmpBack.VerticalResolution)
        Dim fScale As Single = Math.Min(RectF.Width / szf.Width, RectF.Height / szf.Height)
        szf.Width *= fScale
        szf.Height *= fScale

        ppea.Graphics.DrawImage(bmpBack, RectF.X + (RectF.Width - szf.Width) / 2, _
                                RectF.Y + (RectF.Height - szf.Height) / 2, szf.Width, szf.Height)
        '�C�L�欰,�ϧΪ����e�񬰥D,�b������e�񪺭�h�U,��j��̤j���i��C�L�d��,�øm����X
    End Sub
End Class
