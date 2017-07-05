Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class FullScreenShowImage
    '�����O�w�q����\����˵��I�}��,�Ҳ��ͪ��~�[�欰�e��
    Inherits Form
    Private WorkBmp As Bitmap

    Sub New(ByVal bmp As Bitmap)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Location = New Point(0, 0)
        Me.ControlBox = False
        Me.WindowState = FormWindowState.Maximized

        AddHandler Me.KeyDown, AddressOf UserCloseFormOnKeyDown
        AddHandler Me.Click, AddressOf UserCloseFormOnClick
        AddHandler Me.Paint, AddressOf FormOnPaint
        WorkBmp = bmp
    End Sub
    Private Sub UserCloseFormOnKeyDown(ByVal obj As Object, ByVal kea As KeyEventArgs)
        '�ϥΪ̫��U�������
        Me.Close()
    End Sub
    Private Sub UserCloseFormOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�ϥΪ��I���ƹ�
        Me.Close()
    End Sub
    Private Sub FormOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        pea.Graphics.Clear(Color.DarkBlue)

        Dim Rect As Rectangle
        If WorkBmp.Width > Me.Width OrElse WorkBmp.Height > Me.Height Then
            Dim szf As New SizeF(WorkBmp.Width / WorkBmp.HorizontalResolution, _
                     WorkBmp.Height / WorkBmp.VerticalResolution)
            Dim fScale As Single = Math.Min(Me.Width / szf.Width, Me.Height / szf.Height)
            szf.Width *= fScale
            szf.Height *= fScale
            Rect = New Rectangle((Me.Width - szf.Width) / 2, (Me.Height - szf.Height) / 2, _
                                               szf.Width, szf.Height)
        Else
            Rect = New Rectangle((Me.Width - WorkBmp.Width) \ 2, (Me.Height - WorkBmp.Height) \ 2 _
                                 , WorkBmp.Width, WorkBmp.Height)
        End If

        pea.Graphics.DrawImage(WorkBmp, Rect)
        '�o�ӥ��ù��˵����\��,�P�p�e�a���P.�p�e�a���@�k�O,�N�ثe�ϥΪ̩Ҩ��쪺�u�@�Ͻd��
        '�H���ù����覡��ܦb�e���W.�Ӥp�̬O�N�����Y��ܿù��i�H��ܪ��d��.        
    End Sub
End Class
