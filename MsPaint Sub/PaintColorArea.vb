Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintColorArea
    '�����O�w�q���������~�[�Φ欰
    Inherits PaintToolbarSub

    Protected pnColorArea As Panel                    '��m�C
    Protected clrFore, clrBack As Color               '�e����.�I����
    Protected btnColorFore, btnColorBack As Button    '����ϥΪ̩Ҩ����e��,�I�����

    Private pbColor(27) As PictureBox

    Sub New()
        '����Ϥ���غc��
        clrFore = Color.Black
        clrBack = Color.White                        '�w�]�C��,�©��զr.

        Dim sy As Integer
        If bStatusbar Then sy = 22 Else sy = 0

        pnColorArea = New Panel
        With pnColorArea
            .Parent = Me
            .BorderStyle = BorderStyle.None
            .Size = New Size(ClientSize.Width, 50)
            .Location = New Point(0, ClientSize.Height - 50 - sy)
        End With

        Dim pbColorarea As New PictureBox
        With pbColorarea
            .Parent = pnColorArea
            .BorderStyle = BorderStyle.Fixed3D
            .BackColor = Color.Gainsboro
            .Size = New Size(30, 32)
            .Location = New Point(0, 10)
        End With

        btnColorFore = New Button
        With btnColorFore
            .Parent = pbColorarea
            .BackColor = clrFore
            .Size = New Size(14, 14)
            .Location = New Point(3, 3)
            .Enabled = False
        End With

        btnColorBack = New Button                     '�o�G�Ӥ���,��ܨϥΪ̥ثe������C��
        With btnColorBack
            .Parent = pbColorarea
            .BackColor = clrBack
            .Size = New Size(14, 14)
            .Location = New Point(10, 11)
            .Enabled = False
        End With

        Dim i As Integer = 27

        Dim strColor() As String = {"Black", "Gray", "DarkRed", _
        "Olive", "Green", "Teal", "DarkBlue", "Purple", "DarkOliveGreen", _
        "DarkSlateGray", "DodgerBlue", "Navy", "MidnightBlue", "SaddleBrown", _
        "White", "Silver", "Red", "Yellow", "Lime", "Cyan", "Blue", "Fuchsia", _
        "Gold", "SpringGreen", "PaleTurquoise", "DarkTurquoise", "Crimson", "DarkOrange"}
        '���B�w�q,�������Ϫ�26���C��.

        For i = 0 To 27
            pbColor(i) = New PictureBox
            pbColor(i).Parent = pnColorArea
            pbColor(i).BackColor = Color.FromName(strColor(i))
            pbColor(i).Size = New Size(16, 16)
            pbColor(i).BorderStyle = BorderStyle.Fixed3D
            If i <= 13 Then
                pbColor(i).Location = New Point(32 + i * 16, 10)
            Else
                pbColor(i).Location = New Point(32 + (i - 14) * 16, 26)
            End If
            AddHandler pbColor(i).MouseDown, AddressOf SelectColorOnClick
        Next i
    End Sub
    Protected Overrides Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '�w�q�]��ı�����ܧ��,����Ϫ��y�ФΦ첾
        MyBase.FormSizeChanged(obj, ea)
        Dim tx, cy, sy As Integer
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnColorArea.Location = New Point(0, ClientSize.Height - 50 - sy)
    End Sub
    Private Sub SelectColorOnClick(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '����ϥD�n�欰,�ѨϥΪ��I���C��ϥ�.
        Dim pb As PictureBox = DirectCast(obj, PictureBox)

        Select Case mea.Button
            Case MouseButtons.Left
                clrFore = pb.BackColor
                btnColorFore.BackColor = clrFore
            Case MouseButtons.Right
                clrBack = pb.BackColor
                btnColorBack.BackColor = clrBack
        End Select
    End Sub
    Protected Overridable Sub ColorEditOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '�ѥ\�����->��m->�s���m,�Ҽv�T��������C���ܧ�

        pbColor(0).BackColor = clrFore
        btnColorFore.BackColor = clrFore
        '�� PaintWithMenuColor.vb ��ʦ�L�Ĥ@���C��
    End Sub

End Class
