Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class TextPanel
    '�����O�w�q��r�������~�[�ι�@
    Inherits RectControlDot

    Private Txtbox As TextBox
    Private lblsz As Label
    Private FontSz As Size
    Private bmpBack As Bitmap
    Private szBase As Size
    Private RowSize As Integer = 0
    '��r�����,���۷�h����������@,�]�t�~�[,����Y�α��ʤΤ@�ǯS���䪺�欰�B�z
    '�ثe�u���ϥΪ̨ϥμзǦ欰,�~�ॿ�`�u�@
    '�䤤���̤j�����D�O TextChanged �O�b��r��J�����,�~Ĳ�o���欰
    '���O�ݭn���O�b��r�|���b Textbox �ʧ@�e,�H�{���X����

    Property SetFont() As Font
        Get
            Return Txtbox.Font
        End Get
        Set(ByVal Value As Font)
            Txtbox.Font = Value
        End Set
    End Property
    ReadOnly Property GetText() As String()
        Get
            Return Txtbox.Lines
        End Get
    End Property
    Sub New(ByRef obj As Panel, ByVal sz As Size, _
            ByVal fnt As Font, ByVal clr As Color, ByVal TransParent As Boolean)

        Parent = obj

        szBase = New Size(sz.Width * 5 + 6, sz.Height + 6)
        Me.Size = szBase

        AddHandler Me.Paint, AddressOf TextPanelOnPaint

        Txtbox = New TextBox
        With Txtbox
            .Parent = Me
            .Location = New Point(3, 3)
            .AcceptsTab = True
            .AcceptsReturn = False
            .BorderStyle = BorderStyle.None
            .Size = New Size(sz.Width * 5, sz.Height)
            .AutoSize = False
            .Multiline = True
            .Font = fnt
            .ForeColor = clr
            .WordWrap = True
        End With
    End Sub
    Private Sub TextPanelOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '��r�u��C�~�[
        Dim pn As New Pen(Color.DarkBlue, 1)
        pn.DashStyle = Drawing2D.DashStyle.Dash
        Dim grfx As Graphics = pea.Graphics()

        grfx.DrawRectangle(pn, 1, 1, Me.Width - 3, Me.Height - 3)
    End Sub
    Protected Overrides Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        Dim RectEdit As Rectangle
        MyBase.ControldotOnMove(obj, mea)

        If MyBase.Tracking AndAlso _
            MyBase.Rect.X + MyBase.Rect.Width <= Me.Parent.Width AndAlso _
            MyBase.Rect.Y + MyBase.Rect.Height <= Me.Parent.Height Then

            If MyBase.Rect.Width < szBase.Width Then
                MyBase.Rect = New Rectangle(MyBase.Rect.X, MyBase.Rect.Y, _
                                            szBase.Width, MyBase.Rect.Height)
            End If
            If MyBase.Rect.Height < szBase.Height Then
                MyBase.Rect = New Rectangle(MyBase.Rect.X, MyBase.Rect.Y, _
                                            MyBase.Rect.Width, szBase.Height)
            End If

            Me.Location = New Point(MyBase.Rect.X, MyBase.Rect.Y)
            Me.Size = New Size(MyBase.Rect.Width, MyBase.Rect.Height)
            Txtbox.Size = New Size(Me.Width - 6, Me.Height - 6)
        Else
            MyBase.Rect = New Rectangle(Me.Location.X, Me.Location.Y, Me.Width, Me.Height)
        End If
        '�[�W��ɤγ̤p�d�򭭨�,�åB�ק�PEditPanel ���P���j�p�ܧ�欰
    End Sub
    Private Sub TxtboxOnKeyPress(ByVal obj As Object, ByVal kpea As KeyPressEventArgs)
        Dim grfx As Graphics = Txtbox.CreateGraphics()


        If kpea.KeyChar = ChrW(13) AndAlso _
        Me.Location.Y + Me.Height + szBase.Height < Me.Parent.Height Then
            Me.Size = New Size(Me.Width, Me.Height + szBase.Height)
            Txtbox.Size = New Size(Txtbox.Width, Txtbox.Height + szBase.Height)
        End If
    End Sub
End Class
