Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class EditPanel
    '�����O����s�趵���󪺥~�[�Ψ�欰
    Inherits RectControlDot

    Private BmpEditdata As Bitmap
    Private bTracking As Boolean
    Private bTransParent As Boolean
    Private bFormTransParent As Boolean
    Private clrBack As Color
    Private ptLastOffest, ptbmpOffest As Point

    '�o�س̤j�����D�O,����z���⪺�B�z.
    '�b�ϥΤF ControlStyles.SupportsTransparentBackColor, True ��~�o�{ .Net
    '��o�Ӥ�k���Ĳv���۷��}.
    '���ӷQ�b�D�n���Ҳդ�,������ GraphricsItem �ӹ�� EditPanel ���ʧ@.
    '���L�Ҷq��g�o�ӵ{��,�D�n�b�ǲ� .Net �Ϊ���ɦV�Ȥ������

    Property BmpData() As Bitmap
        Set(ByVal Value As Bitmap)
            BmpEditdata = Value
        End Set
        Get
            Return BmpEditdata
        End Get
    End Property
    WriteOnly Property TransParent() As Boolean
        Set(ByVal Value As Boolean)
            bTransParent = Value
            If bTransParent Then
                BmpEditdata.MakeTransparent(clrBack)
                BackColor = Color.FromArgb(0, clrBack.R, clrBack.G, clrBack.B)
            Else
                BmpEditdata.MakeTransparent(Color.Empty)
            End If
        End Set
    End Property
    ReadOnly Property FormTransParent() As Boolean
        Get
            Return bFormTransParent
        End Get
    End Property
   
    Sub New(ByVal rect As Rectangle, ByVal Bmp As Bitmap, _
            ByVal bTrans As Boolean, ByVal bFormTrans As Boolean, ByVal clr As Color)

        setstyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint _
                Or ControlStyles.DoubleBuffer Or _
                ControlStyles.SupportsTransparentBackColor, True)

        bTransParent = bTrans
        bFormTransParent = bFormTrans
        clrBack = clr
        bDraw = True                  '�W�h�ܼ�,���ܦb�ܧ�j�p�欰��,�O�_�nø�s�u�q���

        BmpEditdata = New Bitmap(Bmp)

        If bFormTransParent Then BackColor = Color.FromArgb(0, clrBack.R, clrBack.G, clrBack.B)
        If bTransParent Then BmpEditdata.MakeTransparent(clrBack) '����s�褸��ιϧγz����ϥ�

        Location = New Point(rect.X, rect.Y)
        Size = New Size(rect.Width, rect.Height)
        Cursor = New Cursor(Me.GetType(), "RegionMove.cur")

        AddHandler Me.Paint, AddressOf EditPanelOnPaint
        AddHandler Me.MouseDown, AddressOf EditPanelMouseOnClick
        AddHandler Me.MouseMove, AddressOf EditPanelMouseOnMove
        AddHandler Me.MouseUp, AddressOf EditPanelMouseOnUp
    End Sub
    Private Sub EditPanelOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '�s�趵�~�[�欰
        pea.Graphics.DrawImage(BmpData, 0, 0, Me.Width, Me.Height)

        If bTracking Then                     '��ƹ��l�ܦ欰�}�Үɰ�����ܥ~��
            MyBase.AllControlDotHide()
            Return
        End If
        Dim path As New GraphicsPath
        Dim x, y As Integer
        Dim Width As Integer = Me.Width - 2
        Dim Height As Integer = Me.Height - 2

        For x = 1 To Width Step 8
            path.AddLine(New Point(x, 1), New Point(x + 4, 1))
            path.StartFigure()
            path.AddLine(New Point(x, Height), New Point(x + 4, Height))
            path.StartFigure()
        Next x
        For y = 1 To Height Step 8
            path.AddLine(New Point(1, y), New Point(1, y + 4))
            path.StartFigure()
            path.AddLine(New Point(Width, y), New Point(Width, y + 4))
            path.StartFigure()
        Next y

        pea.Graphics.DrawRectangle(New Pen(Color.White, 1), 1, 1, Width - 1, Height - 1)
        pea.Graphics.DrawPath(New Pen(Color.DarkBlue), path)
        'ø�s�~��
    End Sub
    Private Sub EditPanelMouseOnClick(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If mea.Button <> MouseButtons.Left Then Return

        bTracking = True
        Dim ptMouse As New Point(Me.Parent.MousePosition.X, Me.Parent.MousePosition.Y)
        ptLastOffest = New Point(ptMouse.X - Me.Location.X, ptMouse.Y - Me.Location.Y)
        Me.Refresh()
    End Sub
    Private Sub EditPanelMouseOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '������ʦ欰�B�z.
        If Not bTracking Then Return

        Dim ptMouse As New Point(Me.Parent.MousePosition.X, Me.Parent.MousePosition.Y)
        Me.Location = New Point(ptMouse.X - ptLastOffest.X, ptMouse.Y - ptLastOffest.Y)
    End Sub
    Private Sub EditPanelMouseOnUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '�ƹ��l�ܰ���
        bTracking = False
        MyBase.AllControlDotShow()
        Me.Refresh()
    End Sub
    Protected Overrides Sub PanelResize(ByVal obj As Object, ByVal ea As EventArgs)
        '����j�p�ܧ�
        MyBase.PanelResize(obj, ea)
        Me.Refresh()
    End Sub
End Class
