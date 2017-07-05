Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintWorkareaOnPaint
    '�����O�w�q����ø�Ϥu�@�ϤW,��@��ø�Ϧ欰
    Inherits PaintStatusbar

    Protected Workpath As GraphicsPath                  'ø�Ϥu�@�R�O
    Protected GraphicsData As ArrayList                 'ø�ϰʧ@���X
    Protected UndoCount As Integer = 0                  '�_�즸��

    Sub New()
        GraphicsData = New ArrayList
        Workpath = New GraphicsPath
        AddHandler pnWorkarea.Paint, AddressOf WorkareaOnPaint
    End Sub
    Protected Overridable Sub WorkareaOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '����ϥΪ̩Ҩ���ø�Ϥu�@����ܵe��
        pea.Graphics.DrawImage(bmppic, 0, 0, pnWorkarea.Width, pnWorkarea.Height)

        pea.Graphics.ScaleTransform(ZoomBase, ZoomBase)
        GraphircsDataAnalysis(pea.Graphics)

        pea.Graphics.ScaleTransform(1 / ZoomBase, 1 / ZoomBase)
    End Sub
    Protected Function WorkareaChangeBmp() As Bitmap
        '�Nø�Ϥu�@�ϩ���ܪ��v��,�ഫ�� Bitmap

        Dim BmpBack As Bitmap = bmpPic.Clone
        Dim grfxbmp As Graphics = Graphics.FromImage(BmpBack)

        GraphircsDataAnalysis(grfxbmp)

        grfxbmp.Dispose()
        Return BmpBack
        '�Nø�ϰʧ@�ন Bitmap �b�U��ø�ϩR�O��,�j�q�ϥΦ��@�\��
        '���O�ϥγo�Ӥ�k,�b�ϧΤj��,���Ĳv���Ϊ����~
    End Function
    Private Sub GraphircsDataAnalysis(ByVal grfx As Graphics)
        '���R�v�����ϥΪ�ø�̸�T.�å[�H���.

        Dim GrapData() As GraphicsItem = CType(GraphicsData.ToArray(GetType(GraphicsItem)), GraphicsItem())
        Dim i, ShowCount As Integer

        ShowCount = GrapData.GetUpperBound(0) - UndoCount      '�����_��欰��,�֦���ø�ϰʧ@�`��.
        For i = 0 To ShowCount
            Select Case GrapData(i).DataType
                Case "D"          '���Iø��
                    Dim br As Brush = New SolidBrush(GrapData(i).Color)
                    Dim apt() As PointF = GrapData(i).Data.PathPoints
                    Dim poRect(apt.GetUpperBound(0)) As Rectangle
                    Dim Width As Integer = GrapData(i).pWidth
                    Dim Count As Integer
                    For Count = 0 To apt.GetUpperBound(0)
                        poRect(Count) = New Rectangle(apt(Count).X, _
                        apt(Count).Y, Width, Width)
                    Next Count
                    grfx.FillRectangles(br, poRect)
                Case "I"         '���I�K��
                    Dim apt() As PointF = GrapData(i).Data.PathPoints
                    Dim Count As Integer
                    Dim img As Image = GrapData(i).Image
                    For Count = 0 To apt.GetUpperBound(0)
                        grfx.DrawImage(img, apt(Count))
                    Next Count
                Case "P"         '�̸��|ø��  
                    Dim pn As New Pen(GrapData(i).Color, GrapData(i).pWidth)
                    grfx.DrawPath(pn, GrapData(i).Data)
                Case "F"         '�̸��|��R
                    Dim br As Brush = New SolidBrush(GrapData(i).FillColor)
                    grfx.FillPath(br, GrapData(i).Data)
                Case "R"         '�̸��|��R+ø��
                    Dim br As Brush = New SolidBrush(GrapData(i).FillColor)
                    grfx.FillPath(br, GrapData(i).Data)
                    Dim pn As New Pen(GrapData(i).Color, GrapData(i).pWidth)
                    grfx.DrawPath(pn, GrapData(i).Data)
                Case "B"         '�̼v���K��                     
                    grfx.DrawImage(GrapData(i).Image, _
                    New Rectangle(GrapData(i).ImageLocation.X, _
                    GrapData(i).ImageLocation.Y, _
                    GrapData(i).ImageSize.Width, _
                    GrapData(i).ImageSize.Height))
                    
            End Select
        Next i
        '���R GraphicsItem �����,�ä��H���
        '�o���ӬO���{�����̭��n������.
        '���ܩ��㪺,�p�̦b�o�إΤF�Ӧh�� Type,�����y���V�ä���
    End Sub

End Class
