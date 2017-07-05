Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintToolbarSub
    '�o�����O�t�d�B�z,��ϥΪ̿�ܤu��C�����,�ҹ��������ݤ���غc�ΰʧ@
    Inherits PaintToolBar

    Private ToolbarGrp As Panel
    Private lblToolbarSub() As Label
    Private MaxPixel As Integer = 0
    Private BmpTrans(1) As Bitmap
    Private BmpSpray(2) As Bitmap
    Private bmpPoint As Bitmap

    Protected bTransParent As Boolean = False   '�z����ϥ�
    Protected PenWidth As Integer = 1           '�e���e��
    Protected UserType As Integer = 1           '�X�󫬦�
    Protected EraseSize As Integer = 4          '������j�p
    Protected SprayWidth As Integer = 4         '�Q��d��
    Protected BrushType As Integer = 1          '���ꫬ��
    ' Protected ZoomBase As Integer = 1           '�e�����v  

    Sub New()
        bmpPoint = New Bitmap(1, 1)
        ToolbarGrp = New Panel
        With ToolbarGrp
            .Parent = pnToolbar
            .Location = New Point(8, 210)
            .Size = New Size(42, 68)
            .BorderStyle = BorderStyle.Fixed3D
        End With
        '�o�ت��{���X�w�q,��ϥΪ̿�ܤu��C�\���,�һݭn�����ݸ�T
    End Sub
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '���B�w�q��ϥΪ̿�ܤF�u��C�����,�Ҳ��ͪ����ݤ���.
        '�b�Y�Ө��׬ݨ�,�o�]��O���ݤ��󪺫غc��,���ѩ󥦤��O�ѫغc������
        '�G�]�i�H�����ʺA����.

        MyBase.ToolbarRadbtnSelect(obj, ea)

        DisposelblToolbarSub()                       '������ݤ���
        Dim i As Integer
        Select Case ToolbarCommand
            Case 1, 2, 10   '�s��,��r
                MaxPixel = 1
                ReDim lblToolbarSub(MaxPixel)
                Dim strImageName() As String = {"NoTransParent.bmp", "TransParent.bmp"}
                For i = 0 To MaxPixel
                    BmpTrans(i) = New Bitmap(Me.GetType(), strImageName(i))
                    BmpTrans(i).MakeTransparent(DefaultBackColor) '�w�q�z����,�����ο�ܩҷǳ�
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(39, 29)
                        .Location = New Point(0, 2 + i * 29)
                        .Image = BmpTrans(i)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf TransParentSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf TransParentSelectOnClick
                Next i

            Case 3        '�����
                MaxPixel = 3
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(14, 14)
                        .Location = New Point(10, 2 + i * 15)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf EraseSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf EraseSelectOnClick
                Next i

            Case 6       '��j��
                MaxPixel = 3
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 15)
                        .Location = New Point(3, 2 + i * 15)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf ZoomSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf ZoomSelectOnClick
                Next i

            Case 8       '����
                MaxPixel = 11
                ReDim lblToolbarSub(MaxPixel)
                Dim cx() As Integer = {0, 12, 24, 0, 12, 24, 0, 12, 24, 0, 12, 24}
                Dim cy() As Integer = {0, 0, 0, 12, 12, 12, 24, 24, 24, 36, 36, 36}

                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(12, 12)
                        .Location = New Point(cx(i) + 1, cy(i) + (i + 1))
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf BrushSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf BrushSelectOnClick
                Next i

            Case 9       '�Q��  
                MaxPixel = 2
                ReDim lblToolbarSub(MaxPixel)
                Dim SprayWidth() As Integer = {17, 22, 26}
                Dim SprayLocation() As Point = {New Point(0, 5), _
                                                New Point(17, 5), New Point(3, 29)}
                For i = 0 To MaxPixel
                    BmpSpray(i) = _
                    New Bitmap(Me.GetType(), "Spray" & (i + 1).ToString & ".bmp")
                    BmpSpray(i).MakeTransparent(DefaultBackColor)

                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(SprayWidth(i), 24)
                        .Location = SprayLocation(i)
                        .Image = BmpSpray(i)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf SprayWidthSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf SprayWidthSelectOnClick
                Next i

            Case 11, 12        '�u��
                MaxPixel = 4
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 12)
                        .Location = New Point(2, 2 + i * 12)
                    End With

                    AddHandler lblToolbarSub(i).Paint, AddressOf LineSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf LineSelectOnClick
                Next i

            Case 13, 14, 15, 16    '�X��˦�
                MaxPixel = 2
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 20)
                        .Location = New Point(2, 20 * i + 2)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf RectSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf RectSelectOnClick
                Next i
            Case Else
        End Select
        '�b16�Ӥu��C����,�`�@�|����8�ؤ��P�����ݤ���(�]�t�ť�)
        '�o�ت��{���X,�����ת����Ʃʦs�b,���b�g�L�Ҽ{��.�Ȯɤ�����X��
        '�����o�Ǫ��ݤ���,�]���\�h���ۦP���a��.�o�ئX�֪���,���M�i�H�j�T
        '��C�{���X,���]�|���\Ū�{���X,�ܪ��x���\�h.

    End Sub
    '�H�U���C�դ覡,�禳�ܦh�۪񤧳B,�P�ˬ��{�����c�O�d�Ӥ��@�X�ְʧ@,�Ȧb�S�O�[�H����.
    Private Sub TransParentSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim SubSelect As Integer
        If Not bTransParent Then SubSelect = 1 Else SubSelect = 2

        SetBackColor(SubSelect, CInt(lblPaint.Name))
        '�o��ø�s���󪺤覡,���@�I���P���O.��ܨϥΪ̩ҿ�ܪ�����
        '�O��OnPaint ���欰,�� BackColor �Ӫ��,�Ӥ��O�� OnClick �ӫإ�
    End Sub
    Private Sub TransParentSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)

        If CInt(lblSelect.Name) = 1 Then
            bTransParent = False
        Else
            bTransParent = True
        End If
        ToolbarGrp.Refresh()        '�o�بϥ� Refresh.�O���FĲ�o������ OnPaint �ϬM�ϥΪ̪��ܧ�
    End Sub

    Private Sub EraseSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor((EraseSize - 2) \ 2, CInt(lblPaint.Name))

        Dim Range As Integer = CInt(lblPaint.Name) * 2 + 2
        Dim sy As Integer = lblPaint.Size.Height \ 2 - Range \ 2
        pea.Graphics.FillRectangle(New SolidBrush(clr), sy, sy, Range, Range)
    End Sub
    Private Sub EraseSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)

        EraseSize = CInt(lblSelect.Name) * 2 + 2
        ToolbarGrp.Refresh()
    End Sub

    Private Sub SprayWidthSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        SetBackColor(SprayWidth \ 4, CInt(lblPaint.Name))
        '�o�ت���ܦ欰�������T���a��,�b��ۤ�,��ܰϪ��ϧ����ӷ|�Q�ϦV�B�z.
    End Sub
    Private Sub SprayWidthSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        SprayWidth = Val(lblSelect.Name) * 4

        ToolbarGrp.Refresh()
    End Sub

    Private Sub ZoomSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim emuBase() As Integer = {1, 2, 6, 8}
        Dim SubSelect As Integer = Array.IndexOf(emuBase, ZoomBase) + 1
        Dim clr As Color = SetBackColor(SubSelect, CInt(lblPaint.Name))

        Dim Fnt As New Font("�s�ө���", 7)          '�Ҽ{���ܧ󪺥i��
        pea.Graphics.DrawString(emuBase(CInt(lblPaint.Name) - 1).ToString & "x", _
                                 Fnt, New SolidBrush(clr), New PointF(5, 3))
    End Sub
    Private Sub ZoomSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        Dim emuBase() As Integer = {1, 2, 6, 8}
        If ZoomBase = emuBase(CInt(lblSelect.Name) - 1) Then Return '�I��ۦP�ﶵ��

        ZoomBase = emuBase(CInt(lblSelect.Name) - 1)
        ToolbarGrp.Refresh()
        ZoomBaseChange()

        radbtnToolbar(LastCommand - 1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(LastCommand - 1), ea.Empty)  '��ܤW�@�өR�O
    End Sub

    Private Sub BrushSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)
        Dim clr As Color = SetBackColor(BrushType, CInt(lblPaint.Name))

        Select Case CInt(lblPaint.Name)
            Case 1
                MyDrawEllipse(pea.Graphics, New Point(5, 5), 4, clr)
            Case 2
                MyDrawEllipse(pea.Graphics, New Point(5, 5), 2, clr)
            Case 3
                pea.Graphics.FillRectangle(New SolidBrush(clr), New Rectangle(5, 5, 1, 1))
                '�S��,�o�إH�@�Ӥj�p��1����߯x��,���1��pixel���I
            Case 4
                pea.Graphics.FillRectangle(New SolidBrush(clr), New Rectangle(2, 2, 8, 8))
            Case 5
                pea.Graphics.FillEllipse(New SolidBrush(clr), New Rectangle(3, 3, 5, 5))
            Case 6
                pea.Graphics.DrawRectangle(New Pen(clr), New Rectangle(5, 5, 1, 1))
            Case 7
                pea.Graphics.DrawLine(New Pen(clr), New Point(9, 1), New Point(1, 9))
            Case 8
                pea.Graphics.DrawLine(New Pen(clr), New Point(7, 3), New Point(3, 7))
            Case 9
                pea.Graphics.DrawLine(New Pen(clr), New Point(6, 4), New Point(4, 6))
            Case 10
                pea.Graphics.DrawLine(New Pen(clr), New Point(1, 1), New Point(9, 9))
            Case 11
                pea.Graphics.DrawLine(New Pen(clr), New Point(3, 3), New Point(7, 7))
            Case 12
                pea.Graphics.DrawLine(New Pen(clr), New Point(4, 4), New Point(6, 6))
        End Select
    End Sub
    Private Sub BrushSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        BrushType = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Sub LineSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor(PenWidth, CInt(lblPaint.Name))

        Dim pn As New Pen(clr, CInt(lblPaint.Name))
        Dim sy As Integer = lblPaint.Size.Height \ 2 - (pn.Width \ 2) + 1  '�p��y�а���
        pea.Graphics.DrawLine(pn, 4, sy, 28, sy)
    End Sub
    Private Sub LineSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        PenWidth = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Sub RectSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor(UserType, CInt(lblPaint.Name))

        Select Case CInt(lblPaint.Name)
            Case 1
                pea.Graphics.DrawRectangle(New Pen(clr), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
            Case 2
                pea.Graphics.DrawRectangle(New Pen(clr), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
                pea.Graphics.FillRectangle(New SolidBrush(Color.Gray), 5, 5, lblPaint.Width - 11, lblPaint.Height - 11)
            Case 3
                pea.Graphics.FillRectangle(New SolidBrush(Color.Gray), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
        End Select
    End Sub
    Private Sub RectSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        UserType = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Function SetBackColor(ByVal UserSelect As Integer, ByVal objName As Integer) As Color
        '�o�إH BackColor �ӻs�@�ϥΪ̿�ܪ���᪺����,�åH clr �����Ϭۦ⪺�ĪG
        Dim clr As Color
        If UserSelect = objName Then
            lblToolbarSub(objName - 1).BackColor = Color.DarkBlue
            clr = Color.White
        Else
            lblToolbarSub(objName - 1).BackColor = ToolbarGrp.BackColor
            clr = Color.Black
        End If
        Return clr
    End Function
    Private Sub DisposelblToolbarSub()
        '����Ҧ����ݤ���
        Dim i As Integer
        If MaxPixel = 0 Then Return
        For i = 0 To MaxPixel
            lblToolbarSub(i).Dispose()
        Next i
        MaxPixel = 0
    End Sub

    Protected Sub MyDrawEllipse(ByVal grfx As Graphics, ByVal center As Point, _
                               ByVal radius As Integer, ByVal clr As Color)
        Dim i As Double
        Dim last As Point        
        Dim pt(360) As Point
        Dim Count As Integer = 0
        For i = 0 To 2 * Math.PI Step 2 * Math.PI / 360
            last.X = center.X + CInt(radius * Math.Cos(i))
            last.Y = center.Y + CInt(radius * Math.Sin(i))
            pt(Count) = last           
            Count += 1
        Next i
        grfx.FillClosedCurve(New SolidBrush(clr), pt)
        '������n�ۤv���I�e��O?,�o�O�]�����ӫܩ_�Ǫ��z��
        'Graphics.DrawEllipse �o�өR�O,��������,�b�������x�ΪŶ��ܤp��,�e�X�Ӫ���
        '��,�@�I������...><�ܩ󬰤��򤣪����b�j�餤�e�u�Y�i,�D�n�O��Graphics ���ʧ@
        '���Ǯį�W���ü{,���Ʊ�h���I�s Graphics
    End Sub
    
End Class
