Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintMainArea
    Inherits Form
    '�����O�w�q����ø�Ϥu�@�ϥ~�[�Ϊ��A�ܧ�B�z�覡

    Protected strProgName As String = "�p�e�a���s��"

    Protected bmpPic As Bitmap                '���h�ϧ�
    ' Protected grfxFore As Graphics            '�e���e��(�{�ɩ�)
    Protected szBmp As New Size(512, 384)     '�ϧΤj�p(���ȬO������غc)
    Protected bToolbar, bColorbar, bStatusbar, bTextFrom As Boolean   '�t�m�}��
    Protected pnWorkarea As DoubleBufferPanel    'ø�Ϥu�@��
    Protected ZoomBase As Integer = 1            'ø�ϰ��Y���  

    Private pnAll As Panel                     '���h
    Private DotTL, DotTM, DotTR As ControlDot    '�u�@�Ϥj�p�ܧ󱱨��I
    Private DotML, DotMR As ControlDot
    Private DotDL, DotDM, DotDR As ControlDot

    Sub New()
        'ø�Ϥu�@�ϫغc��
        Dim tx, cy, sy As Integer
        If bToolbar Then tx = 60 Else tx = 0
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0

        Me.Icon = New Icon(GetType(PaintMainArea), "MsPaintTitle.ico")

        pnAll = New Panel
        With pnAll
            .Parent = Me
            .Name = "pnAll"
            .BackColor = Color.Gray
            .BorderStyle = BorderStyle.Fixed3D
            .AutoScroll = True
            .Location = New Point(tx, 0)
            .Size = New Size(ClientSize.Width - tx, ClientSize.Height - cy - sy)
        End With

        pnWorkarea = New DoubleBufferPanel
        With pnWorkarea
            .Parent = pnAll
            .Name = "pnWorkarea"
            .BorderStyle = BorderStyle.None
            .Location = New Point(3, 3)           
            .BackColor = Color.White
            .Size = szBmp
        End With

        'grfxFore = pnWorkarea.CreateGraphics()

        CreateControlDot()    '�إߤu�@�ϥ|�P����j�p�Ϊ������I
        AddHandler pnWorkarea.Resize, AddressOf WorkareaOnReSize
        AddHandler Me.SizeChanged, AddressOf FormSizeChanged   
        
    End Sub
    Protected Overridable Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '�w�q�]��ı�����ܧ�(�u��C,�����)ø�Ϥu�@�Ϯy�аt�m
        Dim tx, cy, sy As Integer
        If bToolbar Then tx = 60 Else tx = 0
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0
        pnAll.Size = New Size(ClientSize.Width - tx, ClientSize.Height - cy - sy)
        pnAll.Location = New Point(tx, 0)
    End Sub
    Protected Sub WorkareaOnReSize(ByVal obj As Object, ByVal ea As EventArgs)
        'ø�Ϥu�@�Ϥj�p�ܧ��,�~�򱱨��I�y�аt�m
        DotTL.Location = New Point(0, 0)
        DotTM.Location = New Point(pnWorkarea.Width \ 2, 0)
        DotTR.Location = New Point(pnWorkarea.Width + 3, 0)
        DotML.Location = New Point(0, pnWorkarea.Height \ 2)
        DotMR.Location = New Point(pnWorkarea.Width + 3, pnWorkarea.Height \ 2)
        DotDL.Location = New Point(0, pnWorkarea.Height + 3)
        DotDM.Location = New Point(pnWorkarea.Width \ 2, pnWorkarea.Height + 3)
        DotDR.Location = New Point(pnWorkarea.Width + 3, pnWorkarea.Height + 3)
    End Sub
    Private Sub CreateControlDot()
        '�إ�ø�Ϥu�@�ϥ~�򱱨��I,�H�վ�ø�Ϥu�@�Ϥj�p
        DotTL = New ControlDot
        With DotTL
            .Parent = pnAll
            .Location = New Point(0, 0)
        End With
        DotTM = New ControlDot
        With DotTM
            .Parent = pnAll
            .Location = New Point(pnWorkarea.Size.Width \ 2, 0)
        End With

        DotTR = New ControlDot
        With DotTR
            .Parent = pnAll
            .Location = New Point(pnWorkarea.Size.Width + 3, 0)
        End With

        DotML = New ControlDot
        With DotML
            .Parent = pnAll
            .Location = New Point(0, pnWorkarea.Height \ 2)
        End With

        DotMR = New ControlDot
        With DotMR
            .Parent = pnAll
            .Name = "LeftRight"
            .Location = New Point(pnWorkarea.Width + 3, pnWorkarea.Height \ 2)
            .Enabled = True
            .Cursor = New Cursor(Me.GetType(), "ChangeSizeLeftRight.cur")
        End With
        AddHandler DotMR.MouseMove, AddressOf ControldotOnMove
        AddHandler DotMR.MouseUp, AddressOf WorkAreaSizeChanged

        DotDL = New ControlDot
        With DotDL
            .Parent = pnAll
            .Location = New Point(0, pnWorkarea.Height + 3)
        End With

        DotDM = New ControlDot
        With DotDM
            .Parent = pnAll
            .Name = "RiseDown"
            .Location = New Point(pnWorkarea.Width \ 2, pnWorkarea.Height + 3)
            .Enabled = True
            .Cursor = New Cursor(Me.GetType(), "ChangeSizeRiseDown.cur")
        End With
        AddHandler DotDM.MouseMove, AddressOf ControldotOnMove
        AddHandler DotDM.MouseUp, AddressOf WorkAreaSizeChanged

        DotDR = New ControlDot
        With DotDR
            .Parent = pnAll
            .Name = "Oblique"
            .Location = New Point(pnWorkarea.Width + 3, pnWorkarea.Height + 3)
            .Enabled = True
            .Cursor = New Cursor(Me.GetType(), "ChangeSizeOblique.cur")
        End With
        AddHandler DotDR.MouseMove, AddressOf ControldotOnMove
        AddHandler DotDR.MouseUp, AddressOf WorkAreaSizeChanged
        '�o�ت��{���X,�������u�@�Ϫ��t�m,���O�D�n�O�b�w�q�]��u�@�Ϫ������I�Ψ�ʧ@.
        '�o�ا@�k�Pı�W�ܲ©�.�ӥB�PRectControlDot.vb ���\�h�������B.
        '���յۧ�o�Ǳ����I�����,���O�a�Ӫ��o�O�\�h����ƥ洫.
        '���ӷ|����n���@�k�~�O.

        '�o�ت��{���X,�̷Ӱ򥻭�h,���ӬO�Q��b�u�@�ϫغc����
        '���B��¦]�����Ʊ�u�@�ϫغc���ݨӤӹL�c��,�G�N�䲾�X�غc��
    End Sub
    Protected Overridable Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '�w�q��ϥΪ̦b�~�򱱨��I�W�@�즲�欰��,�Ұ����B�z
        If mea.Button <> MouseButtons.Left Then Return

        Dim CtrlDot As ControlDot = DirectCast(obj, ControlDot)
        pnAll.Refresh()

        Dim Sz As New Size(pnAll.PointToClient(MousePosition))
        Dim pn As New Pen(Color.Black)
        pn.DashStyle = Drawing2D.DashStyle.Dot
        Dim x, y As Integer
        Select Case CtrlDot.Name
            Case "LeftRight"
                x = Sz.Width
                y = szBmp.Height
            Case "RiseDown"
                x = szBmp.Width
                y = Sz.Height
            Case "Oblique"
                x = Sz.Width
                y = Sz.Height
        End Select
        Dim grfxWork As Graphics = pnWorkarea.CreateGraphics()
        Dim grfxall As Graphics = pnAll.CreateGraphics()

        grfxWork.DrawRectangle(pn, 0, 0, x - 1, y - 1)
        grfxall.DrawRectangle(pn, 4, 4, x - 1, y - 1)
        grfxall.Dispose()
        szBmp = New Size(x, y)

        '�ϧ��Y����ܧ�᪺�@�k���[�J
    End Sub
    Protected Overridable Sub WorkAreaSizeChanged(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '��ø�Ϥu�@�Ϥj�p�����ܧ�n�D��.
        '�� PaintWithMenuImage.vb �мg,����u�@�Ϲ����ܧ󪺵����Υ[�j.
    End Sub
    Protected Sub ZoomBaseChange()
        '�ܧ�u�@�ϵe���Y���
        Dim Sz As New Size(szBmp.Width * ZoomBase, szBmp.Height * ZoomBase)
        pnall.AutoScroll = False                '�קK�] Scroll �첾,�ɭP�y�л~�t.
        pnWorkarea.Size = Sz
        'grfxFore = pnWorkarea.CreateGraphics()
        'grfxFore.ScaleTransform(ZoomBase, ZoomBase)

        pnWorkarea.Refresh()                    'Ĳ�o On_Paint,�ϬM�Y����ܧ�.          
        pnall.AutoScroll = True
    End Sub
    Protected Function GetGraphicsbyFore() As Graphics
        Dim grfx As Graphics = pnWorkarea.CreateGraphics()
        grfx.ScaleTransform(ZoomBase, ZoomBase)
        Return grfx
    End Function
End Class
