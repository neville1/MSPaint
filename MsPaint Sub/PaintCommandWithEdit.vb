Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintCommandWithEdit
    '�����O��@�u��C������s�������B�z.
    Inherits PaintCommandWithGeometry

    Protected bEdit As Boolean = False             '�s�趵�s�b?
    Protected pnEditarea As EditPanel              '�s�趵����
    Protected ConMenuEdit As ContextMenu           '�s�趵���e�\���  
    
    Private bTracking As Boolean = False
    Private ptLast, ptNew As Point
    Private Rect As Rectangle
    Private ptGetRange As ArrayList
    Private ptMin, ptMax As Point
    Sub New()
        ConMenuEdit = New ContextMenu
    End Sub

    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '�H�мg�u��C����ƥ�,�w�q�u��C����Ҳ��ͪ��欰�e���Ψҥ~�B�z.

        If bEdit Then DisposeEditPanel() '��ϥΪ̤����u��C�����,�ҳB�z���欰

        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 1
                pnWorkarea.ContextMenu = Nothing
                RemoveHandler pnWorkarea.MouseDown, AddressOf GetEditRangeOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf GetEditRangeOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf GetEditRangeOnMouseUp
            Case 2
                pnWorkarea.ContextMenu = Nothing
                RemoveHandler pnWorkarea.MouseDown, AddressOf GetEditRectOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf GetEditRectOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf GetEditRectOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 1   '��ܥ��N�d��
                pnWorkarea.ContextMenu = ConMenuEdit
                AddHandler pnWorkarea.MouseDown, AddressOf GetEditRangeOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf GetEditRangeOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf GetEditRangeOnMouseUp
            Case 2   '��ܯx�νd��
                pnWorkarea.ContextMenu = ConMenuEdit
                AddHandler pnWorkarea.MouseDown, AddressOf GetEditRectOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf GetEditRectOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf GetEditRectOnMouseUp
        End Select
    End Sub

    Private Sub GetEditRangeOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If mea.Button <> MouseButtons.Left Then Return

        If bEdit Then DisposeEditPanel()
        ClearCommandRecord()

        bTracking = True
        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        ptMax = ptLast
        ptMin = ptLast

        ptGetRange = New ArrayList
        ptGetRange.Add(ptLast)
    End Sub
    Private Sub GetEditRangeOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        ptMin = New Point(Math.Min(ptNew.X, ptMin.X), Math.Min(ptNew.Y, ptMin.Y))
        ptMax = New Point(Math.Max(ptNew.X, ptMax.X), Math.Max(ptNew.Y, ptMax.Y))
        '���oPoint() �����w���̤j�d��

        SetSbpRect(ptMax, ptMin)
        ptGetRange.Add(ptNew)

        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLine(New Pen(Color.Black, 2), ptLast, ptNew)
        ptLast = ptNew
        '�b�o�ؤS���ͤ@�ӻP�p�e�a�欰���Ū�����,�p�e�a�O�ϥΥثe��m���Ϭۦ�,�@��ø�s�I��
        '�C��,���b�o�ӵ{����,�p�G�ϥΤϬۦ��ø�s,���w�y���į�C��        
    End Sub
    Private Sub GetEditRangeOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        bTracking = False

        ptGetRange.Add(ptGetRange(0))
        Rect = SetSbpRect(ptMax, ptMin)
        If Rect.Width < 3 OrElse Rect.Height < 3 Then Return '����̤p����d��

        Dim apt() As Point = CType(ptGetRange.ToArray(GetType(Point)), Point())
        Dim path As New GraphicsPath
        path.AddPolygon(apt)

        CreateEditPanel(Rect, path, True)
        ClearSbpRect()
    End Sub

    Private Sub GetEditRectOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If mea.Button <> MouseButtons.Left Then Return

        If bEdit Then DisposeEditPanel()
        ClearCommandRecord()

        bTracking = True
        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub GetEditRectOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        Rect = SetSbpRect(ptLast, ptNew)

        Dim pn As New Pen(Color.Black)
        pn.DashStyle = DashStyle.Dash

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawRectangle(pn, Rect)
        grfxFore.Dispose()
    End Sub
    Private Sub GetEditRectOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)    
        bTracking = False

        If Rect.Width < 3 OrElse Rect.Height < 3 Then Return '����̤p����d��

        Dim path As New GraphicsPath
        path.AddRectangle(Rect)

        CreateEditPanel(Rect, path, bTransParent)
        ClearSbpRect()
    End Sub

    Protected Sub CreateEditPanel(ByVal Rect As Rectangle, ByVal Clippath As GraphicsPath, _
                                ByVal bFromTransParent As Boolean)
        '�إ߽s�趵����,�ö�R�]�إ߽s�趵����Ӳ��ͻݲ���������
        Dim bmpBack As Bitmap = WorkareaChangebmp()
        Dim Bmp As New Bitmap(Rect.Width, Rect.Height)
        Dim grfx As Graphics = Graphics.FromImage(Bmp)
        grfx.SetClip(Clippath)
        grfx.TranslateClip(-Rect.X, -Rect.Y)
        grfx.DrawImage(bmpBack, CInt(-Rect.X), CInt(-Rect.Y), bmpBack.Width, bmpBack.Height)
        grfx.Dispose()   '���o�����ϧ�

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "F"
            .FillColor = clrBack
            .Data.AddPath(Clippath, False)
        End With
        GraphicsData.Add(GrapMethod)      '�N�ɮ��d����
        Rect = New Rectangle(Rect.X * ZoomBase, Rect.Y * ZoomBase, _
                             Rect.Width * ZoomBase, Rect.Height * ZoomBase)  ' Rect * ZoomBase

        pnEditarea = New EditPanel(Rect, Bmp, bTransParent, bFromTransParent, clrback)
        pnEditarea.Parent = pnWorkarea
        pnEditarea.ContextMenu = ConMenuEdit
        Bmp.Dispose()
        bEdit = True
        pnWorkarea.Refresh()    '�إ߰_�s�趵����
        '�o�� EditPanel ���غc��,����ΤF�Ӧh���Ѽ�.�ѩ�z���⪺���D,�٨S���o��ѨM.
        '���F���ɤ@�I�į�,�u�n�� bFromTransParent ��@�ѼƾɤJ.
    End Sub
    Protected Sub DisposeEditPanel()
        '����s�趵����ö�Jø�Ϧ欰��

        Dim ptstart As New Point(pnEditarea.Location.X, pnEditarea.Location.Y)

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "B"
            .Image = pnEditarea.BmpData
            .ImageLocation = New Point(ptstart.X \ ZoomBase, ptstart.Y \ ZoomBase)
            .ImageSize = New Size(pnEditarea.Width \ ZoomBase, pnEditarea.Height \ ZoomBase)
        End With

        GraphicsData.Add(GrapMethod)
        pnEditarea.Dispose()
        bEdit = False
    End Sub
    Protected Sub CreateEditPanelWithBmp(ByVal bmp As Bitmap, ByVal FormTransParent As Boolean)
        '������ bitmap �إ߽s�趵����,�ëD��ø�Ϥu�@�ϫإ�
        Dim Rect As New Rectangle(0, 0, bmp.Width * ZoomBase, bmp.Height * ZoomBase)
        pnEditarea = New EditPanel(Rect, bmp, bTransParent, FormTransParent, clrBack)
        pnEditarea.Parent = pnWorkarea
        pnEditarea.ContextMenu = ConMenuEdit
        bmp.Dispose()
        bEdit = True
    End Sub
End Class
