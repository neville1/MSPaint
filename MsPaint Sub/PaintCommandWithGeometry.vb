Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms

Public Class PaintCommandWithGeometry
    '�����O��@�u��C������X��ϧάۦ������B�z.
    '�����O�]��欰�۪�,�����ƻ�����ۦ��欰.
    Inherits PaintCommandWithBrush

    Private bTracking As Boolean = False
    Private ptLast, ptNew As Point
    Private clrWork As Color

    Private BezierCount As Integer = 0
    Private BezierPoint(3) As Point

    Private bPolyStart As Boolean = False
    Private PolyPoint As ArrayList
    Private ReadOnly EumType() As Char = {"P", "R", "F"}

    Private pathForRoundedRect As GraphicsPath    
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '�H�мg�u��C����ƥ�,�w�q�u��C����Ҳ��ͪ��欰�e���Ψҥ~�B�z.

        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 11
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawLineOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawLineOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawLineOnMouseUp

            Case 12                   '��ϥΪ̴��e���_ø�ϰʧ@�ɪ��B�z
                If BezierCount <> 0 Then
                    BezierCount = 3
                    bTracking = True
                    Dim mea As MouseEventArgs
                    DrawBezierOnMouseUp(obj, mea)
                End If
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawBezierOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawBezierOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawBezierOnMouseUp
            Case 13
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawRectangleOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawRectangleOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawRectangleOnMouseUp
            Case 14
                If bPolyStart Then   '��ϥΪ̴��e���_ø�ϰʧ@�ɪ��B�z
                    Dim mea As MouseEventArgs
                    Dim ptStart As Point = DirectCast(PolyPoint(0), Point)
                    ptNew = ptStart
                    bTracking = True
                    DrawPolyOnMouseUp(obj, mea)
                End If
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawPolyOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawPolyOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawPolyOnMouseUp
            Case 15
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawEllipseOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawEllipseOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawEllipseOnMouseUp
            Case 16
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawRoundedRectOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawRoundedRectOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawRoundedRectOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 11  '�e�u
                AddHandler pnWorkarea.MouseDown, AddressOf DrawLineOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawLineOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawLineOnMouseUp
            Case 12  '�e Bezier ���u
                AddHandler pnWorkarea.MouseDown, AddressOf DrawBezierOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawBezierOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawBezierOnMouseUp
            Case 13  '�e�x��
                AddHandler pnWorkarea.MouseDown, AddressOf DrawRectangleOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawRectangleOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawRectangleOnMouseUp
            Case 14  '�e�h���
                AddHandler pnWorkarea.MouseDown, AddressOf DrawPolyOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawPolyOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawPolyOnMouseUp
            Case 15  '�e���
                AddHandler pnWorkarea.MouseDown, AddressOf DrawEllipseOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawEllipseOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawEllipseOnMouseUp
            Case 16  '�e�����x��
                AddHandler pnWorkarea.MouseDown, AddressOf DrawRoundedRectOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawRoundedRectOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawRoundedRectOnMouseUp
        End Select
    End Sub

    Private Sub DrawLineOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        'ø�s���u,���������̰򥻫��A
    End Sub
    Private Sub DrawLineOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        SetSbpRect(ptLast, ptNew)

        pnWorkarea.Refresh()

        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLine(New Pen(clrWork, PenWidth), ptLast, ptNew)
        grfxFore.Dispose()
    End Sub
    Private Sub DrawLineOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        ClearSbpRect()

        Dim GrapMethod As New GraphicsItem

        With GrapMethod
            .DataType = "P"
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddLine(ptLast, ptNew)
            .Data.StartFigure()
        End With

        GraphicsData.Add(GrapMethod)
        '�����Ҩϥμз� path �ӰO�����ϥΪ�ø�ϰʧ@
    End Sub

    Private Sub DrawBezierOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        BezierCount += 1
        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        Select Case BezierCount
            Case 1
                BezierPoint(0) = ptLast
                BezierPoint(1) = ptLast
            Case 2
                BezierPoint(1) = ptLast
                DrawBezierOnMouseMove(obj, mea)        '�b�ϥΪ̿�ܤϦ��I��,�Y�ɤ����ϧ��ܧ�
            Case 3
                BezierPoint(2) = ptLast
                DrawBezierOnMouseMove(obj, mea)
        End Select
        'Bezier ���u�o�O�@�إѥ|���I�Ҳզ������u,���L�ѹ껡,���S��Ĳ�L�q�����Ѫ��ϥΪ̦Ө�
        '�ӤH�Z�h��,�p�e�a�� Bezier ���u,���h�֤H����F�Ѩ�Ϊk.
    End Sub
    Private Sub DrawBezierOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))

        Select Case BezierCount
            Case 1
                BezierPoint(2) = ptNew
                BezierPoint(3) = ptNew
            Case 2
                BezierPoint(1) = ptNew
            Case 3
                BezierPoint(2) = ptNew
        End Select

        SetSbpRect(ptLast, ptNew)

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawBezier(New Pen(clrWork, PenWidth), BezierPoint(0), BezierPoint(1), _
                            BezierPoint(2), BezierPoint(3))
        grfxFore.Dispose()
    End Sub
    Private Sub DrawBezierOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        ClearSbpRect()

        If BezierCount < 3 Then Return ' ø�Ϧ欰�|���פ�

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "P"
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddBezier(BezierPoint(0), BezierPoint(1), BezierPoint(2), BezierPoint(3))
            .Data.StartFigure()
        End With
        GraphicsData.Add(GrapMethod)
        BezierCount = 0

    End Sub

    Private Sub DrawRectangleOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub DrawRectangleOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        Dim Rect As Rectangle = SetSbpRect(ptLast, ptNew)

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        Select Case Usertype
            Case 1
                grfxFore.DrawRectangle(New Pen(clrWork, PenWidth), Rect)
            Case 2
                grfxFore.FillRectangle(New SolidBrush(clrBack), Rect)
                grfxFore.DrawRectangle(New Pen(clrWork, PenWidth), Rect)
            Case 3
                grfxFore.FillRectangle(New SolidBrush(clrBack), Rect)
        End Select
        grfxFore.Dispose()
    End Sub
    Private Sub DrawRectangleOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        Dim Rect As Rectangle = SetSbpRect(ptLast, ptNew)
        ClearSbpRect()

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            GrapMethod.DataType = EumType(UserType - 1)
            GrapMethod.FillColor = clrBack
            GrapMethod.Color = clrWork
            GrapMethod.pWidth = PenWidth
            GrapMethod.Data.AddRectangle(Rect)
        End With
        GraphicsData.Add(GrapMethod)
    End Sub

    Private Sub DrawPolyOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        bTracking = True

        If Not bPolyStart Then
            PolyPoint = New ArrayList
            PolyPoint.Add(ptLast)
            bPolyStart = True         '�h���ø�s�欰�}�l
        Else
            ptLast = ptNew
        End If
    End Sub
    Private Sub DrawPolyOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        SetSbpRect(ptLast, ptNew)

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLine(New Pen(clrWork, PenWidth), ptLast, ptNew)
        grfxFore.Dispose()
    End Sub
    Private Sub DrawPolyOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return

        ClearSbpRect()

        Dim TempGrapMethod As New GraphicsItem
        With TempGrapMethod
            .DataType = "P"
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddLine(ptLast, ptNew)
        End With

        GraphicsData.Add(TempGrapMethod)

        PolyPoint.Add(ptNew)        '�إ�ø�s�h��Τ��Ҳ��ͪ��{�ɩʽu�q

        Dim ptStart As Point = DirectCast(PolyPoint(0), Point)

        If (Math.Abs(ptNew.X - ptStart.X) > PenWidth AndAlso _
           Math.Abs(ptNew.Y - ptStart.Y) > Penwidth) Then Return

        PolyPoint(PolyPoint.Count - 1) = PolyPoint(0)
        Dim apt() As Point = CType(PolyPoint.ToArray(GetType(Point)), Point())

        GraphicsData.RemoveRange(GraphicsData.Count - _
                                  apt.GetUpperBound(0), apt.GetUpperBound(0))
        '�R���{�ɩʽu�q 

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = EumType(UserType - 1)
            .FillColor = clrBack
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddPolygon(apt)
        End With

        GraphicsData.Add(GrapMethod)
        bPolyStart = False
        PolyPoint.Clear()

        pnWorkarea.Refresh()
    End Sub

    Private Sub DrawEllipseOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub DrawEllipseOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        Dim Rect As Rectangle = SetSbpRect(ptLast, ptNew)

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        Select Case Usertype
            Case 1
                grfxFore.DrawEllipse(New Pen(clrWork, PenWidth), Rect)
            Case 2
                grfxFore.FillEllipse(New SolidBrush(clrBack), Rect)
                grfxFore.DrawEllipse(New Pen(clrWork, PenWidth), Rect)
            Case 3
                grfxFore.FillEllipse(New SolidBrush(clrBack), Rect)
        End Select
        grfxFore.Dispose()
    End Sub
    Private Sub DrawEllipseOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        Dim Rect As Rectangle = SetSbpRect(ptLast, ptNew)

        ClearSbpRect()

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = EumType(UserType - 1)
            .FillColor = clrBack
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddEllipse(Rect)
        End With
        GraphicsData.Add(GrapMethod)
    End Sub

    Private Sub DrawRoundedRectOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()

        clrWork = SelectUserColor(mea)
        bTracking = True
        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub DrawRoundedRectOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        Dim rect As Rectangle = SetSbpRect(ptLast, ptNew)

        Dim RangeWidth As Integer = Math.Abs(ptLast.X - ptNew.X)
        Dim RangeHeight As Integer = Math.Abs(ptLast.Y - ptNew.Y)

        Dim Szf As New Size(RangeWidth, RangeHeight)
        Szf.Width = Math.Min(RangeWidth, 18)
        Szf.Height = Math.Min(RangeHeight, 18)
        '�o�ئ��ӷN�Q���쪺����,���Ӥp�̥H��,�ҿ׶ꨤ�x��,�ꨤ���j�p���өM�x�Ϊ��j�p����.
        '�i�O�p�e�a��ڪ����浲�G,�ꨤ���j�p,�O�Q����b�@�ӽd�򤺪�.

        pathForRoundedRect = New GraphicsPath
        pathForRoundedRect.AddLine(rect.Left + Szf.Width \ 2, rect.Top, rect.Right - Szf.Width \ 2, rect.Top)
        If Szf.Width > 0 AndAlso Szf.Height > 0 Then
            pathForRoundedRect.AddArc(rect.Right - Szf.Width, rect.Top, Szf.Width, Szf.Height, 270, 90)
        End If
        pathForRoundedRect.AddLine(rect.Right, rect.Top + Szf.Height \ 2, rect.Right, rect.Bottom - Szf.Height \ 2)
        If Szf.Width > 0 AndAlso Szf.Height > 0 Then
            pathForRoundedRect.AddArc(rect.Right - Szf.Width, rect.Bottom - Szf.Height, Szf.Width, Szf.Height, 0, 90)
        End If
        pathForRoundedRect.AddLine(rect.Right - Szf.Width \ 2, rect.Bottom, rect.Left + Szf.Width \ 2, rect.Bottom)
        If Szf.Width > 0 AndAlso Szf.Height > 0 Then
            pathForRoundedRect.AddArc(rect.Left, rect.Bottom - Szf.Height, Szf.Width, Szf.Height, 90, 90)
        End If
        pathForRoundedRect.AddLine(rect.Left, rect.Bottom - Szf.Height \ 2, rect.Left, rect.Top + Szf.Height \ 2)
        If Szf.Width > 0 AndAlso Szf.Height > 0 Then
            pathForRoundedRect.AddArc(rect.Left, rect.Top, Szf.Width, Szf.Height, 180, 90)
        End If
        '�̷Ӷ�������V,�N�ꨤ�x�Ϊ����u�Ϊ��u�[�J���|

        pnWorkarea.Refresh()
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        Select Case Usertype
            Case 1
                grfxFore.DrawPath(New Pen(clrWork, PenWidth), pathForRoundedRect)
            Case 2
                grfxFore.FillPath(New SolidBrush(clrBack), pathForRoundedRect)
                grfxFore.DrawPath(New Pen(clrWork, PenWidth), pathForRoundedRect)
            Case 3
                grfxFore.FillPath(New SolidBrush(clrBack), pathForRoundedRect)
        End Select
        grfxFore.Dispose()
    End Sub
    Private Sub DrawRoundedRectOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        ClearSbpRect()

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = EumType(UserType - 1)
            .FillColor = clrBack
            .Color = clrWork
            .pWidth = PenWidth
            .Data = pathForRoundedRect
        End With
        GraphicsData.Add(GrapMethod)
    End Sub

    Protected Function SetSbpRect(ByVal poSt As Point, ByVal poEnd As Point) As Rectangle
        '�N�ϧΪ�ø�s�d��,��ܦb���A�C�W,�ñN��~��x�Φ�m�Ǧ^
        Dim RangeWidth As Integer = Math.Abs(poSt.X - poEnd.X)
        Dim RangeHeight As Integer = Math.Abs(poSt.Y - poEnd.Y)

        sbpRectangle.Text = RangeWidth.ToString & "x" & RangeHeight.ToString

        Return New Rectangle(Math.Min(poSt.X, poEnd.X), _
                             Math.Min(poSt.Y, poEnd.Y), RangeWidth, RangeHeight)
    End Function
    Protected Sub ClearSbpRect()
        '�M�����A�C������ø�s�d���T,�ð���ƹ��l��
        bTracking = False
        sbpRectangle.Text = ""
    End Sub
End Class
