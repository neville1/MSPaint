Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class RectControlDot
    '此類別定義一個由外圍控制點包圍,可供改變大小的控制項
    '本控制項僅有大小位置變更的實作,故不能直接使用,必需繼承延伸.
    Inherits Control

    Private DotWithPanel(7) As ControlDot
    Private sz As Size
    Private bTracking As Boolean = False
    Private RectEdit As Rectangle

    Protected bDraw As Boolean = False         '定義當控制項大小變更期間,外觀表面形式

    ReadOnly Property Tracking() As Boolean
        Get
            Return bTracking
        End Get
    End Property
    Property Rect() As Rectangle
        Get
            Return RectEdit
        End Get
        Set(ByVal Value As Rectangle)
            RectEdit = Value
        End Set
    End Property
    Sub New()
        '控制項建構元,建構包圍該控制項的定位點
        AddHandler Me.Resize, AddressOf PanelResize

        sz = New Size(Me.Width, Me.Height)

        Dim strCursor() As String = {"ChangeSizeOblique.cur", "ChangeSizeRiseDown.cur", "ChangeSizeRightOblique.cur", _
                                    "ChangeSizeLeftRight.cur", "ChangeSizeLeftRight.cur", "ChangeSizeRightOblique.cur", _
                                    "ChangeSizeRiseDown.cur", "ChangeSizeOblique.cur"}

        Dim DotLo() As Point = {New Point(0, 0), New Point(sz.Width \ 2, 0), New Point(sz.Width - 3, 0), _
                       New Point(0, sz.Height \ 2), New Point(sz.Width - 3, sz.Height \ 2), _
                       New Point(0, sz.Height - 3), New Point(sz.Width \ 2, sz.Height - 3), _
                       New Point(sz.Width - 3, sz.Height - 3)}

        Dim strDotName() As String = {"SW", "S", "SE", "W", "E", "NW", "N", "NE"}

        Dim i As Integer
        For i = 0 To 7
            DotWithPanel(i) = New ControlDot
            With DotWithPanel(i)
                .Parent = Me
                .Name = strDotName(i)
                .Enabled = True
                .Location = DotLo(i)
                .Cursor = New Cursor(Me.GetType(), strCursor(i))
            End With

            AddHandler DotWithPanel(i).MouseDown, AddressOf ControlDotOnMouseDown
            AddHandler DotWithPanel(i).MouseMove, AddressOf ControldotOnMove
            AddHandler DotWithPanel(i).MouseUp, AddressOf PanelSizeChanged
        Next i
    End Sub
    Protected Sub AllControlDotHide()
        '所有控制點隱藏
        Dim i As Integer
        For i = 0 To DotWithPanel.GetUpperBound(0)
            DotWithPanel(i).Hide()
        Next i
    End Sub
    Protected Sub AllControlDotShow()
        '所有控制點顯示
        Dim i As Integer
        For i = 0 To DotWithPanel.GetUpperBound(0)
            DotWithPanel(i).Show()
        Next i
    End Sub
    Private Sub ControlDotOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '變更控制項大小位置行為啟動
        If mea.Button = MouseButtons.Left Then
            bTracking = True
        End If
    End Sub
    Protected Overridable Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '由移動控制點行為,決定控制項大小,位置變更
        If Not bTracking Then Return

        Dim CtrlDot As ControlDot = DirectCast(obj, ControlDot)
        Me.Parent.Refresh()

        Dim poMouse As Point = (Me.Parent.PointToClient(MousePosition))
        Dim RectLo As New Rectangle(Me.Location.X, Me.Location.Y, Me.Size.Width, Me.Size.Height)
        Select Case CtrlDot.Name
            Case "SW"
                RectEdit = New Rectangle(poMouse.X, poMouse.Y, _
                               RectLo.Width + RectLo.X - poMouse.X, RectLo.Height + RectLo.Y - poMouse.Y)
            Case "S"
                RectEdit = New Rectangle(RectLo.X, poMouse.Y, _
                               RectLo.Width, RectLo.Height + RectLo.Y - poMouse.Y)
            Case "SE"
                RectEdit = New Rectangle(RectLo.X, poMouse.Y, _
                               poMouse.X - RectLo.X, RectLo.Height + RectLo.Y - poMouse.Y)
            Case "W"
                RectEdit = New Rectangle(poMouse.X, RectLo.Y, _
                               RectLo.Width + RectLo.X - poMouse.X, RectLo.Height)
            Case "E"
                RectEdit = New Rectangle(RectLo.X, RectLo.Y, poMouse.X - RectLo.X, RectLo.Height)
            Case "NW"
                RectEdit = New Rectangle(poMouse.X, RectLo.Y, _
                               RectLo.Width + RectLo.X - poMouse.X, poMouse.Y - RectLo.Y)
            Case "N"
                RectEdit = New Rectangle(RectLo.X, RectLo.Y, RectLo.Width, poMouse.Y - RectLo.Y)
            Case "NE"
                RectEdit = New Rectangle(RectLo.X, RectLo.Y, _
                               poMouse.X - RectLo.X, poMouse.Y - RectLo.Y)
        End Select
        Dim pn As New Pen(Color.Black)
        pn.DashStyle = Drawing2D.DashStyle.Dot
        Dim Rect = New Rectangle(RectEdit.X - RectLo.X, RectEdit.Y - RectLo.Y, RectEdit.Width - 1, RectEdit.Height - 1)

        If bDraw Then
            Dim grfx As Graphics = Me.CreateGraphics()
            Dim grfxParent As Graphics = Me.Parent.CreateGraphics()
            grfx.DrawRectangle(pn, Rect)
            grfxParent.DrawRectangle(pn, RectEdit.X, RectEdit.Y, RectEdit.Width - 2, RectEdit.Height - 2)
            grfx.Dispose()
            grfxParent.Dispose()
        End If
    End Sub
    Protected Overridable Sub PanelSizeChanged(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '控制項位置變更確定
        If Not bTracking Then Return

        bTracking = False
        Me.Location = New Point(RectEdit.X, RectEdit.Y)
        Me.Size = New Size(RectEdit.Width, RectEdit.Height)
    End Sub
    Protected Overridable Sub PanelResize(ByVal obj As Object, ByVal ea As EventArgs)
        '定位點位置重新定義.
        Dim Sz As New Size(Me.Width, Me.Height)
        Dim DotLo() As Point = {New Point(0, 0), New Point(Sz.Width \ 2, 0), New Point(Sz.Width - 3, 0), _
                       New Point(0, Sz.Height \ 2), New Point(Sz.Width - 3, Sz.Height \ 2), _
                       New Point(0, Sz.Height - 3), New Point(Sz.Width \ 2, Sz.Height - 3), _
                       New Point(Sz.Width - 3, Sz.Height - 3)}
        Dim i As Integer
        For i = 0 To 7
            DotWithPanel(i).Location = DotLo(i)
        Next i
    End Sub
End Class
