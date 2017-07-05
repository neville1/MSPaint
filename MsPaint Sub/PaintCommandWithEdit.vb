Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintCommandWithEdit
    '本類別實作工具列中關於編輯類的處理.
    Inherits PaintCommandWithGeometry

    Protected bEdit As Boolean = False             '編輯項存在?
    Protected pnEditarea As EditPanel              '編輯項物件
    Protected ConMenuEdit As ContextMenu           '編輯項內容功能表  
    
    Private bTracking As Boolean = False
    Private ptLast, ptNew As Point
    Private Rect As Rectangle
    Private ptGetRange As ArrayList
    Private ptMin, ptMax As Point
    Sub New()
        ConMenuEdit = New ContextMenu
    End Sub

    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '以覆寫工具列元件事件,定義工具列元件所產生的行為委任及例外處理.

        If bEdit Then DisposeEditPanel() '當使用者切換工具列元件時,所處理的行為

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
            Case 1   '選擇任意範圍
                pnWorkarea.ContextMenu = ConMenuEdit
                AddHandler pnWorkarea.MouseDown, AddressOf GetEditRangeOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf GetEditRangeOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf GetEditRangeOnMouseUp
            Case 2   '選擇矩形範圍
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
        '取得Point() 中指定的最大範圍

        SetSbpRect(ptMax, ptMin)
        ptGetRange.Add(ptNew)

        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLine(New Pen(Color.Black, 2), ptLast, ptNew)
        ptLast = ptNew
        '在這裏又產生一個與小畫家行為不符的情形,小畫家是使用目前位置的反相色,作為繪製點的
        '顏色,但在這個程式中,如果使用反相色來繪製,必定造成效能低落        
    End Sub
    Private Sub GetEditRangeOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        bTracking = False

        ptGetRange.Add(ptGetRange(0))
        Rect = SetSbpRect(ptMax, ptMin)
        If Rect.Width < 3 OrElse Rect.Height < 3 Then Return '限制最小選取範圍

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

        If Rect.Width < 3 OrElse Rect.Height < 3 Then Return '限制最小選取範圍

        Dim path As New GraphicsPath
        path.AddRectangle(Rect)

        CreateEditPanel(Rect, path, bTransParent)
        ClearSbpRect()
    End Sub

    Protected Sub CreateEditPanel(ByVal Rect As Rectangle, ByVal Clippath As GraphicsPath, _
                                ByVal bFromTransParent As Boolean)
        '建立編輯項元件,並填充因建立編輯項元件而產生需移除的部份
        Dim bmpBack As Bitmap = WorkareaChangebmp()
        Dim Bmp As New Bitmap(Rect.Width, Rect.Height)
        Dim grfx As Graphics = Graphics.FromImage(Bmp)
        grfx.SetClip(Clippath)
        grfx.TranslateClip(-Rect.X, -Rect.Y)
        grfx.DrawImage(bmpBack, CInt(-Rect.X), CInt(-Rect.Y), bmpBack.Width, bmpBack.Height)
        grfx.Dispose()   '取得裁切圖形

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "F"
            .FillColor = clrBack
            .Data.AddPath(Clippath, False)
        End With
        GraphicsData.Add(GrapMethod)      '將補捉範圍填空
        Rect = New Rectangle(Rect.X * ZoomBase, Rect.Y * ZoomBase, _
                             Rect.Width * ZoomBase, Rect.Height * ZoomBase)  ' Rect * ZoomBase

        pnEditarea = New EditPanel(Rect, Bmp, bTransParent, bFromTransParent, clrback)
        pnEditarea.Parent = pnWorkarea
        pnEditarea.ContextMenu = ConMenuEdit
        Bmp.Dispose()
        bEdit = True
        pnWorkarea.Refresh()    '建立起編輯項元件
        '這個 EditPanel 的建構元,明顯用了太多的參數.由於透明色的問題,還沒有得到解決.
        '為了提升一點效能,只好把 bFromTransParent 當作參數導入.
    End Sub
    Protected Sub DisposeEditPanel()
        '釋放編輯項元件並填入繪圖行為中

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
        '直接由 bitmap 建立編輯項元件,並非由繪圖工作區建立
        Dim Rect As New Rectangle(0, 0, bmp.Width * ZoomBase, bmp.Height * ZoomBase)
        pnEditarea = New EditPanel(Rect, bmp, bTransParent, FormTransParent, clrBack)
        pnEditarea.Parent = pnWorkarea
        pnEditarea.ContextMenu = ConMenuEdit
        bmp.Dispose()
        bEdit = True
    End Sub
End Class
