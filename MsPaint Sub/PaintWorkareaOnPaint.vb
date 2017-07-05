Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintWorkareaOnPaint
    '此類別定義關於繪圖工作區上,實作的繪圖行為
    Inherits PaintStatusbar

    Protected Workpath As GraphicsPath                  '繪圖工作命令
    Protected GraphicsData As ArrayList                 '繪圖動作集合
    Protected UndoCount As Integer = 0                  '復原次數

    Sub New()
        GraphicsData = New ArrayList
        Workpath = New GraphicsPath
        AddHandler pnWorkarea.Paint, AddressOf WorkareaOnPaint
    End Sub
    Protected Overridable Sub WorkareaOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '控制使用者所見的繪圖工作區顯示畫面
        pea.Graphics.DrawImage(bmppic, 0, 0, pnWorkarea.Width, pnWorkarea.Height)

        pea.Graphics.ScaleTransform(ZoomBase, ZoomBase)
        GraphircsDataAnalysis(pea.Graphics)

        pea.Graphics.ScaleTransform(1 / ZoomBase, 1 / ZoomBase)
    End Sub
    Protected Function WorkareaChangeBmp() As Bitmap
        '將繪圖工作區所顯示的影像,轉換成 Bitmap

        Dim BmpBack As Bitmap = bmpPic.Clone
        Dim grfxbmp As Graphics = Graphics.FromImage(BmpBack)

        GraphircsDataAnalysis(grfxbmp)

        grfxbmp.Dispose()
        Return BmpBack
        '將繪圖動作轉成 Bitmap 在各項繪圖命令中,大量使用此一功能
        '但是使用這個方法,在圖形大時,有效率不佳的隱憂
    End Function
    Private Sub GraphircsDataAnalysis(ByVal grfx As Graphics)
        '分析己有的使用者繪者資訊.並加以實行.

        Dim GrapData() As GraphicsItem = CType(GraphicsData.ToArray(GetType(GraphicsItem)), GraphicsItem())
        Dim i, ShowCount As Integer

        ShowCount = GrapData.GetUpperBound(0) - UndoCount      '扣除復原行為後,擁有的繪圖動作總數.
        For i = 0 To ShowCount
            Select Case GrapData(i).DataType
                Case "D"          '依點繪圖
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
                Case "I"         '依點貼圖
                    Dim apt() As PointF = GrapData(i).Data.PathPoints
                    Dim Count As Integer
                    Dim img As Image = GrapData(i).Image
                    For Count = 0 To apt.GetUpperBound(0)
                        grfx.DrawImage(img, apt(Count))
                    Next Count
                Case "P"         '依路徑繪圖  
                    Dim pn As New Pen(GrapData(i).Color, GrapData(i).pWidth)
                    grfx.DrawPath(pn, GrapData(i).Data)
                Case "F"         '依路徑填充
                    Dim br As Brush = New SolidBrush(GrapData(i).FillColor)
                    grfx.FillPath(br, GrapData(i).Data)
                Case "R"         '依路徑填充+繪圖
                    Dim br As Brush = New SolidBrush(GrapData(i).FillColor)
                    grfx.FillPath(br, GrapData(i).Data)
                    Dim pn As New Pen(GrapData(i).Color, GrapData(i).pWidth)
                    grfx.DrawPath(pn, GrapData(i).Data)
                Case "B"         '依影像貼圖                     
                    grfx.DrawImage(GrapData(i).Image, _
                    New Rectangle(GrapData(i).ImageLocation.X, _
                    GrapData(i).ImageLocation.Y, _
                    GrapData(i).ImageSize.Width, _
                    GrapData(i).ImageSize.Height))
                    
            End Select
        Next i
        '分析 GraphicsItem 的資料,並予以實行
        '這應該是本程式中最重要的部份.
        '但很明顯的,小弟在這裏用了太多的 Type,有易造成混亂之虞
    End Sub

End Class
