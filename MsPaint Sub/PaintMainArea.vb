Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintMainArea
    Inherits Form
    '本類別定義關於繪圖工作區外觀及狀態變更處理方式

    Protected strProgName As String = "小畫家重製版"

    Protected bmpPic As Bitmap                '底層圖形
    ' Protected grfxFore As Graphics            '前景畫布(臨時性)
    Protected szBmp As New Size(512, 384)     '圖形大小(給值是為初期建構)
    Protected bToolbar, bColorbar, bStatusbar, bTextFrom As Boolean   '配置開關
    Protected pnWorkarea As DoubleBufferPanel    '繪圖工作區
    Protected ZoomBase As Integer = 1            '繪圖區縮放比  

    Private pnAll As Panel                     '底層
    Private DotTL, DotTM, DotTR As ControlDot    '工作區大小變更控制點
    Private DotML, DotMR As ControlDot
    Private DotDL, DotDM, DotDR As ControlDot

    Sub New()
        '繪圖工作區建構元
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

        CreateControlDot()    '建立工作區四周控制大小用的控制點
        AddHandler pnWorkarea.Resize, AddressOf WorkareaOnReSize
        AddHandler Me.SizeChanged, AddressOf FormSizeChanged   
        
    End Sub
    Protected Overridable Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '定義因視覺元件變更(工具列,色塊區)繪圖工作區座標配置
        Dim tx, cy, sy As Integer
        If bToolbar Then tx = 60 Else tx = 0
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0
        pnAll.Size = New Size(ClientSize.Width - tx, ClientSize.Height - cy - sy)
        pnAll.Location = New Point(tx, 0)
    End Sub
    Protected Sub WorkareaOnReSize(ByVal obj As Object, ByVal ea As EventArgs)
        '繪圖工作區大小變更後,外圍控制點座標配置
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
        '建立繪圖工作區外圍控制點,以調整繪圖工作區大小
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
        '這裏的程式碼,雖說的工作區的配置,但是主要是在定義包圍工作區的控制點及其動作.
        '這種作法感覺上很笨拙.而且與RectControlDot.vb 有許多類似之處.
        '曾試著把這些控制點物件化,但是帶來的卻是許多的資料交換.
        '應該會有更好的作法才是.

        '這裏的程式碼,依照基本原則,應該是被放在工作區建構元內
        '此處單純因為不希望工作區建構元看來太過繁雜,故將其移出建構元
    End Sub
    Protected Overridable Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '定義當使用者在外圍控制點上作拖曳行為時,所做的處理
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

        '圖形縮放比變更後的作法未加入
    End Sub
    Protected Overridable Sub WorkAreaSizeChanged(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '當繪圖工作區大小受到變更要求時.
        '由 PaintWithMenuImage.vb 覆寫,控制工作區像素變更的裁切或加大.
    End Sub
    Protected Sub ZoomBaseChange()
        '變更工作區畫面縮放比
        Dim Sz As New Size(szBmp.Width * ZoomBase, szBmp.Height * ZoomBase)
        pnall.AutoScroll = False                '避免因 Scroll 位移,導致座標誤差.
        pnWorkarea.Size = Sz
        'grfxFore = pnWorkarea.CreateGraphics()
        'grfxFore.ScaleTransform(ZoomBase, ZoomBase)

        pnWorkarea.Refresh()                    '觸發 On_Paint,反映縮放比變更.          
        pnall.AutoScroll = True
    End Sub
    Protected Function GetGraphicsbyFore() As Graphics
        Dim grfx As Graphics = pnWorkarea.CreateGraphics()
        grfx.ScaleTransform(ZoomBase, ZoomBase)
        Return grfx
    End Function
End Class
