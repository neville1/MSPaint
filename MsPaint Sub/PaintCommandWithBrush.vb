Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms

Public Class PaintCommandWithBrush
    '本類別實作工具列中關於筆刷相似類的處理.
    '此類別因其行為相近,不重複說明其相似行為.
    Inherits PaintWorkareaOnPaint

    Private ptLast, ptNew As Point
    Private ptArea As Point
    Private clrWork As Color
    Private bmpBrush As Bitmap
    Private bmpOffest As Integer
    Private tmr As Timer
    Private arrlstPoint As ArrayList

    Private bTracking As Boolean = False        '滑鼠拖曳

    Sub New()
        '定義程式啟動時,預設的工具列選項.
        arrlstPoint = New ArrayList

        AddHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
        AddHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
        AddHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
    End Sub
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '以覆寫工具列元件事件,定義工具列元件所產生的行為委任及例外處理.

        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 3
                RemoveHandler pnWorkarea.MouseDown, AddressOf EraseOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf EraseOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf EraseOnMouseUp
                RemoveHandler pnWorkarea.MouseLeave, AddressOf EraseOnMouseLeave
            Case 7
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
            Case 8
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawBrushOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawBrushOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawBrushOnMouseUp
                RemoveHandler pnWorkarea.MouseLeave, AddressOf DrawBrushOnMouseLeave
            Case 9
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawSprayOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawSprayOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawSprayOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 3 '橡皮擦
                AddHandler pnWorkarea.MouseDown, AddressOf EraseOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf EraseOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf EraseOnMouseUp
                AddHandler pnWorkarea.MouseLeave, AddressOf EraseOnMouseLeave
            Case 7 '畫筆
                AddHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
            Case 8 '筆刷 
                AddHandler pnWorkarea.MouseDown, AddressOf DrawBrushOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawBrushOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawBrushOnMouseUp
                AddHandler pnWorkarea.MouseLeave, AddressOf DrawBrushOnMouseLeave
            Case 9 '噴鎗
                AddHandler pnWorkarea.MouseDown, AddressOf DrawSprayOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawSprayOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawSprayOnMouseUp
        End Select
        '改寫數次後,小弟最後決定用這個方式,來定義工具列功能的委任
        '雖然繪圖行為中,有許多相似的方法,用這種方式來定義,較不容易來共用相似的程式碼
        '也易造成程式碼的肥大,但目前考慮的重點是,儘量讓流程清楚,易於閱讀
    End Sub
    Private Sub EraseOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()          '清除因復原,重複功能所產生的多餘命令

        bTracking = True
        ptArea = New Point(mea.X, mea.Y)
        ptLast = GetScalePoint(ptArea)

        arrlstPoint.Add(ptLast)
        '橡皮擦功能,本來以為這樣的功能,只需要利用畫筆大小的改變及使用底色即可.
        '後來發現,如果用這樣的方式製作,會產生不少的例外情形,也不能完全正確的表示使用者的動作
        '這裏使用一連串連續的矩形,來表示橡皮擦行為的動作
    End Sub
    Private Sub EraseOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()


        Dim UpdateArea As New Rectangle(ptArea.X - 2, ptArea.Y - 2, _
                                        EraseSize * ZoomBase + 4, EraseSize * ZoomBase + 4)
        pnWorkarea.Invalidate(UpdateArea)
        pnWorkarea.Update()
        ptArea = New Point(mea.X, mea.Y)
        ptNew = GetScalePoint(ptArea)
        Dim grfxCursor As Graphics = pnWorkarea.CreateGraphics()
        Dim RangeCuror As Integer = EraseSize * ZoomBase
        grfxCursor.FillRectangle(New SolidBrush(clrBack), ptArea.X, ptArea.Y, _
                                 EraseSize * ZoomBase, EraseSize * ZoomBase)
        grfxCursor.DrawRectangle(New Pen(Color.Black), ptArea.X, ptArea.Y, _
                                 EraseSize * ZoomBase, EraseSize * ZoomBase)
        grfxCursor.Dispose()
        '繪製臨時遊標
        '這點實在有點微詞,目前 .Net 的 Cursor,只能從 *.cur 檔中建立,在執行期中,僅有改變
        'Cursor 大小的方法可供使用.當然這不是代表,只能用這種浪費效能,製造閃爍的方法來建
        '立遊標.    
        If Not bTracking Then Return
    
        arrlstPoint.Add(ptNew)

        InsertCountinePoint()                  '插入點使其連續
        Dim br As Brush = New SolidBrush(clrBack)
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        Dim Rect() As Rectangle = PointToRect(apt, EraseSize)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.FillRectangles(br, Rect)
        grfxFore.Dispose()
        ptLast = ptNew

        '這裏用 .Invalidate 取代 .Refush 是希望能夠減少閃爍.
    End Sub
    Private Sub EraseOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        bTracking = False

        '使用者行為確定,存入繪圖行為集合
        Dim GrapMethod As New GraphicsItem
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        If apt.Length = 1 Then
            With GrapMethod
                .DataType = "D"
                .Color = ClrBack
                .pWidth = EraseSize
                .Data.AddLines(apt)
            End With
        Else
            Dim Rect() As Rectangle = PointToRect(apt, EraseSize - 1)
            With GrapMethod
                .DataType = "P"
                .Color = ClrBack
                .pWidth = 1
                .Data.AddRectangles(Rect)
            End With
        End If

        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        pnWorkarea.Refresh()
        '此類產生了一個例外現象
        '因為在Graphics.DrawLine 的定義中 (0,0)->(0,1) 應該是包含起點及終點的.
        '但是在 (0,0)->(0,0) 這種情況,卻視為無效行為
        '故要特別處理,只有一點的狀況
    End Sub
    Private Sub EraseOnMouseLeave(ByVal obj As Object, ByVal ea As EventArgs)
        pnWorkarea.Refresh()   '配合建立的臨時遊標以清除臨時遊標
    End Sub

    Private Sub DrawPointOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        arrlstPoint.Add(ptLast)
    End Sub
    Private Sub DrawPointOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return
        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        arrlstPoint.Add(ptNew)

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim pn As New Pen(clrWork)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLines(pn, apt)
        grfxFore.Dispose()

        ptLast = ptNew
        '繪圖動作的最基本型態,如果有人利用這個程式來參考繪圖作法(可能嗎?程式技巧實在不怎樣)
        '這裏是最理想的切入點
    End Sub
    Private Sub DrawPointOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        bTracking = False
      
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        If apt.Length = 1 Then
            GrapMethod.DataType = "D"
        Else
            GrapMethod.DataType = "P"
        End If
        With GrapMethod
            .Color = clrWork
            .pWidth = 1                  '畫筆特性,寬度不變
            .Data.AddLines(apt)
            .Data.StartFigure()
        End With

        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        pnWorkarea.Refresh()
        '同 EraseOnMouseUp 原因
    End Sub

    Private Sub DrawBrushOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptArea = New Point(mea.X - bmpOffest * ZoomBase, mea.Y - bmpOffest * ZoomBase)
        ptLast = GetScalePoint(ptArea)
        arrlstPoint.Add(ptLast)
        '這裏使用一個 Point() 和一個代表 Brush 的 bitmap,來表示筆刷的繪圖行為
        '可以想見,這種表示方法,帶來的一定是效能及記憶體的不良影響.
        '不過在 .Net 筆刷型態如此豐富時,還找不到一個適用於此項功能的方式.
        '推想小畫家的此項功能,應是配合當時的方法而製作的.
    End Sub
    Private Sub DrawBrushOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        SetBmpBrush(mea)        '設定使用的Brush(以含有透明色的Bitmap建立)
        
        Dim UpdateArea As New Rectangle(ptArea.X, ptArea.Y, bmpBrush.Width * ZoomBase + 1, bmpBrush.Height * ZoomBase + 1)
        pnWorkarea.Invalidate(UpdateArea)
        pnWorkarea.Update()
        ptArea = New Point(mea.X - bmpOffest * ZoomBase, mea.Y - bmpOffest * ZoomBase)
        ptNew = GetScalePoint(ptArea)

        Dim grfxCursor As Graphics = pnWorkarea.CreateGraphics()
        grfxCursor.DrawImage(bmpBrush, ptArea.X, ptArea.Y, bmpBrush.Width * ZoomBase, bmpBrush.Height * ZoomBase)
        grfxCursor.Dispose()
        '繪製遊標,理由同 DrawEraseOnMouseMove

        If Not bTracking Then Return

        InsertCountinePoint()                '在兩點間插入點,使其呈連續狀態

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        Dim i As Integer  
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        For i = 0 To apt.GetUpperBound(0)
            grfxFore.DrawImage(bmpBrush, apt(i).X, apt(i).Y, bmpBrush.Width, bmpBrush.Height)
        Next i
        grfxFore.Dispose()

        ptLast = ptNew
    End Sub
    Private Sub DrawBrushOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        bTracking = False

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "I"
            .Image = bmpBrush
            .Data.AddLines(apt)
        End With
        GraphicsData.Add(GrapMethod)

        arrlstPoint.Clear()
    End Sub
    Private Sub DrawBrushOnMouseLeave(ByVal obj As Object, ByVal ea As EventArgs)
        pnWorkarea.Refresh()
    End Sub
    Private Sub SetBmpBrush(ByVal mea As MouseEventArgs)
        '設定筆刷型態

        Select Case (BrushType - 1) Mod 3
            Case 0
                bmpBrush = New Bitmap(8, 8)
                bmpOffest = 4
            Case 1
                bmpBrush = New Bitmap(5, 5)
                bmpOffest = 2
            Case 2
                bmpBrush = New Bitmap(2, 2)
                bmpOffest = 1
        End Select

        If mea.Button = MouseButtons.Right Then
            clrWork = clrBack
        Else
            clrWork = clrFore
        End If

        Dim grfx As Graphics = Graphics.FromImage(bmpBrush)
        Dim ClrNot As Color = Color.FromArgb(clrWork.A, Not clrWork.R, Not clrWork.G, Not clrWork.B)
        grfx.Clear(ClrNot)   '用反向色作為背景

        Select Case BrushType
            Case 1
                MyDrawEllipse(grfx, New Point(3, 3), 3, clrWork)
            Case 2
                MyDrawEllipse(grfx, New Point(2, 2), 2, clrWork)
            Case 3
                MyDrawEllipse(grfx, New Point(1, 1), 1, clrWork)
            Case 4
                grfx.FillRectangle(New SolidBrush(clrWork), 0, 0, 7, 7)
            Case 5
                grfx.FillRectangle(New SolidBrush(clrWork), 0, 0, 4, 4)
            Case 6
                grfx.DrawRectangle(New Pen(clrWork), 0, 0, 1, 1)
            Case 7
                grfx.DrawLine(New Pen(clrWork, 2), 7, 0, 0, 7)
            Case 8
                grfx.DrawLine(New Pen(clrWork, 2), 4, 0, 0, 4)
            Case 9
                grfx.DrawLine(New Pen(clrWork, 2), 1, 0, 0, 1)
            Case 10
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 7, 7)
            Case 11
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 4, 4)
            Case 12
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 1, 1)
        End Select
        grfx.Dispose()
        bmpBrush.MakeTransparent(ClrNot)       '筆刷背景透明
        '這裏本來想從 PaintToolbarSub 中,把 bitmap 傳遞過來即可,畢竟這段程式碼看來.
        '實在有些愚笨和礙眼,不過為了資料傳遞上的一致性而言,只好暫時複製一份
        '寫在這裏作為更動程式結構時的考慮
    End Sub

    Private Sub DrawSprayOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))

        tmr = New Timer
        tmr.Interval = 20
        tmr.Enabled = True
        AddHandler tmr.Tick, AddressOf DrawSpray
        '以一個 Timer 物件來建立噴鎗效果
    End Sub
    Private Sub DrawSprayOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub DrawSprayOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return

        bTracking = False
        tmr.Dispose()    
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "D"
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddLines(apt)
            .Data.StartFigure()
        End With
        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        '這裏以一個 Point() 表示噴鎗行為
    End Sub
    Private Sub DrawSpray(ByVal obj As Object, ByVal ea As EventArgs)
        Dim RndAngle As Single
        Dim RndLength As Integer
        Dim RndPoint As Point
        Dim i As Integer

        Dim br As Brush = New SolidBrush(clrWork)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        For i = 1 To 2 * SprayWidth
            RndAngle = (2 * Math.PI * Rnd())
            RndLength = CInt((SprayWidth * Rnd()))
            RndPoint = New Point(ptLast.X + RndLength * Math.Cos(RndAngle), _
                                 ptLast.Y + RndLength * Math.Sin(RndAngle))
            Dim Rect As New Rectangle(RndPoint.X, RndPoint.Y, 1, 1)
            grfxFore.FillRectangle(br, Rect)
            arrlstPoint.Add(RndPoint)
        Next i
        grfxFore.Dispose()
        '亂數取點,不知道與小畫家的方法是否一致,但執行效果看來十分接近.
    End Sub

    Protected Function SelectUserColor(ByVal mea As MouseEventArgs) As Color
        '依使用者按鍵選擇顏色
        Select Case mea.Button
            Case MouseButtons.Left
                Return clrFore
            Case MouseButtons.Right
                Return ClrBack
            Case Else
                Return clrFore
        End Select
    End Function
    Protected Sub ClearCommandRecord()
        '清除由復原/重覆所造成的多餘繪圖命令
        If UndoCount > 0 Then
            GraphicsData.RemoveRange(GraphicsData.Count - UndoCount, UndoCount)
            UndoCount = 0
        End If
    End Sub
    Protected Sub SetMousePosition()
        '顯示滑鼠座標
        sbpLocation.Text = pnWorkarea.PointToClient(MousePosition).X.ToString & "," & _
        pnWorkarea.PointToClient(MousePosition).Y.ToString
    End Sub
    Private Sub InsertCountinePoint()
        '在二個 Point 中,插入不定數的 Point,使其成為一個由連續點所組成的 Point()
        If ptLast.Equals(ptNew) Then Return

        Dim ShiftPoint As Point = New Point(Math.Abs(ptLast.X - ptNew.X), Math.Abs(ptLast.Y - ptNew.Y))
        Dim OffestX, OffestY As Integer
        Dim SingleX, SingleY As Integer
        Dim InsPoint As Point

        If ptLast.X > ptNew.X Then SingleX = -1 Else SingleX = 1
        If ptLast.Y > ptNew.Y Then SingleY = -1 Else SingleY = 1

        If ShiftPoint.X > ShiftPoint.Y Then
            For OffestX = 1 To ShiftPoint.X Step 1
                OffestY = Math.Round((ShiftPoint.Y / ShiftPoint.X) * OffestX)
                InsPoint = New Point(ptLast.X + (OffestX * SingleX), ptLast.Y + (OffestY * SingleY))
                arrlstPoint.Add(InsPoint)
            Next OffestX
        Else
            For OffestY = 1 To ShiftPoint.Y Step 1
                OffestX = Math.Round((ShiftPoint.X / ShiftPoint.Y) * OffestY) * SingleX
                InsPoint = New Point(ptLast.X + OffestX, ptLast.Y + (OffestY * SingleY))
                arrlstPoint.Add(InsPoint)
            Next OffestY
        End If
        '根據直線方程式斜率計算
    End Sub    
    Private Function PointToRect(ByVal apt() As Point, ByVal Range As Integer) As Rectangle()
        Dim Rect(apt.GetUpperBound(0)) As Rectangle

        Dim i As Integer
        For i = 0 To apt.GetUpperBound(0)
            Rect(i) = New Rectangle(apt(i).X, apt(i).Y, Range, Range)
        Next i
        Return Rect
        'Point() 轉成 Rectangle()
    End Function
    Protected Function GetScalePoint(ByVal pt As Point) As Point
        Return New Point(pt.X / ZoomBase, pt.Y / ZoomBase)
    End Function


End Class
