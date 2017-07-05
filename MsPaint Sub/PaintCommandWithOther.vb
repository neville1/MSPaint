Imports System
Imports System.Drawing
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintCommandWithOther
    '本類別實作工具列中尚未處理的部份(填色,取色,放大鏡)
    Inherits PaintCommandWithText

    Private clrWork As Color

    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 4
                RemoveHandler pnWorkarea.MouseDown, AddressOf FillColorOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf FillColorOnMouseMove
            Case 5
                RemoveHandler pnWorkarea.MouseDown, AddressOf GetPointColorOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf GetPointColorOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf GetPointColorOnMouseUp
            Case 6
                RemoveHandler pnWorkarea.MouseMove, AddressOf ZoomChangeOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf ZoomChangeOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 4   '填充顏色
                AddHandler pnWorkarea.MouseDown, AddressOf FillColorOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf FillColorOnMouseMove
            Case 5   '取得顏色
                AddHandler pnWorkarea.MouseDown, AddressOf GetPointColorOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf GetPointColorOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf GetPointColorOnMouseUp
            Case 6    '放大鏡
                AddHandler pnWorkarea.MouseMove, AddressOf ZoomChangeOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf ZoomChangeOnMouseUp
        End Select
    End Sub
    Private Sub GetPointColorOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '根據繪圖情形取色

        Dim bmpback As Bitmap = WorkareaChangeBmp()

        Dim GetPoint As Point = GetScalePoint(New Point(mea.X, mea.Y))
        Select Case mea.Button
            Case MouseButtons.Left
                clrFore = bmpback.GetPixel(GetPoint.X, GetPoint.Y)
                btnColorFore.BackColor = clrFore
            Case MouseButtons.Right
                clrBack = bmpback.GetPixel(GetPoint.X, GetPoint.Y)
                btnColorBack.BackColor = clrBack
        End Select

        bmpback.Dispose()
    End Sub
    Private Sub GetPointColorOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub GetPointColorOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        radbtnToolbar(LastCommand - 1).Checked = True
        Dim ea As EventArgs
        ToolbarRadbtnSelect(radbtnToolbar(LastCommand - 1), ea.Empty)  '選擇上一個命令
    End Sub

    Private Sub FillColorOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '填色功能實作
        Dim bmpback As Bitmap = WorkareaChangeBmp()

        Dim ptPo As Point = GetScalePoint(New Point(mea.X, mea.Y))
        Dim pt, ptSum As Point
        Dim Width, Height As Integer
        Width = bmpback.Width
        Height = bmpback.Height

        clrWork = SelectUserColor(mea)

        Dim clrFill As Color = bmpback.GetPixel(ptPo.X, ptPo.Y)
        If clrWork.ToArgb = clrFill.ToArgb Then Return
        Dim Points As New Stack

        Dim ptDirEction() As Point = {New Point(0, -1), New Point(1, 0), _
                                      New Point(0, 1), New Point(-1, 0)} '定義方向
        Dim i As Integer
        Points.Push(ptPo)
        Do
            pt = Points.Pop()
            bmpback.SetPixel(pt.X, pt.Y, clrWork)

            For i = 0 To 3
                ptSum = New Point(pt.X + ptDirEction(i).X, pt.Y + ptDirEction(i).Y)
                If ptSum.X >= 0 AndAlso ptSum.X < Width AndAlso _
                   ptSum.Y >= 0 AndAlso ptSum.Y < Height Then
                    If bmpback.GetPixel(ptSum.X, ptSum.Y).ToArgb = clrFill.ToArgb Then
                        Points.Push(ptSum)
                    End If
                End If
            Next i
        Loop While Points.Count > 0

        WorkareaAddGraphicsItem(bmpback)
        '這段程式碼,簡直是集合所有程式碼的壞處了.
        '速度極慢,記憶體損失大.自從寫這個程式以來,這裏一直是自己考慮的重點.
        '因為 ExtFloodFill 這個API,並沒有被包含在 .Net 中.
        '所以一直落入交戰中,到底要不要呼叫 API來達成這個功能.
        '後來試圖利用 API 來製作時發現,要完成這個功能.至少要使用 CreateSolidBrush 
        'CreateCompatibleDC,CreateCompatibleBitmap,SelectObject,DeleteObject,及
        'ExtFloodFill.雖然可以完成,速度上也相當不錯,但這樣已經不是在學習 VB.Net.
        '不但把與裝置無關的特性破壞怠盡,也讓程式中 grax = Graphics.FormImage() 
        '這樣的一致性消失.所以先以這樣的方法來製作這個功能.
    End Sub
    Private Sub FillColorOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub ZoomChangeOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub ZoomChangeOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
    End Sub
    Protected Overridable Sub WorkareaAddGraphicsItem(ByVal bmp As Bitmap)
        '由 PaintWithMenuImage 實作,將全影影存入繪圖行為
    End Sub
End Class
