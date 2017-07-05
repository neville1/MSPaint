Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class FullScreenShowImage
    '本類別定義關於功能表中檢視點陣圖,所產生的外觀行為畫面
    Inherits Form
    Private WorkBmp As Bitmap

    Sub New(ByVal bmp As Bitmap)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Location = New Point(0, 0)
        Me.ControlBox = False
        Me.WindowState = FormWindowState.Maximized

        AddHandler Me.KeyDown, AddressOf UserCloseFormOnKeyDown
        AddHandler Me.Click, AddressOf UserCloseFormOnClick
        AddHandler Me.Paint, AddressOf FormOnPaint
        WorkBmp = bmp
    End Sub
    Private Sub UserCloseFormOnKeyDown(ByVal obj As Object, ByVal kea As KeyEventArgs)
        '使用者按下任何按鍵
        Me.Close()
    End Sub
    Private Sub UserCloseFormOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '使用者點擊滑鼠
        Me.Close()
    End Sub
    Private Sub FormOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        pea.Graphics.Clear(Color.DarkBlue)

        Dim Rect As Rectangle
        If WorkBmp.Width > Me.Width OrElse WorkBmp.Height > Me.Height Then
            Dim szf As New SizeF(WorkBmp.Width / WorkBmp.HorizontalResolution, _
                     WorkBmp.Height / WorkBmp.VerticalResolution)
            Dim fScale As Single = Math.Min(Me.Width / szf.Width, Me.Height / szf.Height)
            szf.Width *= fScale
            szf.Height *= fScale
            Rect = New Rectangle((Me.Width - szf.Width) / 2, (Me.Height - szf.Height) / 2, _
                                               szf.Width, szf.Height)
        Else
            Rect = New Rectangle((Me.Width - WorkBmp.Width) \ 2, (Me.Height - WorkBmp.Height) \ 2 _
                                 , WorkBmp.Width, WorkBmp.Height)
        End If

        pea.Graphics.DrawImage(WorkBmp, Rect)
        '這個全螢幕檢視的功能,與小畫家不同.小畫家的作法是,將目前使用者所見到的工作區範圍
        '以全螢幕的方式顯示在畫面上.而小弟是將全圖縮放至螢幕可以顯示的範圍內.        
    End Sub
End Class
