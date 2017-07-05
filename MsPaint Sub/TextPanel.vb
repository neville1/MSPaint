Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class TextPanel
    '本類別定義文字方塊控制項外觀及實作
    Inherits RectControlDot

    Private Txtbox As TextBox
    Private lblsz As Label
    Private FontSz As Size
    Private bmpBack As Bitmap
    Private szBase As Size
    Private RowSize As Integer = 0
    '文字方塊中,有相當多未完成的實作,包含外觀,防止內縮或捲動及一些特殊鍵的行為處理
    '目前只有使用者使用標準行為,才能正常工作
    '其中的最大的問題是 TextChanged 是在文字輸入完後後,才觸發的行為
    '但是需要的是在文字尚未在 Textbox 動作前,以程式碼控制

    Property SetFont() As Font
        Get
            Return Txtbox.Font
        End Get
        Set(ByVal Value As Font)
            Txtbox.Font = Value
        End Set
    End Property
    ReadOnly Property GetText() As String()
        Get
            Return Txtbox.Lines
        End Get
    End Property
    Sub New(ByRef obj As Panel, ByVal sz As Size, _
            ByVal fnt As Font, ByVal clr As Color, ByVal TransParent As Boolean)

        Parent = obj

        szBase = New Size(sz.Width * 5 + 6, sz.Height + 6)
        Me.Size = szBase

        AddHandler Me.Paint, AddressOf TextPanelOnPaint

        Txtbox = New TextBox
        With Txtbox
            .Parent = Me
            .Location = New Point(3, 3)
            .AcceptsTab = True
            .AcceptsReturn = False
            .BorderStyle = BorderStyle.None
            .Size = New Size(sz.Width * 5, sz.Height)
            .AutoSize = False
            .Multiline = True
            .Font = fnt
            .ForeColor = clr
            .WordWrap = True
        End With
    End Sub
    Private Sub TextPanelOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        '文字工具列外觀
        Dim pn As New Pen(Color.DarkBlue, 1)
        pn.DashStyle = Drawing2D.DashStyle.Dash
        Dim grfx As Graphics = pea.Graphics()

        grfx.DrawRectangle(pn, 1, 1, Me.Width - 3, Me.Height - 3)
    End Sub
    Protected Overrides Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        Dim RectEdit As Rectangle
        MyBase.ControldotOnMove(obj, mea)

        If MyBase.Tracking AndAlso _
            MyBase.Rect.X + MyBase.Rect.Width <= Me.Parent.Width AndAlso _
            MyBase.Rect.Y + MyBase.Rect.Height <= Me.Parent.Height Then

            If MyBase.Rect.Width < szBase.Width Then
                MyBase.Rect = New Rectangle(MyBase.Rect.X, MyBase.Rect.Y, _
                                            szBase.Width, MyBase.Rect.Height)
            End If
            If MyBase.Rect.Height < szBase.Height Then
                MyBase.Rect = New Rectangle(MyBase.Rect.X, MyBase.Rect.Y, _
                                            MyBase.Rect.Width, szBase.Height)
            End If

            Me.Location = New Point(MyBase.Rect.X, MyBase.Rect.Y)
            Me.Size = New Size(MyBase.Rect.Width, MyBase.Rect.Height)
            Txtbox.Size = New Size(Me.Width - 6, Me.Height - 6)
        Else
            MyBase.Rect = New Rectangle(Me.Location.X, Me.Location.Y, Me.Width, Me.Height)
        End If
        '加上邊界及最小範圍限制,並且修改與EditPanel 不同的大小變更行為
    End Sub
    Private Sub TxtboxOnKeyPress(ByVal obj As Object, ByVal kpea As KeyPressEventArgs)
        Dim grfx As Graphics = Txtbox.CreateGraphics()


        If kpea.KeyChar = ChrW(13) AndAlso _
        Me.Location.Y + Me.Height + szBase.Height < Me.Parent.Height Then
            Me.Size = New Size(Me.Width, Me.Height + szBase.Height)
            Txtbox.Size = New Size(Txtbox.Width, Txtbox.Height + szBase.Height)
        End If
    End Sub
End Class
