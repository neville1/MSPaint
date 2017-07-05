Imports System
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports System.Windows.Forms
Public Class PropertyDialog
    '本類別定義關於圖形屬性對話方塊外觀及動作
    Inherits Form

    Const strTitle As String = "小畫家重製版"

    Private BmpUnit As Integer = 3
    Private BmpColor As Integer = 2
    Private Dpi As Point
    Private SzBmp As SizeF
    Private txtboxBmpSize(1) As TextBox
    Private oldSize As Size
    Private NewSize As Size

    ReadOnly Property BmpSize() As Size
        Get
            Return NewSize
        End Get
    End Property

    Sub New(ByVal strFileName As String, ByVal Sz As Size)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        Size = New Size(466, 332)
        AutoScaleBaseSize = New Size(6, 18)
        Text = "屬性"

        Dim strFileInfoStatus(1) As String
        Dim nfi As New NumberFormatInfo
        nfi.NumberDecimalDigits = 0

        Dim fs As FileStream
        If File.Exists(strFileName) Then
            strFileInfoStatus(0) = File.GetLastWriteTime(strFileName).ToString
            fs = File.Open(strFileName, FileMode.Open)
            strFileInfoStatus(1) = CInt(fs.Length).ToString("N", nfi) & " 位元組"
            fs.Close()
        Else
            strFileInfoStatus(0) = "無法使用"
            strFileInfoStatus(1) = "無法使用"
        End If

        Dim lblFileInfo(1), lblFileInfoStatus(1) As Label
        Dim strFileInfo() As String = {"上次存檔時間:", "檔案大小:"}
        Dim i As Integer
        For i = 0 To 1
            lblFileInfo(i) = New Label
            With lblFileInfo(i)
                .Parent = Me
                .Text = strFileInfo(i)
                .Location = New Point(24, 16 + i * 16)
                .Size = New Size(104, 16)
            End With

            lblFileInfoStatus(i) = New Label
            With lblFileInfoStatus(i)
                .Parent = Me
                .Text = strFileInfoStatus(i)
                .Location = New Point(152, 16 + i * 16)
                .Size = New Size(160, 16)
            End With
        Next i
        Dim lblBmpSize(1) As Label
        Dim strBmpSize() As String = {"寬度(&W):", "高度(&H):"}

        For i = 0 To 1
            lblBmpSize(i) = New Label
            With lblBmpSize(i)
                .Parent = Me
                .Text = strBmpSize(i)
                .Location = New Point(24 + i * 160, 64)
                .Size = New Size(64, 16)
            End With

            txtboxBmpSize(i) = New TextBox
            With txtboxBmpSize(i)
                .Parent = Me
                Select Case i
                    Case 0
                        .Text = Sz.Width.ToString
                    Case 1
                        .Text = Sz.Height.ToString
                End Select
                .Location = New Point(96 + i * 152, 56)
                .Size = New Size(64, 25)
                .MaxLength = 4
            End With
            AddHandler txtboxBmpSize(i).KeyUp, AddressOf BmpSizeKeyUp
            AddHandler txtboxBmpSize(i).KeyDown, AddressOf BmpSizeKeyDown
        Next i

        Dim grpBmpUnit As New GroupBox
        With grpBmpUnit
            .Parent = Me
            .Text = "單位"
            .Location = New Point(16, 88)
            .Size = New Size(304, 48)
        End With

        Dim strUnit() As String = {"英吋(&I)", "公分(&M)", "像素(&P)"}
        Dim radbtnUnit(2) As RadioButton
        For i = 0 To 2
            radbtnUnit(i) = New RadioButton
            With radbtnUnit(i)
                .Parent = grpBmpUnit
                .Name = (i + 1).ToString
                .Text = strUnit(i)
                .Location = New Point(16 + i * 96, 16)
                .Size = New Size(88, 24)
                If BmpUnit - 1 = i Then .Checked = True
            End With
            AddHandler radbtnUnit(i).CheckedChanged, AddressOf BmpUnitChanged
        Next i

        Dim grpBmpColor As New GroupBox
        With grpBmpColor
            .Parent = Me
            .Text = "色彩"
            .Location = New Point(16, 144)
            .Size = New Size(304, 48)
        End With

        Dim strColorText() As String = {"黑白(&B)", "彩色(&L)"}
        Dim radbtnBmpColor(1) As RadioButton
        For i = 0 To 1
            radbtnBmpColor(i) = New RadioButton
            With radbtnBmpColor(i)
                .Parent = grpBmpColor
                .Text = strColorText(i)
                .Location = New Point(16 + i * 192, 16)
                .Size = New Size(80, 24)
                If BmpColor - 1 = i Then .Checked = True
            End With
        Next i

        Dim grpBmpTransparent As New GroupBox
        With grpBmpTransparent
            .Parent = Me
            .Text = "透明效果"
            .Location = New Point(16, 200)
            .Size = New Size(304, 80)
            .Enabled = False
        End With

        Dim chkTransparent As New CheckBox
        With chkTransparent
            .Parent = grpBmpTransparent
            .Text = "使用透明背景色彩(&T)"
            .Location = New Point(16, 16)
            .Size = New Size(176, 24)
        End With

        Dim btnSelectTransparent As New Button
        With btnSelectTransparent
            .Parent = grpBmpTransparent
            .Text = "選取色彩(&C)"
            .Location = New Point(56, 40)
            .Size = New Size(144, 32)
        End With

        Dim palTransparentColor As New Panel
        With palTransparentColor
            .Parent = grpBmpTransparent
            .Location = New Point(224, 40)
            .Size = New Size(32, 32)
            .BorderStyle = BorderStyle.Fixed3D
        End With

        Dim btnOK As New Button
        With btnOK
            .Parent = Me
            .Text = "確定"
            .Location = New Point(352, 16)
            .Size = New Size(96, 32)
        End With
        AddHandler btnOK.Click, AddressOf ButtonOKOnClick

        Dim btnCancel As New Button
        With btnCancel
            .Parent = Me
            .Text = "取消"
            .Location = New Point(352, 56)
            .Size = New Size(96, 32)
        End With
        AddHandler btnCancel.Click, AddressOf ButtonCancelOnClick

        Dim btnDefault As New Button
        With btnDefault
            .Parent = Me
            .Text = "預設(&D)"
            .Location = New Point(352, 96)
            .Size = New Size(96, 32)
        End With
        AddHandler btnDefault.Click, AddressOf ButtonDefaultOnClick

        SzBmp = New SizeF(Sz.Width, Sz.Height)
        oldSize = Sz
        Dim grfx As Graphics = Me.CreateGraphics()
        Dpi = New Point(grfx.DpiX, grfx.DpiY)
    End Sub
    Private Sub BmpUnitChanged(ByVal obj As Object, ByVal ea As EventArgs)
        Dim radbtn As RadioButton = DirectCast(obj, RadioButton)
        BmpUnit = CInt(radbtn.Name)
        UpdateText()
    End Sub
    Private Sub BmpSizeKeyDown(ByVal obj As Object, ByVal kea As KeyEventArgs)
        If kea.KeyData = Keys.Enter Then
            Dim ea As EventArgs
            ButtonOKOnClick(obj, ea.Empty)
        End If
    End Sub
    Private Sub BmpSizeKeyUp(ByVal obj As Object, ByVal kea As KeyEventArgs)
        Try
            SzBmp.Width = Val(txtboxBmpSize(0).Text)
            SzBmp.Height = Val(txtboxBmpSize(1).Text)

        Catch ex As Exception
            MessageBox.Show("請輸入一數字.", strTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End Try

    End Sub
    Private Sub ButtonOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If SzBmp.Width < 1 OrElse SzBmp.Width > 3000 OrElse _
           SzBmp.Height < 1 OrElse SzBmp.Height > 3000 Then
            MessageBox.Show("圖形一邊的像素必需在1-3000之間", strTitle, _
                             MessageBoxButtons.OK, MessageBoxIcon.Information)
            NewSize = oldSize
        Else
            NewSize = New Size(SzBmp.Width, SzBmp.Height)
        End If
        DialogResult = DialogResult.OK
    End Sub
    Private Sub ButtonCancelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.Cancel
    End Sub
    Private Sub ButtonDefaultOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        SzBmp = New SizeF(512, 384)
        UpdateText()
    End Sub
    Private Sub UpdateText()
        Select Case CInt(BmpUnit)
            Case 1
                txtboxBmpSize(0).Text = (SzBmp.Width / Dpi.X).ToString("N")
                txtboxBmpSize(1).Text = (SzBmp.Height / Dpi.Y).ToString("N")
            Case 2
                txtboxBmpSize(0).Text = (SzBmp.Width / Dpi.X * 2.54).ToString("N")
                txtboxBmpSize(1).Text = (SzBmp.Height / Dpi.X * 2.54).ToString("N")
            Case 3
                txtboxBmpSize(0).Text = SzBmp.Width.ToString
                txtboxBmpSize(1).Text = SzBmp.Height.ToString
        End Select
        '這裏有個令人難以理解的問題,在公分的部份.小畫家在 512,384 的Pixel 時,顯示為
        '16,12.等於每平方公分有 32*32 的 Pixel. 為什麼?
    End Sub
End Class
