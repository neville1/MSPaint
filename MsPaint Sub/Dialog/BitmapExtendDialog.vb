Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class BitmapExtendDialog
    '本類定義關於延伸扭曲對話方塊外觀及動作
    Inherits Form

    Private txtboxExtend(1), txtboxDistort(1) As TextBox
    Private ExtX, ExtY, DisX, DisY As Integer
    Private bExt, bDis As Boolean
    ReadOnly Property IsExtend() As Boolean
        Get
            Return bExt
        End Get
    End Property
    ReadOnly Property IsDistort() As Boolean
        Get
            Return bDis
        End Get
    End Property
    ReadOnly Property ExtendX() As Integer
        Get
            Return ExtX
        End Get
    End Property
    ReadOnly Property ExtendY() As Integer
        Get
            Return ExtY
        End Get
    End Property
    ReadOnly Property DistortX() As Integer
        Get
            Return DisX
        End Get
    End Property
    ReadOnly Property DistortY() As Integer
        Get
            Return DisY
        End Get
    End Property

    Sub New()
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        Size = New Size(464, 318)
        Text = "延伸及扭曲"

        Dim grpMain(1) As GroupBox
        Dim i As Integer
        Dim strGrpText() As String = {"延伸", "扭曲"}
        For i = 0 To 1
            grpMain(i) = New GroupBox
            With grpMain(i)
                .Parent = Me
                .Text = strGrpText(i)
                .Location = New Point(16, 16 + i * 128)
                .Size = New Size(320, 118)
            End With
        Next i

        Dim strlblKey() As String = {"H", "O", "V", "E"}
        Dim lblHorizon(1) As Label
        For i = 0 To 1
            lblHorizon(i) = New Label
            With lblHorizon(i)
                .Parent = grpMain(i)
                .Text = "水平(&" & strlblKey(i) & "):"
                .Location = New Point(88, 32)
                .Size = New Size(80, 23)
            End With
        Next i

        Dim lblVertical(1) As Label
        For i = 0 To 1
            lblVertical(i) = New Label
            With lblVertical(i)
                .Parent = grpMain(i)
                .Text = "垂直(&" & strlblKey(i) & "):"
                .Location = New Point(88, 72)
                .Size = New Size(80, 23)
            End With
        Next i

        Dim strPicFileName() As String = {"mspaintExtend", "mspaintDistort"}

        Dim picExtend(1) As PictureBox
        Dim lblUnit1(1) As Label
        For i = 0 To 1
            picExtend(i) = New PictureBox
            With picExtend(i)
                .Parent = grpMain(0)
                .Image = New Bitmap(Me.GetType(), strPicFileName(0) & (i + 1).ToString & ".ico")
                .Size = New Size(32, 32)
                .Location = New Point(16, 24 + i * 40)
            End With

            txtboxExtend(i) = New TextBox
            With txtboxExtend(i)
                .Parent = grpMain(0)
                .Text = "100"
                .Location = New Point(176, 32 + i * 40)
                .Size = New Size(64, 25)
                .MaxLength = 4
            End With
            AddHandler txtboxExtend(i).KeyDown, AddressOf TxtboxKeyDown

            lblUnit1(i) = New Label
            With lblUnit1(i)
                .Parent = grpMain(0)
                .Text = "%"
                .Location = New Point(248, 40 + i * 40)
                .Size = New Size(16, 23)
            End With
        Next i

        Dim picDistort(1) As PictureBox
        Dim lblUnit2(1) As Label
        For i = 0 To 1
            picDistort(i) = New PictureBox
            With picDistort(i)
                .Parent = grpMain(1)
                .Image = New Bitmap(Me.GetType(), strPicFileName(1) & (i + 1).ToString & ".ico")
                .Size = New Size(32, 32)
                .Location = New Point(16, 24 + i * 40)
            End With
            txtboxDistort(i) = New TextBox
            With txtboxDistort(i)
                .Parent = grpMain(1)
                .Text = "0"
                .Location = New Point(176, 32 + i * 40)
                .Size = New Size(64, 25)
                .MaxLength = 3
            End With
            AddHandler txtboxDistort(i).KeyDown, AddressOf TxtboxKeyDown

            lblUnit2(i) = New Label
            With lblUnit2(i)
                .Parent = grpMain(1)
                .Text = "度"
                .Location = New Point(248, 40 + i * 40)
                .Size = New Size(16, 23)
            End With
        Next i

        Dim btnOK As New Button
        With btnOK
            .Parent = Me
            .Text = "確定"
            .Location = New Point(352, 16)
            .Size = New Size(96, 32)
        End With
        AddHandler btnOK.Click, AddressOf btnOKOnClick

        Dim btnCancel As New Button
        With btnCancel
            .Parent = Me
            .Text = "取消"
            .Location = New Point(352, 56)
            .Size = New Size(96, 32)
        End With
        AddHandler btnCancel.Click, AddressOf btnCancelOnClick
    End Sub
    Private Sub TxtboxKeyDown(ByVal obj As Object, ByVal kea As KeyEventArgs)
        If kea.KeyData = Keys.Enter Then
            Dim ea As EventArgs
            btnOKOnClick(obj, ea.Empty)
        End If
    End Sub
    Private Sub btnOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Const strTitle As String = "小畫家重製版"
        Try
            ExtX = Val(txtboxExtend(0).Text)
            ExtY = Val(txtboxExtend(1).Text)
            DisX = Val(txtboxDistort(0).Text)
            DisY = Val(txtboxDistort(1).Text)

        Catch ex As Exception
            MessageBox.Show("請輸入一整數.", strTitle, _
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End Try

        If ExtX < 1 OrElse ExtX > 500 OrElse ExtY < 1 OrElse ExtY > 500 Then
            MessageBox.Show("請輸入一個在1到500之間的整數.", strTitle, _
                             MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If DisX < -89 OrElse DisX > 89 OrElse DisY < -89 OrElse DisY > 89 Then
            MessageBox.Show("請輸入一個在-89到89之間的整數.", strTitle, _
                             MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If ExtX <> 100 OrElse ExtY <> 100 Then bExt = True Else bExt = False
        If DisX <> 0 OrElse DisY <> 0 Then bDis = True Else bDis = False

        DialogResult = DialogResult.OK
    End Sub

    Private Sub btnCancelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.Cancel
    End Sub
End Class
