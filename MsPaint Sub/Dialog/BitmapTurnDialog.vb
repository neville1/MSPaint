Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class BitmapTurnDialog
    '本類別定義關於旋轉/翻轉對話方塊的外觀及動作
    Inherits Form

    Private SelectCmd As Integer = 1
    Private btnOK, btnCancel As Button
    Private grpRangeAngle As GroupBox
    ReadOnly Property SelectCommand() As Integer
        Get
            Return SelectCmd
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
        Size = New Size(408, 236)
        Text = "翻轉及旋轉"

        Dim cxText As Integer = AutoScaleBaseSize.Width
        Dim cyText As Integer = AutoScaleBaseSize.Height

        grpRangeAngle = New GroupBox
        With grpRangeAngle
            .Parent = Me
            .Location = New Point(90, 105)
            .Size = New Size(120, 66)
        End With

        Dim GrpOption As New GroupBox
        With GrpOption
            .Parent = Me
            .Text = Me.Text
            .Location = New Point(10, 15)
            .Size = New Size(254, 170)
        End With

        Dim i As Integer
        Dim strRadbtn() As String = {"水平翻轉(&F)", "垂直翻轉(&V)", "旋轉角度(&R)", _
        "90 度旋轉(&9)", "180 度旋轉(&1)", "270 度旋轉(&2)"}
        Dim radbtnCommand(5) As RadioButton
        For i = 0 To 5
            radbtnCommand(i) = New RadioButton
            If i < 3 Then
                With radbtnCommand(i)
                    .Name = (i + 1).ToString
                    .Parent = GrpOption
                    .Text = strRadbtn(i)
                    .Location = New Point(10, 20 + i * 24)
                    .Size = New Size(120, 24)
                End With
            Else
                With radbtnCommand(i)
                    .Name = (i + 1).ToString
                    .Parent = grpRangeAngle
                    .Text = strRadbtn(i)
                    .Location = New Point(0, (i - 3) * 24)
                    .Size = New Size(140, 24)
                End With
            End If
            AddHandler radbtnCommand(i).CheckedChanged, AddressOf RadioButtonCheckChanged
        Next i
        radbtnCommand(0).Checked = True
        radbtnCommand(3).Checked = True
        grpRangeAngle.Enabled = False
        SelectCmd = 1

        btnOK = New Button
        With btnOK
            .Parent = Me
            .Text = "確定"
            .TextAlign = ContentAlignment.BottomCenter
            .Location = New Point(290, 20)
            .Size = New Size(100, 30)
        End With
        AddHandler btnOK.Click, AddressOf btnOKOnClick

        btnCancel = New Button
        With btnCancel
            .Parent = Me
            .Text = "取消"
            .TextAlign = ContentAlignment.BottomCenter
            .Location = New Point(290, 60)
            .Size = New Size(100, 30)
        End With
        AddHandler btnCancel.Click, AddressOf btnCancelOnClick
    End Sub
    Private Sub RadioButtonCheckChanged(ByVal obj As Object, ByVal ea As EventArgs)
        Dim radbtn As RadioButton = DirectCast(obj, RadioButton)
        If radbtn.Checked Then
            SelectCmd = CInt(radbtn.Name)
        End If
        If SelectCmd >= 3 Then
            grpRangeAngle.Enabled = True
        Else
            grpRangeAngle.Enabled = False
        End If
    End Sub
    Private Sub btnOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.OK
    End Sub
    Private Sub btnCancelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.Cancel
    End Sub
End Class
