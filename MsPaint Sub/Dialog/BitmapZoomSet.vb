Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class BitmapZoomSet
    '本類別定義關於自訂縮放對話方塊的外觀及動作
    Inherits Form

    Private radbtnZoomBase(4) As RadioButton
    Private lblCurrentSize As Label
    Private ZoomSet As Integer

    ReadOnly Property ZoomBase() As Integer
        Get
            Return ZoomSet
        End Get
    End Property

    Sub New(ByVal Zoom As Integer)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        AutoScaleBaseSize = New Size(6, 18)
        Size = New Size(408, 192)
        Text = "自訂縮放"
        ZoomSet = Zoom

        Dim lblCurrent As New Label
        With lblCurrent
            .Parent = Me
            .Text = "目前縮放比例:"
            .Location = New Point(24, 16)
            .Size = New Size(104, 16)
        End With

        lblCurrentSize = New Label
        With lblCurrentSize
            .Parent = Me
            .Text = (Zoom * 100).ToString & "%"
            .Location = New Point(208, 16)
            .Size = New Size(40, 16)
        End With

        Dim grpZoomBase As New GroupBox
        With grpZoomBase
            .Parent = Me
            .Text = "縮放成"
            .Location = New Point(16, 40)
            .Size = New Size(264, 88)
        End With

        Dim strZoomBase() As String = {"100%(&1)", "200%(&2)", "400%(&4)", "600%(&6)", "800%(&8)"}

        Dim i, cx, cy As Integer
        For i = 0 To 4
            radbtnZoomBase(i) = New RadioButton
            With radbtnZoomBase(i)
                .Parent = grpZoomBase
                .Name = strZoomBase(i).Chars(0).ToString
                .Text = strZoomBase(i)
                .Size = New Size(80, 16)
                If (i + 1) Mod 2 <> 0 Then cy = 32 Else cy = 56
                .Location = New Point((i \ 2) * 80 + 8, cy)
                If .Name = Zoom.ToString Then .Checked = True
            End With
            AddHandler radbtnZoomBase(i).CheckedChanged, AddressOf ZoomBaseChangedOnClick
        Next i

        Dim btnOK As New Button
        With btnOK
            .Parent = Me
            .Text = "確定"
            .Location = New Point(288, 8)
            .Size = New Size(104, 32)
        End With
        AddHandler btnOK.Click, AddressOf ButtonOKOnClick

        Dim btnCancel As New Button
        With btnCancel
            .Parent = Me
            .Text = "取消"
            .Location = New Point(288, 48)
            .Size = New Size(104, 32)
        End With
        AddHandler btnCancel.Click, AddressOf ButtonCancelOnClick

    End Sub
    Private Sub ZoomBaseChangedOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim radbtn As RadioButton = DirectCast(obj, RadioButton)
        ZoomSet = CInt(radbtn.Name)
        lblCurrentSize.Text = (ZoomSet * 100).ToString & "%"
    End Sub
    Private Sub ButtonOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.OK
    End Sub
    Private Sub ButtonCancelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        DialogResult = DialogResult.Cancel
    End Sub
End Class
