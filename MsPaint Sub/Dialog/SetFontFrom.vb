Imports System
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms

Public Class SetFontFrom
    Inherits Form

    Private cboboxFontType, cboboxFontSize, cboboxFontCode As ComboBox
    Private chkboxFontStyle(2) As NoFocusChkeckBox
    Private chkboxFontVertical As NoFocusChkeckBox

    Private nBold, nItalic, nUnderline As Integer
    Private nFontStyle As Integer = 0
    Private fmt() As FontFamily = FontFamily.Families

    ReadOnly Property FontName() As String
        Get
            Return cboboxFontType.Text
        End Get
    End Property
    ReadOnly Property FontSize() As Single
        Get
            Return CSng(cboboxFontSize.Text)
        End Get
    End Property
    ReadOnly Property MyFontStyle() As Integer
        Get
            Return nFontStyle
        End Get
    End Property

    Event UserFontChanged As EventHandler

    Sub New()
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        AutoScaleBaseSize = New Size(6, 18)
        Size = New Size(592, 72)
        Text = "¦r«¬"

        cboboxFontType = New ComboBox
        With cboboxFontType
            .Parent = Me
            .Location = New Point(8, 8)
            .Size = New Size(216, 23)
            .DropDownStyle = ComboBoxStyle.DropDownList
        End With

        Dim i As Integer
        Dim strFontName(fmt.GetUpperBound(0)) As String
        For i = 0 To fmt.GetUpperBound(0)
            strFontName(i) = fmt(i).Name
            cboboxFontType.Items.Add(strFontName(i))
        Next i
        cboboxFontType.Text = cboboxFontType.Items.Item(fmt.GetUpperBound(0) - 2)
        AddHandler cboboxFontType.SelectionChangeCommitted, AddressOf FontTypeChanged

        cboboxFontSize = New ComboBox
        With cboboxFontSize
            .Parent = Me
            .Location = New Point(240, 8)
            .Size = New Size(72, 23)
            .DropDownStyle = ComboBoxStyle.DropDownList
        End With
        Dim FontSize() As Integer = {8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72}
        For i = 0 To FontSize.GetUpperBound(0)
            cboboxFontSize.Items.Add(FontSize(i).ToString)
        Next i
        cboboxFontSize.Text = cboboxFontSize.Items.Item(1)
        AddHandler cboboxFontSize.SelectionChangeCommitted, AddressOf FontSizeChanged

        cboboxFontCode = New ComboBox
        With cboboxFontCode
            .Parent = Me
            .Location = New Point(320, 8)
            .Size = New Size(160, 23)
            .Text = "Big-5"
            .DropDownStyle = ComboBoxStyle.DropDownList
        End With

        Dim strFontStyle() As String = {"Bold", "Italic", "Underline"}
        For i = 0 To 2
            chkboxFontStyle(i) = New NoFocusChkeckBox
            With chkboxFontStyle(i)
                .Parent = Me
                .Name = strFontStyle(i)
                .Appearance = Appearance.Button
                .Location = New Point(488 + i * 24, 8)
                .Size = New Size(24, 24)
                .Image = New Bitmap(Me.GetType(), "Font" & strFontStyle(i) & ".bmp")
            End With
            AddHandler chkboxFontStyle(i).CheckedChanged, AddressOf FontStyleChanged
        Next i


        chkboxFontVertical = New NoFocusChkeckBox
        With chkboxFontVertical
            .Parent = Me
            .Appearance = Appearance.Button
            .Location = New Point(560, 8)
            .Size = New Size(24, 24)
            .Image = New Bitmap(Me.GetType(), "FontVertical.bmp")
        End With

        nBold = 0
        nItalic = 0
        nUnderline = 0
    End Sub
    Private Sub FontTypeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        Dim cbobox As ComboBox = DirectCast(obj, ComboBox)

        chkboxFontStyle(0).Enabled = fmt(cbobox.SelectedIndex()).IsStyleAvailable(FontStyle.Bold)
        chkboxFontStyle(1).Enabled = fmt(cbobox.SelectedIndex()).IsStyleAvailable(FontStyle.Italic)
        chkboxFontStyle(2).Enabled = fmt(cbobox.SelectedIndex()).IsStyleAvailable(FontStyle.Underline)

        RaiseEvent UserFontChanged(Me, EventArgs.Empty)
    End Sub
    Private Sub FontSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        RaiseEvent UserFontChanged(Me, EventArgs.Empty)
    End Sub
    Private Sub FontStyleChanged(ByVal obj As Object, ByVal ea As EventArgs)
        Dim chkbox As NoFocusChkeckBox = DirectCast(obj, NoFocusChkeckBox)
        Select Case chkbox.Name
            Case "Bold"
                If chkbox.Checked Then nBold = 1 Else nBold = 0
            Case "Italic"
                If chkbox.Checked Then nItalic = 2 Else nItalic = 0
            Case "Underline"
                If chkbox.Checked Then nUnderline = 4 Else nUnderline = 0
        End Select

        nFontStyle = nBold Or nItalic Or nUnderline

        RaiseEvent UserFontChanged(Me, EventArgs.Empty)
    End Sub
End Class
