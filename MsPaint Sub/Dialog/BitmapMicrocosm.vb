Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class BitmapMicrocosm
    '本類別定義關於縮圖Form中的外觀及動作
    Inherits DoubleBufferForm

    Sub New(ByVal bmp As Bitmap)
        FormBorderStyle = FormBorderStyle.SizableToolWindow
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        Size = New Size(182, 242)
        Text = "縮圖"

        Dim pb As New PictureBox
        With pb
            .Parent = Me
            .BorderStyle = BorderStyle.Fixed3D
            .Dock = DockStyle.Fill
            .Image = bmp
        End With
    End Sub
End Class
