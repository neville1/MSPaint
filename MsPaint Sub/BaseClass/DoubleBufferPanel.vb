Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class DoubleBufferPanel
    Inherits Panel

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or _
        ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True)
    End Sub
End Class
