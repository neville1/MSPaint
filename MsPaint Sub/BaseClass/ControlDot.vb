Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class ControlDot
    Inherits Control

    Sub New()
        setstyle(ControlStyles.Opaque Or ControlStyles.Selectable, False)

        Me.Size = New Size(3, 3)
        Me.Enabled = False
        AddHandler Me.Paint, AddressOf ControlDotOnPaint
    End Sub
    Private Sub ControlDotOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        If Me.Enabled Then
            Me.BackColor = Color.DarkBlue
        Else
            Me.BackColor = Color.White
        End If
    End Sub
End Class
