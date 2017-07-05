Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class DoubleBufferForm
    Inherits Form
    Private CursorPath As String

    Property GetCursorPath() As String
        Get
            Return CursorPath
        End Get
        Set(ByVal Value As String)
            CursorPath = Value
        End Set
    End Property

    Sub New()
        setstyle( _
           ControlStyles.AllPaintingInWmPaint Or _
           ControlStyles.DoubleBuffer Or _
           ControlStyles.ResizeRedraw Or _
           ControlStyles.UserPaint, _
           True)
    End Sub
End Class
