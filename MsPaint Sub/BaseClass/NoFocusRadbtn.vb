Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class NoFocusRadbtn
    Inherits RadioButton

    Sub New()
        SetStyle(ControlStyles.Selectable, False)
    End Sub

End Class
