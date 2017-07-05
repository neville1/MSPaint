Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class NoFocusChkeckBox
    Inherits CheckBox
    Sub New()
        setstyle(ControlStyles.Selectable, False)
    End Sub
End Class
