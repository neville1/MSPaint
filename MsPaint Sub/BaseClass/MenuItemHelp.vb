Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class MenuItemHelp
    Inherits MenuItem

    Private sbpHelpPanel As StatusBarPanel
    Private sbpHelpText As String
    Sub New()
    End Sub
    Sub New(ByVal strText As String)
        MyBase.New(strText)
    End Sub
    Sub New(ByVal strText As String, ByVal HelpBar As StatusBarPanel)
        MyBase.New(strText)

        sbpHelpPanel = HelpBar
    End Sub
    Property HelpPanel() As StatusBarPanel
        Get
            Return sbpHelpPanel
        End Get
        Set(ByVal Value As StatusBarPanel)
            sbpHelpPanel = Value
        End Set
    End Property
    Property HelpText() As String
        Get
            Return sbpHelpText
        End Get
        Set(ByVal Value As String)
            sbpHelpText = Value
        End Set
    End Property
    Protected Overrides Sub OnSelect(ByVal ea As EventArgs)
        MyBase.OnSelect(ea)
        If Not HelpPanel Is Nothing Then
            sbpHelpPanel.Text = HelpText
        End If
    End Sub
End Class
