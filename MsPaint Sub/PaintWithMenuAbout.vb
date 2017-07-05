Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithMenuAbout
    Inherits PaintWithMenuColor

    Sub New()
        Dim mi As New MenuItem("說明(&H)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miAboutProg As New MenuItemHelp("關於小畫家重製版(&A)")
        With miAboutProg
            .HelpPanel = sbpMenu
            .HelpText = "顯示程式資訊，版本編號以及版權。"
        End With
        Menu.MenuItems(index).MenuItems.Add(miAboutProg)
        AddHandler miAboutProg.Click, AddressOf MenuAboutProgOnClick
    End Sub
    Private Sub MenuAboutProgOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim dlg As New AboutProgram

        dlg.ShowDialog()
    End Sub

End Class
