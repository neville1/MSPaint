Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithMenuAbout
    Inherits PaintWithMenuColor

    Sub New()
        Dim mi As New MenuItem("����(&H)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miAboutProg As New MenuItemHelp("����p�e�a���s��(&A)")
        With miAboutProg
            .HelpPanel = sbpMenu
            .HelpText = "��ܵ{����T�A�����s���H�Ϊ��v�C"
        End With
        Menu.MenuItems(index).MenuItems.Add(miAboutProg)
        AddHandler miAboutProg.Click, AddressOf MenuAboutProgOnClick
    End Sub
    Private Sub MenuAboutProgOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim dlg As New AboutProgram

        dlg.ShowDialog()
    End Sub

End Class
