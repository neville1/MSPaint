Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithMenuColor
    Inherits PaintWithMenuImage

    Private clrdlg As ColorDialog

    Sub New()
        Dim mi As New MenuItem("色彩(&C)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miColorEdit As New MenuItemHelp("編輯色彩(&E)...")
        With miColorEdit
            .HelpPanel = sbpMenu
            .HelpText = "建立新的顏色。"
        End With
        Menu.MenuItems(index).MenuItems.Add(miColorEdit)
        AddHandler miColorEdit.Click, AddressOf ColorEditOnClick

        clrdlg = New ColorDialog     '在程式執行期,僅使用單一物件處理
    End Sub
    Protected Overrides Sub ColorEditOnClick(ByVal obj As Object, ByVal ea As EventArgs)

        If clrdlg.ShowDialog() = DialogResult.OK Then
            clrFore = clrdlg.Color
        End If

        MyBase.ColorEditOnClick(obj, ea)
        '這行為有點讓我難以理解.原作小畫家,既不會將自訂顏色用來取代 ColorArea所存在的顏色
        '甚至也不會有儲存自訂顏色的情況,只是單純的將所選的顏色,用來替代色盤的第一個顏色.
        '怎麼會這樣設計呢?
    End Sub
End Class
