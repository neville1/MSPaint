Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithMenuColor
    Inherits PaintWithMenuImage

    Private clrdlg As ColorDialog

    Sub New()
        Dim mi As New MenuItem("��m(&C)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miColorEdit As New MenuItemHelp("�s���m(&E)...")
        With miColorEdit
            .HelpPanel = sbpMenu
            .HelpText = "�إ߷s���C��C"
        End With
        Menu.MenuItems(index).MenuItems.Add(miColorEdit)
        AddHandler miColorEdit.Click, AddressOf ColorEditOnClick

        clrdlg = New ColorDialog     '�b�{�������,�Ȩϥγ�@����B�z
    End Sub
    Protected Overrides Sub ColorEditOnClick(ByVal obj As Object, ByVal ea As EventArgs)

        If clrdlg.ShowDialog() = DialogResult.OK Then
            clrFore = clrdlg.Color
        End If

        MyBase.ColorEditOnClick(obj, ea)
        '�o�欰���I�������H�z��.��@�p�e�a,�J���|�N�ۭq�C��ΨӨ��N ColorArea�Ҧs�b���C��
        '�Ʀܤ]���|���x�s�ۭq�C�⪺���p,�u�O��ª��N�ҿ諸�C��,�ΨӴ��N��L���Ĥ@���C��.
        '���|�o�˳]�p�O?
    End Sub
End Class
