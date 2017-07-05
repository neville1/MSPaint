Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintStatusbar
    '�����w�q���󪬺A�C���~�[�Φ欰
    Inherits PaintColorArea
    Protected sbar As StatusBar         '���A�C
    Protected sbpMenu, sbpLocation, sbpRectangle As StatusBarPanel   '���O����,�e����m,ø�s�d��
    Sub New()
        '���A�C�غc��

        sbar = New StatusBar
        sbar.Parent = Me
        sbar.ShowPanels = True

        sbpMenu = New StatusBarPanel
        sbpMenu.Text = "�p�ݻ����A�Ы��@�U[����]�\�����[�����D�D]�C"
        sbpMenu.AutoSize = StatusBarPanelAutoSize.Spring

        sbpLocation = New StatusBarPanel
        sbpRectangle = New StatusBarPanel

        sbar.Panels.AddRange(New StatusBarPanel() {sbpMenu, sbpLocation, sbpRectangle})
    End Sub
    Protected Overrides Sub ControldotOnMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '�w�q����ø�Ϥu�@�Ϥj�p��,�b���A�C����ܪ���T
        MyBase.ControldotOnMove(obj, mea)

        If mea.Button <> MouseButtons.Left Then Return
        Dim sz As Point = pnWorkarea.Parent.PointToClient(MousePosition)
        sbpRectangle.Text = sz.X.ToString & "x" & sz.Y.ToString
    End Sub
    Protected Overrides Sub ToolbarHelp(ByVal obj As Object, ByVal ea As EventArgs)
        '�мg PaintToolbar.vb,�w�q��ƹ����ܤu��C����W���,���A�C�ҭn��ܪ��T��.
        Dim radbtnHelp As NoFocusRadbtn = DirectCast(obj, NoFocusRadbtn)

        Select Case CInt(radbtnHelp.Name)
            Case 1
                sbpMenu.Text = "��Ϥ������󳡥��A�[�H���ʡA�ƻs�νs��C"
            Case 2
                sbpMenu.Text = "��Ϥ����x�γ����A�[�H���ʡA�ƻs�νs��C"
            Case 3
                sbpMenu.Text = "�ϥΩҿ諸������Ϊ��A�M�������Ϥ��C"
            Case 4
                sbpMenu.Text = "�ϥΥثe��ø�Ϧ�ΡA��J�Y�Ӱϰ�C"
            Case 5
                sbpMenu.Text = "�D��Ϥ������@���C���ø�ϡC"
            Case 6
                sbpMenu.Text = "�ܧ󭿲v�C"
            Case 7
                sbpMenu.Text = "�H�@�Ӽe�׹����j�pø�X���N�Ϊ����u���C"
            Case 8
                sbpMenu.Text = "�ϥΩҿ諸��l�Ϊ��M�j�p��ø�ϡC"
            Case 9
                sbpMenu.Text = "�Ωҿ諸�Q��j�p��ø�ϡC"
            Case 10
                sbpMenu.Text = "�N��r�K��Ϥ����C"
            Case 11
                sbpMenu.Text = "�Ωҿ諸�u���e��ø�s���u�C"
            Case 12
                sbpMenu.Text = "�Ωҿ諸�u���e��ø�s���u�C"
            Case 13
                sbpMenu.Text = "�Ωҿ諸��J�˦�ø�s�x�ΡC"
            Case 14
                sbpMenu.Text = "�Ωҿ諸��J�˦�ø�s�h��ΡC"
            Case 15
                sbpMenu.Text = "�Ωҿ諸��J�˦�ø�s���ΡC"
            Case 16
                sbpMenu.Text = "�Ωҿ諸��J�˦�ø�s�ꨤ�x�ΡC"
        End Select
    End Sub
    Protected Overrides Sub ToolbarHelpLeave(ByVal obj As Object, ByVal ea As EventArgs)
        '���}�u��C�����,���A�C����ܪ��T��
        sbpMenu.Text = "�p�ݻ����A�Ы��@�U[����]�\�����[�����D�D]�C"
    End Sub
End Class
