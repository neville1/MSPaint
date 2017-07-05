Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintToolBar
    '�����O�w�q����u��C�~�[�Φ欰

    Inherits PaintWithRegistry

    Protected ToolBarCommand As Integer = 7             '�ϥΪ̤u��C�R�O���
    Protected LastCommand As Integer                    '�W�@�өR�O
    Protected pnToolbar As Panel                        '�u��C
    Protected radbtnToolbar(15) As NoFocusRadbtn        '�u��C�\��
    Sub New()
        '�u��C�غc��
        Dim tx, coly, sy As Integer
        If bToolbar Then tx = 60 Else tx = 0
        If bColorbar Then coly = 50 Else coly = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnToolbar = New Panel
        With pnToolbar
            .Parent = Me
            .BorderStyle = BorderStyle.None
            .Location = New Point(0, 0)
            .Size = New Size(60, ClientSize.Height - coly - sy)
        End With

        Dim TotTool As New ToolTip
        Dim strTotTool() As String = {"��ܥ��N�d��", "���", "�����/�m������", "��J��m", _
                                      "�D���C��", "��j��", "�]��", "����", "�Q�j", "��r", _
                                      "���u", "���u", "�x��", "�h���", "����", "�ꨤ�x��"}

        Dim i, cx As Integer
        For i = 0 To 15
            radbtnToolbar(i) = New NoFocusRadbtn
            With radbtnToolbar(i)
                .Name = (i + 1).ToString
                .Parent = pnToolbar
                .Appearance = Appearance.Button
                .Size = New Size(26, 26)
                .Image = New Bitmap(Me.GetType(), "toolbutton" & (i + 1).ToString & ".bmp")
                If (i + 1) Mod 2 <> 0 Then cx = 3 Else cx = 29
                .Location = New Point(cx, (i \ 2) * 26)
            End With
            TotTool.SetToolTip(radbtnToolbar(i), strTotTool(i))
            AddHandler radbtnToolbar(i).MouseEnter, AddressOf ToolbarHelp
            AddHandler radbtnToolbar(i).MouseLeave, AddressOf ToolbarHelpLeave
            AddHandler radbtnToolbar(i).Click, AddressOf ToolbarRadbtnSelect
        Next i
        radbtnToolbar(6).Checked = True
        pnWorkarea.Cursor = New Cursor(Me.GetType(), "Pen.cur")

    End Sub
    Protected Overrides Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '�w�q�]��ı�����ܧ�(�u��C,�����)�u��C�y�аt�m
        MyBase.FormSizeChanged(obj, ea)

        Dim tx, cy, sy As Integer
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnToolbar.Size = New Size(60, ClientSize.Height - cy - sy)
    End Sub
    Protected Overridable Sub ToolbarHelp(ByVal obj As Object, ByVal ea As EventArgs)
        '��ƹ����ܤu��C����W��,��ܸӤ���b���A�C����T.�� PaintStatusbar �мg
    End Sub
    Protected Overridable Sub ToolbarHelpLeave(ByVal obj As Object, ByVal ea As EventArgs)
        '��ƹ����X�u��C�����,��ܸӤ���b���A�C����T.�� PaintStatusbar �мg
    End Sub
    Protected Overridable Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '�w�q��ϥΪ̿�ܤu��C�����,�Ҳ��ͪ��欰
        Dim radbtnSelect As NoFocusRadbtn = DirectCast(obj, NoFocusRadbtn)

        LastCommand = ToolBarCommand               '���o�W�@�ӥ\��

        ToolBarCommand = CInt(radbtnSelect.Name)   '���o�ϥΪ̩ҿ�ܪ��\��
        Dim CursorName As String
        Select Case ToolBarCommand
            Case 1, 2, 10, 11, 12, 13, 14, 15, 16
                CursorName = "Cross.cur"
            Case 3
                CursorName = "Null.cur"
            Case 4
                CursorName = "FillColor.cur"
            Case 5
                CursorName = "Color.cur"
            Case 6
                CursorName = "ZoomSet.cur"
            Case 7
                CursorName = "Pen.cur"
            Case 8
                CursorName = "Brush.cur"
            Case 9
                CursorName = "Spray.cur"
        End Select
        pnWorkarea.Cursor = New Cursor(Me.GetType(), CursorName)
        '�o��ثe����,�ȩw�q�F,������ø�Ϥu�@�Ϫ��C�����,�D�n������y�{,�� Overrides�мg
    End Sub
End Class
