Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class AboutProgram
    '�����O�Ȭ�'����p�e�a'��ܹ�ܤ��
    Inherits Form

    Sub New()
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Location = Point.op_Addition(ActiveForm.Location, _
        Size.op_Addition(SystemInformation.CaptionButtonSize, _
        SystemInformation.FrameBorderSize))
        Size = New Size(432, 274)
        Text = "����p�e�a���s��"

        Dim lblProgName As New Label
        With lblProgName
            .Parent = Me
            .Location = New Point(72, 16)
            .Text = "�p�e�a���s�� For VB .Net"
            .AutoSize = True
        End With

        Dim lblProgVer As New Label
        With lblProgVer
            .Parent = Me
            .Location = New Point(256, 16)
            .Text = "����:V 1.2"
            .AutoSize = True
        End With

        Dim lblby As New Label
        With lblby
            .Parent = Me
            .Location = New Point(72, 182)
            .Text = "by"
            .AutoSize = True
        End With

        Dim lblAuthor As New Label
        With lblAuthor
            .Parent = Me
            .Location = New Point(96, 182)
            .Text = "Neville(neville.gee@msa.hinet.net)"
            .AutoSize = True
        End With

        Dim txtbox As New TextBox
        With txtbox
            .Parent = Me
            .Multiline = True
            .Location = New Point(72, 48)
            .Size = New Size(272, 132)
            .Text = "���{�����ǲ� Programming Mircosoft Windows With Visual Basic .Net" & _
                    "�һs�@" & vbLf & "     �w��ϥΡA�ק�A���G�ΰŶK�䤤���{���X�C" & vbLf & _
                    "                                                     " & _
                    "���ѩ�䤤���ϧΡA�C�СA�ΰt�m�覡�ҥ� Mircosoft �p�e�a�o�ӡA" & _
                    "�G�Ф��n�@���ӷ~�γ~�C                                         " & _
                    "�]�w��g�H�����p�̷N��"
            .ReadOnly = True
            .TabStop = False
        End With

        Dim btnOK As New Button
        With btnOK
            .Parent = Me
            .Location = New Point(176, 206)
            .Size = New Size(75, 32)
            .Text = "�T�w"
            .TabIndex = 1
        End With
        AddHandler btnOK.Click, AddressOf ButtonOKOnClick
    End Sub
    Private Sub ButtonOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Me.Close()
    End Sub

End Class
