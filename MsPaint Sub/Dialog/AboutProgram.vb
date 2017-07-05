Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class AboutProgram
    '本類別僅為'關於小畫家'顯示對話方塊
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
        Text = "關於小畫家重製版"

        Dim lblProgName As New Label
        With lblProgName
            .Parent = Me
            .Location = New Point(72, 16)
            .Text = "小畫家重製版 For VB .Net"
            .AutoSize = True
        End With

        Dim lblProgVer As New Label
        With lblProgVer
            .Parent = Me
            .Location = New Point(256, 16)
            .Text = "版本:V 1.2"
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
            .Text = "本程式為學習 Programming Mircosoft Windows With Visual Basic .Net" & _
                    "所製作" & vbLf & "     歡迎使用，修改，散佈或剪貼其中的程式碼。" & vbLf & _
                    "                                                     " & _
                    "但由於其中的圖形，遊標，及配置方式皆由 Mircosoft 小畫家得來，" & _
                    "故請不要作為商業用途。                                         " & _
                    "也歡迎寫信給予小弟意見"
            .ReadOnly = True
            .TabStop = False
        End With

        Dim btnOK As New Button
        With btnOK
            .Parent = Me
            .Location = New Point(176, 206)
            .Size = New Size(75, 32)
            .Text = "確定"
            .TabIndex = 1
        End With
        AddHandler btnOK.Click, AddressOf ButtonOKOnClick
    End Sub
    Private Sub ButtonOKOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Me.Close()
    End Sub

End Class
