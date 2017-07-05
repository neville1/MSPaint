Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintColorArea
    '本類別定義關於色塊的外觀及行為
    Inherits PaintToolbarSub

    Protected pnColorArea As Panel                    '色彩列
    Protected clrFore, clrBack As Color               '前景色.背景色
    Protected btnColorFore, btnColorBack As Button    '控制使用者所見的前景,背景表示

    Private pbColor(27) As PictureBox

    Sub New()
        '色塊區元件建構元
        clrFore = Color.Black
        clrBack = Color.White                        '預設顏色,黑底白字.

        Dim sy As Integer
        If bStatusbar Then sy = 22 Else sy = 0

        pnColorArea = New Panel
        With pnColorArea
            .Parent = Me
            .BorderStyle = BorderStyle.None
            .Size = New Size(ClientSize.Width, 50)
            .Location = New Point(0, ClientSize.Height - 50 - sy)
        End With

        Dim pbColorarea As New PictureBox
        With pbColorarea
            .Parent = pnColorArea
            .BorderStyle = BorderStyle.Fixed3D
            .BackColor = Color.Gainsboro
            .Size = New Size(30, 32)
            .Location = New Point(0, 10)
        End With

        btnColorFore = New Button
        With btnColorFore
            .Parent = pbColorarea
            .BackColor = clrFore
            .Size = New Size(14, 14)
            .Location = New Point(3, 3)
            .Enabled = False
        End With

        btnColorBack = New Button                     '這二個元件,表示使用者目前選取的顏色
        With btnColorBack
            .Parent = pbColorarea
            .BackColor = clrBack
            .Size = New Size(14, 14)
            .Location = New Point(10, 11)
            .Enabled = False
        End With

        Dim i As Integer = 27

        Dim strColor() As String = {"Black", "Gray", "DarkRed", _
        "Olive", "Green", "Teal", "DarkBlue", "Purple", "DarkOliveGreen", _
        "DarkSlateGray", "DodgerBlue", "Navy", "MidnightBlue", "SaddleBrown", _
        "White", "Silver", "Red", "Yellow", "Lime", "Cyan", "Blue", "Fuchsia", _
        "Gold", "SpringGreen", "PaleTurquoise", "DarkTurquoise", "Crimson", "DarkOrange"}
        '此處定義,關於色塊區的26種顏色.

        For i = 0 To 27
            pbColor(i) = New PictureBox
            pbColor(i).Parent = pnColorArea
            pbColor(i).BackColor = Color.FromName(strColor(i))
            pbColor(i).Size = New Size(16, 16)
            pbColor(i).BorderStyle = BorderStyle.Fixed3D
            If i <= 13 Then
                pbColor(i).Location = New Point(32 + i * 16, 10)
            Else
                pbColor(i).Location = New Point(32 + (i - 14) * 16, 26)
            End If
            AddHandler pbColor(i).MouseDown, AddressOf SelectColorOnClick
        Next i
    End Sub
    Protected Overrides Sub FormSizeChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '定義因視覺元件變更時,色塊區的座標及位移
        MyBase.FormSizeChanged(obj, ea)
        Dim tx, cy, sy As Integer
        If bColorbar Then cy = 50 Else cy = 0
        If bStatusbar Then sy = 22 Else sy = 0

        pnColorArea.Location = New Point(0, ClientSize.Height - 50 - sy)
    End Sub
    Private Sub SelectColorOnClick(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '色塊區主要行為,由使用者點選顏色使用.
        Dim pb As PictureBox = DirectCast(obj, PictureBox)

        Select Case mea.Button
            Case MouseButtons.Left
                clrFore = pb.BackColor
                btnColorFore.BackColor = clrFore
            Case MouseButtons.Right
                clrBack = pb.BackColor
                btnColorBack.BackColor = clrBack
        End Select
    End Sub
    Protected Overridable Sub ColorEditOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '由功能表中的->色彩->編輯色彩,所影響的色塊區顏色變更

        pbColor(0).BackColor = clrFore
        btnColorFore.BackColor = clrFore
        '由 PaintWithMenuColor.vb 更動色盤第一個顏色
    End Sub

End Class
