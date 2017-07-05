Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintCommandWithText
    '�����O�B�z�u��C����r����\�઺��@
    Inherits PaintCommandWithEdit

    Private pnText As TextPanel                 '��r����
    Protected frmFontSet As SetFontFrom         '��r�u��C
    Protected bText As Boolean                  '��r����O�_�s�b 
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)

        If bText Then TextAreaComplete()

        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 10
                RemoveHandler pnWorkarea.MouseDown, AddressOf TxtboxOnMouseDown
        End Select

        Select Case ToolbarCommand
            Case 10                 '��r���
                AddHandler pnWorkarea.MouseDown, AddressOf TxtboxOnMouseDown
                If ZoomBase <> 1 Then         '��r�\��ȯ�b�L�Y��񪺱��p�U�ϥ� 
                    Beep()
                    radbtnToolbar(LastCommand - 1).Checked = True
                    ToolbarRadbtnSelect(radbtnToolbar(LastCommand - 1), ea.Empty)
                End If
        End Select
    End Sub
    Private Sub TxtboxOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If bText Then
            TextAreaComplete()
        Else
            CreateTextArea(mea)
        End If
    End Sub
    Private Sub CreateTextArea(ByVal mea As MouseEventArgs)
        '�إߤ�r������
        frmFontSet = New SetFontFrom
        Me.AddOwnedForm(frmFontSet)
        frmFontSet.Show()
        AddHandler frmFontSet.UserFontChanged, AddressOf UserFontChanged

        Dim fnt As New Font("�s�ө���", 9)
        Dim grfx As Graphics = Me.CreateGraphics()
        Dim sz As Size = New Size(CInt(grfx.MeasureString("A", fnt).Width), _
                                  CInt(grfx.MeasureString("A", fnt).Height))
        grfx.Dispose()

        pnText = New TextPanel(pnWorkarea, sz, fnt, clrFore, bTransParent)
        pnText.Location = New Point(mea.X, mea.Y)
        bText = True
    End Sub
    Private Sub UserFontChanged(ByVal obj As Object, ByVal ea As EventArgs)
        '��r�u��C,�r���]�w�ܧ�
        Dim fs As FontStyle = frmFontSet.MyFontStyle
        pnText.SetFont = New Font(frmFontSet.FontName, frmFontSet.FontSize, fs, GraphicsUnit.Point)
    End Sub
    Protected Sub TextAreaComplete()
        '�����r������,�ñN�欰��ƥ[�Jø�Ϧ欰��
        frmFontSet.Dispose()

        Dim bmpBack As New Bitmap(pnText.Width - 6, pnText.Height - 6)
        Dim grfx As Graphics = Graphics.FromImage(bmpBack)
        Dim br As New SolidBrush(clrFore)

        If bTransParent Then
            Dim rectSrc As New Rectangle(pnText.Location.X + 3, pnText.Location.Y + 3, _
                                         bmpBack.Width, bmpBack.Height)
            Dim rectDst As New Rectangle(0, 0, bmpBack.Width, bmpBack.Height)
            grfx.DrawImage(WorkareaChangeBmp(), rectDst, rectSrc, GraphicsUnit.Pixel)
        Else
            grfx.Clear(clrBack)
        End If
        Dim strText() As String = pnText.GetText

        Dim i As Integer
        Dim Height As Integer = 0
        For i = 0 To strText.GetUpperBound(0)
            grfx.DrawString(strText(i), pnText.SetFont, br, 0, Height)
            Height += CInt(grfx.MeasureString(strText(i), pnText.SetFont).Height)
        Next i
        grfx.Dispose()

        If strText.Length > 0 Then
            Dim GrapMethod As New GraphicsItem
            With GrapMethod
                .DataType = "B"
                .IsFont = True
                .Image = bmpBack
                .ImageLocation = New Point(pnText.Location.X + 1, pnText.Location.Y + 3)
                .ImageSize = bmpBack.Size
            End With
            GraphicsData.Add(GrapMethod)
        End If

        pnText.Dispose()
        bText = False
        pnWorkarea.Refresh()
        '�N��r��������ϧ�,�g�J GraphircsData ��
    End Sub
End Class
