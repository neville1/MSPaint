Imports System
Imports System.Drawing
Imports System.Collections
Imports System.Windows.Forms
Public Class PaintCommandWithOther
    '�����O��@�u��C���|���B�z������(���,����,��j��)
    Inherits PaintCommandWithText

    Private clrWork As Color

    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 4
                RemoveHandler pnWorkarea.MouseDown, AddressOf FillColorOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf FillColorOnMouseMove
            Case 5
                RemoveHandler pnWorkarea.MouseDown, AddressOf GetPointColorOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf GetPointColorOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf GetPointColorOnMouseUp
            Case 6
                RemoveHandler pnWorkarea.MouseMove, AddressOf ZoomChangeOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf ZoomChangeOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 4   '��R�C��
                AddHandler pnWorkarea.MouseDown, AddressOf FillColorOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf FillColorOnMouseMove
            Case 5   '���o�C��
                AddHandler pnWorkarea.MouseDown, AddressOf GetPointColorOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf GetPointColorOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf GetPointColorOnMouseUp
            Case 6    '��j��
                AddHandler pnWorkarea.MouseMove, AddressOf ZoomChangeOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf ZoomChangeOnMouseUp
        End Select
    End Sub
    Private Sub GetPointColorOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '�ھ�ø�ϱ��Ψ���

        Dim bmpback As Bitmap = WorkareaChangeBmp()

        Dim GetPoint As Point = GetScalePoint(New Point(mea.X, mea.Y))
        Select Case mea.Button
            Case MouseButtons.Left
                clrFore = bmpback.GetPixel(GetPoint.X, GetPoint.Y)
                btnColorFore.BackColor = clrFore
            Case MouseButtons.Right
                clrBack = bmpback.GetPixel(GetPoint.X, GetPoint.Y)
                btnColorBack.BackColor = clrBack
        End Select

        bmpback.Dispose()
    End Sub
    Private Sub GetPointColorOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub GetPointColorOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        radbtnToolbar(LastCommand - 1).Checked = True
        Dim ea As EventArgs
        ToolbarRadbtnSelect(radbtnToolbar(LastCommand - 1), ea.Empty)  '��ܤW�@�өR�O
    End Sub

    Private Sub FillColorOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '���\���@
        Dim bmpback As Bitmap = WorkareaChangeBmp()

        Dim ptPo As Point = GetScalePoint(New Point(mea.X, mea.Y))
        Dim pt, ptSum As Point
        Dim Width, Height As Integer
        Width = bmpback.Width
        Height = bmpback.Height

        clrWork = SelectUserColor(mea)

        Dim clrFill As Color = bmpback.GetPixel(ptPo.X, ptPo.Y)
        If clrWork.ToArgb = clrFill.ToArgb Then Return
        Dim Points As New Stack

        Dim ptDirEction() As Point = {New Point(0, -1), New Point(1, 0), _
                                      New Point(0, 1), New Point(-1, 0)} '�w�q��V
        Dim i As Integer
        Points.Push(ptPo)
        Do
            pt = Points.Pop()
            bmpback.SetPixel(pt.X, pt.Y, clrWork)

            For i = 0 To 3
                ptSum = New Point(pt.X + ptDirEction(i).X, pt.Y + ptDirEction(i).Y)
                If ptSum.X >= 0 AndAlso ptSum.X < Width AndAlso _
                   ptSum.Y >= 0 AndAlso ptSum.Y < Height Then
                    If bmpback.GetPixel(ptSum.X, ptSum.Y).ToArgb = clrFill.ToArgb Then
                        Points.Push(ptSum)
                    End If
                End If
            Next i
        Loop While Points.Count > 0

        WorkareaAddGraphicsItem(bmpback)
        '�o�q�{���X,²���O���X�Ҧ��{���X���a�B�F.
        '�t�׷��C,�O����l���j.�۱q�g�o�ӵ{���H��,�o�ؤ@���O�ۤv�Ҽ{�����I.
        '�]�� ExtFloodFill �o��API,�èS���Q�]�t�b .Net ��.
        '�ҥH�@�����J��Ԥ�,�쩳�n���n�I�s API�ӹF���o�ӥ\��.
        '��ӸչϧQ�� API �ӻs�@�ɵo�{,�n�����o�ӥ\��.�ܤ֭n�ϥ� CreateSolidBrush 
        'CreateCompatibleDC,CreateCompatibleBitmap,SelectObject,DeleteObject,��
        'ExtFloodFill.���M�i�H����,�t�פW�]�۷���,���o�ˤw�g���O�b�ǲ� VB.Net.
        '������P�˸m�L�����S�ʯ}�a���,�]���{���� grax = Graphics.FormImage() 
        '�o�˪��@�P�ʮ���.�ҥH���H�o�˪���k�ӻs�@�o�ӥ\��.
    End Sub
    Private Sub FillColorOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub ZoomChangeOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
    End Sub
    Private Sub ZoomChangeOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
    End Sub
    Protected Overridable Sub WorkareaAddGraphicsItem(ByVal bmp As Bitmap)
        '�� PaintWithMenuImage ��@,�N���v�v�s�Jø�Ϧ欰
    End Sub
End Class
