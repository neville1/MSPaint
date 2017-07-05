Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Windows.Forms

Public Class PaintCommandWithBrush
    '�����O��@�u��C�����󵧨�ۦ������B�z.
    '�����O�]��欰�۪�,�����ƻ�����ۦ��欰.
    Inherits PaintWorkareaOnPaint

    Private ptLast, ptNew As Point
    Private ptArea As Point
    Private clrWork As Color
    Private bmpBrush As Bitmap
    Private bmpOffest As Integer
    Private tmr As Timer
    Private arrlstPoint As ArrayList

    Private bTracking As Boolean = False        '�ƹ��즲

    Sub New()
        '�w�q�{���Ұʮ�,�w�]���u��C�ﶵ.
        arrlstPoint = New ArrayList

        AddHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
        AddHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
        AddHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
    End Sub
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '�H�мg�u��C����ƥ�,�w�q�u��C����Ҳ��ͪ��欰�e���Ψҥ~�B�z.

        MyBase.ToolbarRadbtnSelect(obj, ea)

        Select Case LastCommand
            Case 3
                RemoveHandler pnWorkarea.MouseDown, AddressOf EraseOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf EraseOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf EraseOnMouseUp
                RemoveHandler pnWorkarea.MouseLeave, AddressOf EraseOnMouseLeave
            Case 7
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
            Case 8
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawBrushOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawBrushOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawBrushOnMouseUp
                RemoveHandler pnWorkarea.MouseLeave, AddressOf DrawBrushOnMouseLeave
            Case 9
                RemoveHandler pnWorkarea.MouseDown, AddressOf DrawSprayOnMouseDown
                RemoveHandler pnWorkarea.MouseMove, AddressOf DrawSprayOnMouseMove
                RemoveHandler pnWorkarea.MouseUp, AddressOf DrawSprayOnMouseUp
        End Select

        Select Case ToolbarCommand
            Case 3 '�����
                AddHandler pnWorkarea.MouseDown, AddressOf EraseOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf EraseOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf EraseOnMouseUp
                AddHandler pnWorkarea.MouseLeave, AddressOf EraseOnMouseLeave
            Case 7 '�e��
                AddHandler pnWorkarea.MouseDown, AddressOf DrawPointOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawPointOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawPointOnMouseUp
            Case 8 '���� 
                AddHandler pnWorkarea.MouseDown, AddressOf DrawBrushOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawBrushOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawBrushOnMouseUp
                AddHandler pnWorkarea.MouseLeave, AddressOf DrawBrushOnMouseLeave
            Case 9 '�Q��
                AddHandler pnWorkarea.MouseDown, AddressOf DrawSprayOnMouseDown
                AddHandler pnWorkarea.MouseMove, AddressOf DrawSprayOnMouseMove
                AddHandler pnWorkarea.MouseUp, AddressOf DrawSprayOnMouseUp
        End Select
        '��g�Ʀ���,�p�̳̫�M�w�γo�Ӥ覡,�өw�q�u��C�\�઺�e��
        '���Mø�Ϧ欰��,���\�h�ۦ�����k,�γo�ؤ覡�өw�q,�����e���Ӧ@�άۦ����{���X
        '�]���y���{���X���Τj,���ثe�Ҽ{�����I�O,���q���y�{�M��,����\Ū
    End Sub
    Private Sub EraseOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()          '�M���]�_��,���ƥ\��Ҳ��ͪ��h�l�R�O

        bTracking = True
        ptArea = New Point(mea.X, mea.Y)
        ptLast = GetScalePoint(ptArea)

        arrlstPoint.Add(ptLast)
        '������\��,���ӥH���o�˪��\��,�u�ݭn�Q�εe���j�p�����ܤΨϥΩ���Y�i.
        '��ӵo�{,�p�G�γo�˪��覡�s�@,�|���ͤ��֪��ҥ~����,�]���৹�����T����ܨϥΪ̪��ʧ@
        '�o�بϥΤ@�s��s�򪺯x��,�Ӫ�ܾ�����欰���ʧ@
    End Sub
    Private Sub EraseOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()


        Dim UpdateArea As New Rectangle(ptArea.X - 2, ptArea.Y - 2, _
                                        EraseSize * ZoomBase + 4, EraseSize * ZoomBase + 4)
        pnWorkarea.Invalidate(UpdateArea)
        pnWorkarea.Update()
        ptArea = New Point(mea.X, mea.Y)
        ptNew = GetScalePoint(ptArea)
        Dim grfxCursor As Graphics = pnWorkarea.CreateGraphics()
        Dim RangeCuror As Integer = EraseSize * ZoomBase
        grfxCursor.FillRectangle(New SolidBrush(clrBack), ptArea.X, ptArea.Y, _
                                 EraseSize * ZoomBase, EraseSize * ZoomBase)
        grfxCursor.DrawRectangle(New Pen(Color.Black), ptArea.X, ptArea.Y, _
                                 EraseSize * ZoomBase, EraseSize * ZoomBase)
        grfxCursor.Dispose()
        'ø�s�{�ɹC��
        '�o�I��b���I�L��,�ثe .Net �� Cursor,�u��q *.cur �ɤ��إ�,�b�������,�Ȧ�����
        'Cursor �j�p����k�i�Ѩϥ�.��M�o���O�N��,�u��γo�خ��O�į�,�s�y�{�{����k�ӫ�
        '�߹C��.    
        If Not bTracking Then Return
    
        arrlstPoint.Add(ptNew)

        InsertCountinePoint()                  '���J�I�Ϩ�s��
        Dim br As Brush = New SolidBrush(clrBack)
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        Dim Rect() As Rectangle = PointToRect(apt, EraseSize)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.FillRectangles(br, Rect)
        grfxFore.Dispose()
        ptLast = ptNew

        '�o�إ� .Invalidate ���N .Refush �O�Ʊ�����ְ{�{.
    End Sub
    Private Sub EraseOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        bTracking = False

        '�ϥΪ̦欰�T�w,�s�Jø�Ϧ欰���X
        Dim GrapMethod As New GraphicsItem
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        If apt.Length = 1 Then
            With GrapMethod
                .DataType = "D"
                .Color = ClrBack
                .pWidth = EraseSize
                .Data.AddLines(apt)
            End With
        Else
            Dim Rect() As Rectangle = PointToRect(apt, EraseSize - 1)
            With GrapMethod
                .DataType = "P"
                .Color = ClrBack
                .pWidth = 1
                .Data.AddRectangles(Rect)
            End With
        End If

        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        pnWorkarea.Refresh()
        '�������ͤF�@�Өҥ~�{�H
        '�]���bGraphics.DrawLine ���w�q�� (0,0)->(0,1) ���ӬO�]�t�_�I�β��I��.
        '���O�b (0,0)->(0,0) �o�ر��p,�o�����L�Ħ欰
        '�G�n�S�O�B�z,�u���@�I�����p
    End Sub
    Private Sub EraseOnMouseLeave(ByVal obj As Object, ByVal ea As EventArgs)
        pnWorkarea.Refresh()   '�t�X�إߪ��{�ɹC�ХH�M���{�ɹC��
    End Sub

    Private Sub DrawPointOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
        arrlstPoint.Add(ptLast)
    End Sub
    Private Sub DrawPointOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        If Not bTracking Then Return
        ptNew = GetScalePoint(New Point(mea.X, mea.Y))
        arrlstPoint.Add(ptNew)

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim pn As New Pen(clrWork)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        grfxFore.DrawLines(pn, apt)
        grfxFore.Dispose()

        ptLast = ptNew
        'ø�ϰʧ@���̰򥻫��A,�p�G���H�Q�γo�ӵ{���ӰѦ�ø�ϧ@�k(�i���?�{���ޥ���b�����)
        '�o�جO�̲z�Q�����J�I
    End Sub
    Private Sub DrawPointOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        bTracking = False
      
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        If apt.Length = 1 Then
            GrapMethod.DataType = "D"
        Else
            GrapMethod.DataType = "P"
        End If
        With GrapMethod
            .Color = clrWork
            .pWidth = 1                  '�e���S��,�e�פ���
            .Data.AddLines(apt)
            .Data.StartFigure()
        End With

        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        pnWorkarea.Refresh()
        '�P EraseOnMouseUp ��]
    End Sub

    Private Sub DrawBrushOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptArea = New Point(mea.X - bmpOffest * ZoomBase, mea.Y - bmpOffest * ZoomBase)
        ptLast = GetScalePoint(ptArea)
        arrlstPoint.Add(ptLast)
        '�o�بϥΤ@�� Point() �M�@�ӥN�� Brush �� bitmap,�Ӫ�ܵ��ꪺø�Ϧ欰
        '�i�H�Q��,�o�ت�ܤ�k,�a�Ӫ��@�w�O�į�ΰO���骺���}�v�T.
        '���L�b .Net ���ꫬ�A�p���״I��,�٧䤣��@�ӾA�Ω󦹶��\�઺�覡.
        '���Q�p�e�a�������\��,���O�t�X��ɪ���k�ӻs�@��.
    End Sub
    Private Sub DrawBrushOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()

        SetBmpBrush(mea)        '�]�w�ϥΪ�Brush(�H�t���z���⪺Bitmap�إ�)
        
        Dim UpdateArea As New Rectangle(ptArea.X, ptArea.Y, bmpBrush.Width * ZoomBase + 1, bmpBrush.Height * ZoomBase + 1)
        pnWorkarea.Invalidate(UpdateArea)
        pnWorkarea.Update()
        ptArea = New Point(mea.X - bmpOffest * ZoomBase, mea.Y - bmpOffest * ZoomBase)
        ptNew = GetScalePoint(ptArea)

        Dim grfxCursor As Graphics = pnWorkarea.CreateGraphics()
        grfxCursor.DrawImage(bmpBrush, ptArea.X, ptArea.Y, bmpBrush.Width * ZoomBase, bmpBrush.Height * ZoomBase)
        grfxCursor.Dispose()
        'ø�s�C��,�z�ѦP DrawEraseOnMouseMove

        If Not bTracking Then Return

        InsertCountinePoint()                '�b���I�����J�I,�Ϩ�e�s�򪬺A

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())
        Dim i As Integer  
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        For i = 0 To apt.GetUpperBound(0)
            grfxFore.DrawImage(bmpBrush, apt(i).X, apt(i).Y, bmpBrush.Width, bmpBrush.Height)
        Next i
        grfxFore.Dispose()

        ptLast = ptNew
    End Sub
    Private Sub DrawBrushOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return
        bTracking = False

        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "I"
            .Image = bmpBrush
            .Data.AddLines(apt)
        End With
        GraphicsData.Add(GrapMethod)

        arrlstPoint.Clear()
    End Sub
    Private Sub DrawBrushOnMouseLeave(ByVal obj As Object, ByVal ea As EventArgs)
        pnWorkarea.Refresh()
    End Sub
    Private Sub SetBmpBrush(ByVal mea As MouseEventArgs)
        '�]�w���ꫬ�A

        Select Case (BrushType - 1) Mod 3
            Case 0
                bmpBrush = New Bitmap(8, 8)
                bmpOffest = 4
            Case 1
                bmpBrush = New Bitmap(5, 5)
                bmpOffest = 2
            Case 2
                bmpBrush = New Bitmap(2, 2)
                bmpOffest = 1
        End Select

        If mea.Button = MouseButtons.Right Then
            clrWork = clrBack
        Else
            clrWork = clrFore
        End If

        Dim grfx As Graphics = Graphics.FromImage(bmpBrush)
        Dim ClrNot As Color = Color.FromArgb(clrWork.A, Not clrWork.R, Not clrWork.G, Not clrWork.B)
        grfx.Clear(ClrNot)   '�ΤϦV��@���I��

        Select Case BrushType
            Case 1
                MyDrawEllipse(grfx, New Point(3, 3), 3, clrWork)
            Case 2
                MyDrawEllipse(grfx, New Point(2, 2), 2, clrWork)
            Case 3
                MyDrawEllipse(grfx, New Point(1, 1), 1, clrWork)
            Case 4
                grfx.FillRectangle(New SolidBrush(clrWork), 0, 0, 7, 7)
            Case 5
                grfx.FillRectangle(New SolidBrush(clrWork), 0, 0, 4, 4)
            Case 6
                grfx.DrawRectangle(New Pen(clrWork), 0, 0, 1, 1)
            Case 7
                grfx.DrawLine(New Pen(clrWork, 2), 7, 0, 0, 7)
            Case 8
                grfx.DrawLine(New Pen(clrWork, 2), 4, 0, 0, 4)
            Case 9
                grfx.DrawLine(New Pen(clrWork, 2), 1, 0, 0, 1)
            Case 10
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 7, 7)
            Case 11
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 4, 4)
            Case 12
                grfx.DrawLine(New Pen(clrWork, 2), 0, 0, 1, 1)
        End Select
        grfx.Dispose()
        bmpBrush.MakeTransparent(ClrNot)       '����I���z��
        '�o�إ��ӷQ�q PaintToolbarSub ��,�� bitmap �ǻ��L�ӧY�i,�����o�q�{���X�ݨ�.
        '��b���ǷM�©Mê��,���L���F��ƶǻ��W���@�P�ʦӨ�,�u�n�Ȯɽƻs�@��
        '�g�b�o�ا@����ʵ{�����c�ɪ��Ҽ{
    End Sub

    Private Sub DrawSprayOnMouseDown(ByVal obj As Object, ByVal mea As MouseEventArgs)
        ClearCommandRecord()
        clrWork = SelectUserColor(mea)
        bTracking = True

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))

        tmr = New Timer
        tmr.Interval = 20
        tmr.Enabled = True
        AddHandler tmr.Tick, AddressOf DrawSpray
        '�H�@�� Timer ����ӫإ߼Q��ĪG
    End Sub
    Private Sub DrawSprayOnMouseMove(ByVal obj As Object, ByVal mea As MouseEventArgs)
        SetMousePosition()
        If Not bTracking Then Return

        ptLast = GetScalePoint(New Point(mea.X, mea.Y))
    End Sub
    Private Sub DrawSprayOnMouseUp(ByVal obj As Object, ByVal mea As MouseEventArgs)
        If Not bTracking Then Return

        bTracking = False
        tmr.Dispose()    
        Dim apt() As Point = CType(arrlstPoint.ToArray(GetType(Point)), Point())

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "D"
            .Color = clrWork
            .pWidth = PenWidth
            .Data.AddLines(apt)
            .Data.StartFigure()
        End With
        GraphicsData.Add(GrapMethod)
        arrlstPoint.Clear()
        '�o�إH�@�� Point() ��ܼQ��欰
    End Sub
    Private Sub DrawSpray(ByVal obj As Object, ByVal ea As EventArgs)
        Dim RndAngle As Single
        Dim RndLength As Integer
        Dim RndPoint As Point
        Dim i As Integer

        Dim br As Brush = New SolidBrush(clrWork)
        Dim grfxFore As Graphics = GetGraphicsbyFore()
        For i = 1 To 2 * SprayWidth
            RndAngle = (2 * Math.PI * Rnd())
            RndLength = CInt((SprayWidth * Rnd()))
            RndPoint = New Point(ptLast.X + RndLength * Math.Cos(RndAngle), _
                                 ptLast.Y + RndLength * Math.Sin(RndAngle))
            Dim Rect As New Rectangle(RndPoint.X, RndPoint.Y, 1, 1)
            grfxFore.FillRectangle(br, Rect)
            arrlstPoint.Add(RndPoint)
        Next i
        grfxFore.Dispose()
        '�üƨ��I,�����D�P�p�e�a����k�O�_�@�P,������ĪG�ݨӤQ������.
    End Sub

    Protected Function SelectUserColor(ByVal mea As MouseEventArgs) As Color
        '�̨ϥΪ̫������C��
        Select Case mea.Button
            Case MouseButtons.Left
                Return clrFore
            Case MouseButtons.Right
                Return ClrBack
            Case Else
                Return clrFore
        End Select
    End Function
    Protected Sub ClearCommandRecord()
        '�M���Ѵ_��/���Щҳy�����h�lø�ϩR�O
        If UndoCount > 0 Then
            GraphicsData.RemoveRange(GraphicsData.Count - UndoCount, UndoCount)
            UndoCount = 0
        End If
    End Sub
    Protected Sub SetMousePosition()
        '��ܷƹ��y��
        sbpLocation.Text = pnWorkarea.PointToClient(MousePosition).X.ToString & "," & _
        pnWorkarea.PointToClient(MousePosition).Y.ToString
    End Sub
    Private Sub InsertCountinePoint()
        '�b�G�� Point ��,���J���w�ƪ� Point,�Ϩ䦨���@�ӥѳs���I�Ҳզ��� Point()
        If ptLast.Equals(ptNew) Then Return

        Dim ShiftPoint As Point = New Point(Math.Abs(ptLast.X - ptNew.X), Math.Abs(ptLast.Y - ptNew.Y))
        Dim OffestX, OffestY As Integer
        Dim SingleX, SingleY As Integer
        Dim InsPoint As Point

        If ptLast.X > ptNew.X Then SingleX = -1 Else SingleX = 1
        If ptLast.Y > ptNew.Y Then SingleY = -1 Else SingleY = 1

        If ShiftPoint.X > ShiftPoint.Y Then
            For OffestX = 1 To ShiftPoint.X Step 1
                OffestY = Math.Round((ShiftPoint.Y / ShiftPoint.X) * OffestX)
                InsPoint = New Point(ptLast.X + (OffestX * SingleX), ptLast.Y + (OffestY * SingleY))
                arrlstPoint.Add(InsPoint)
            Next OffestX
        Else
            For OffestY = 1 To ShiftPoint.Y Step 1
                OffestX = Math.Round((ShiftPoint.X / ShiftPoint.Y) * OffestY) * SingleX
                InsPoint = New Point(ptLast.X + OffestX, ptLast.Y + (OffestY * SingleY))
                arrlstPoint.Add(InsPoint)
            Next OffestY
        End If
        '�ھڪ��u��{���ײv�p��
    End Sub    
    Private Function PointToRect(ByVal apt() As Point, ByVal Range As Integer) As Rectangle()
        Dim Rect(apt.GetUpperBound(0)) As Rectangle

        Dim i As Integer
        For i = 0 To apt.GetUpperBound(0)
            Rect(i) = New Rectangle(apt(i).X, apt(i).Y, Range, Range)
        Next i
        Return Rect
        'Point() �ন Rectangle()
    End Function
    Protected Function GetScalePoint(ByVal pt As Point) As Point
        Return New Point(pt.X / ZoomBase, pt.Y / ZoomBase)
    End Function


End Class
