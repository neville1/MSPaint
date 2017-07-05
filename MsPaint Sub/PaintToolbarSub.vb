Imports System
Imports System.Drawing
Imports System.Windows.Forms
Public Class PaintToolbarSub
    '這個類別負責處理,當使用者選擇工具列元件後,所對應的附屬元件建構及動作
    Inherits PaintToolBar

    Private ToolbarGrp As Panel
    Private lblToolbarSub() As Label
    Private MaxPixel As Integer = 0
    Private BmpTrans(1) As Bitmap
    Private BmpSpray(2) As Bitmap
    Private bmpPoint As Bitmap

    Protected bTransParent As Boolean = False   '透明色使用
    Protected PenWidth As Integer = 1           '畫筆寬度
    Protected UserType As Integer = 1           '幾何型式
    Protected EraseSize As Integer = 4          '橡皮擦大小
    Protected SprayWidth As Integer = 4         '噴鎗範圍
    Protected BrushType As Integer = 1          '筆刷型式
    ' Protected ZoomBase As Integer = 1           '畫面倍率  

    Sub New()
        bmpPoint = New Bitmap(1, 1)
        ToolbarGrp = New Panel
        With ToolbarGrp
            .Parent = pnToolbar
            .Location = New Point(8, 210)
            .Size = New Size(42, 68)
            .BorderStyle = BorderStyle.Fixed3D
        End With
        '這裏的程式碼定義,當使用者選擇工具列功能後,所需要的附屬資訊
    End Sub
    Protected Overrides Sub ToolbarRadbtnSelect(ByVal obj As Object, ByVal ea As EventArgs)
        '此處定義當使用者選擇了工具列元件後,所產生的附屬元件.
        '在某個角度看來,這也算是附屬元件的建構元,但由於它不是由建構元產生
        '故也可以視為動態元件.

        MyBase.ToolbarRadbtnSelect(obj, ea)

        DisposelblToolbarSub()                       '釋放附屬元件
        Dim i As Integer
        Select Case ToolbarCommand
            Case 1, 2, 10   '編輯,文字
                MaxPixel = 1
                ReDim lblToolbarSub(MaxPixel)
                Dim strImageName() As String = {"NoTransParent.bmp", "TransParent.bmp"}
                For i = 0 To MaxPixel
                    BmpTrans(i) = New Bitmap(Me.GetType(), strImageName(i))
                    BmpTrans(i).MakeTransparent(DefaultBackColor) '定義透明色,為光棒選擇所準備
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(39, 29)
                        .Location = New Point(0, 2 + i * 29)
                        .Image = BmpTrans(i)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf TransParentSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf TransParentSelectOnClick
                Next i

            Case 3        '橡皮擦
                MaxPixel = 3
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(14, 14)
                        .Location = New Point(10, 2 + i * 15)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf EraseSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf EraseSelectOnClick
                Next i

            Case 6       '放大鏡
                MaxPixel = 3
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 15)
                        .Location = New Point(3, 2 + i * 15)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf ZoomSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf ZoomSelectOnClick
                Next i

            Case 8       '筆刷
                MaxPixel = 11
                ReDim lblToolbarSub(MaxPixel)
                Dim cx() As Integer = {0, 12, 24, 0, 12, 24, 0, 12, 24, 0, 12, 24}
                Dim cy() As Integer = {0, 0, 0, 12, 12, 12, 24, 24, 24, 36, 36, 36}

                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(12, 12)
                        .Location = New Point(cx(i) + 1, cy(i) + (i + 1))
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf BrushSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf BrushSelectOnClick
                Next i

            Case 9       '噴鎗  
                MaxPixel = 2
                ReDim lblToolbarSub(MaxPixel)
                Dim SprayWidth() As Integer = {17, 22, 26}
                Dim SprayLocation() As Point = {New Point(0, 5), _
                                                New Point(17, 5), New Point(3, 29)}
                For i = 0 To MaxPixel
                    BmpSpray(i) = _
                    New Bitmap(Me.GetType(), "Spray" & (i + 1).ToString & ".bmp")
                    BmpSpray(i).MakeTransparent(DefaultBackColor)

                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(SprayWidth(i), 24)
                        .Location = SprayLocation(i)
                        .Image = BmpSpray(i)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf SprayWidthSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf SprayWidthSelectOnClick
                Next i

            Case 11, 12        '線條
                MaxPixel = 4
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 12)
                        .Location = New Point(2, 2 + i * 12)
                    End With

                    AddHandler lblToolbarSub(i).Paint, AddressOf LineSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf LineSelectOnClick
                Next i

            Case 13, 14, 15, 16    '幾何樣式
                MaxPixel = 2
                ReDim lblToolbarSub(MaxPixel)
                For i = 0 To MaxPixel
                    lblToolbarSub(i) = New Label
                    With lblToolbarSub(i)
                        .Name = (i + 1).ToString
                        .Parent = ToolbarGrp
                        .Size = New Size(34, 20)
                        .Location = New Point(2, 20 * i + 2)
                    End With
                    AddHandler lblToolbarSub(i).Paint, AddressOf RectSelectOnPaint
                    AddHandler lblToolbarSub(i).Click, AddressOf RectSelectOnClick
                Next i
            Case Else
        End Select
        '在16個工具列元件中,總共會產生8種不同的附屬元件(包含空白)
        '這裏的程式碼,有高度的重複性存在,但在經過考慮後.暫時不予於合併
        '畢竟這些附屬元件,也有許多不相同的地方.這裏合併的話,雖然可以大幅
        '減低程式碼,但也會讓閱讀程式碼,變的困難許多.

    End Sub
    '以下的七組方式,亦有很多相近之處,同樣為程式結構保留而不作合併動作,僅在特別加以說明.
    Private Sub TransParentSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim SubSelect As Integer
        If Not bTransParent Then SubSelect = 1 Else SubSelect = 2

        SetBackColor(SubSelect, CInt(lblPaint.Name))
        '這裏繪製元件的方式,有一點不同的是.表示使用者所選擇的光棒
        '是由OnPaint 的行為,用 BackColor 來表示,而不是由 OnClick 來建立
    End Sub
    Private Sub TransParentSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)

        If CInt(lblSelect.Name) = 1 Then
            bTransParent = False
        Else
            bTransParent = True
        End If
        ToolbarGrp.Refresh()        '這裏使用 Refresh.是為了觸發對應的 OnPaint 反映使用者的變更
    End Sub

    Private Sub EraseSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor((EraseSize - 2) \ 2, CInt(lblPaint.Name))

        Dim Range As Integer = CInt(lblPaint.Name) * 2 + 2
        Dim sy As Integer = lblPaint.Size.Height \ 2 - Range \ 2
        pea.Graphics.FillRectangle(New SolidBrush(clr), sy, sy, Range, Range)
    End Sub
    Private Sub EraseSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)

        EraseSize = CInt(lblSelect.Name) * 2 + 2
        ToolbarGrp.Refresh()
    End Sub

    Private Sub SprayWidthSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        SetBackColor(SprayWidth \ 4, CInt(lblPaint.Name))
        '這裏的顯示行為有不正確的地方,在原著中,顯示區的圖形應該會被反向處理.
    End Sub
    Private Sub SprayWidthSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        SprayWidth = Val(lblSelect.Name) * 4

        ToolbarGrp.Refresh()
    End Sub

    Private Sub ZoomSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim emuBase() As Integer = {1, 2, 6, 8}
        Dim SubSelect As Integer = Array.IndexOf(emuBase, ZoomBase) + 1
        Dim clr As Color = SetBackColor(SubSelect, CInt(lblPaint.Name))

        Dim Fnt As New Font("新細明體", 7)          '考慮有變更的可能
        pea.Graphics.DrawString(emuBase(CInt(lblPaint.Name) - 1).ToString & "x", _
                                 Fnt, New SolidBrush(clr), New PointF(5, 3))
    End Sub
    Private Sub ZoomSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        Dim emuBase() As Integer = {1, 2, 6, 8}
        If ZoomBase = emuBase(CInt(lblSelect.Name) - 1) Then Return '點選相同選項時

        ZoomBase = emuBase(CInt(lblSelect.Name) - 1)
        ToolbarGrp.Refresh()
        ZoomBaseChange()

        radbtnToolbar(LastCommand - 1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(LastCommand - 1), ea.Empty)  '選擇上一個命令
    End Sub

    Private Sub BrushSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)
        Dim clr As Color = SetBackColor(BrushType, CInt(lblPaint.Name))

        Select Case CInt(lblPaint.Name)
            Case 1
                MyDrawEllipse(pea.Graphics, New Point(5, 5), 4, clr)
            Case 2
                MyDrawEllipse(pea.Graphics, New Point(5, 5), 2, clr)
            Case 3
                pea.Graphics.FillRectangle(New SolidBrush(clr), New Rectangle(5, 5, 1, 1))
                '特例,這裏以一個大小為1的實心矩形,表示1個pixel的點
            Case 4
                pea.Graphics.FillRectangle(New SolidBrush(clr), New Rectangle(2, 2, 8, 8))
            Case 5
                pea.Graphics.FillEllipse(New SolidBrush(clr), New Rectangle(3, 3, 5, 5))
            Case 6
                pea.Graphics.DrawRectangle(New Pen(clr), New Rectangle(5, 5, 1, 1))
            Case 7
                pea.Graphics.DrawLine(New Pen(clr), New Point(9, 1), New Point(1, 9))
            Case 8
                pea.Graphics.DrawLine(New Pen(clr), New Point(7, 3), New Point(3, 7))
            Case 9
                pea.Graphics.DrawLine(New Pen(clr), New Point(6, 4), New Point(4, 6))
            Case 10
                pea.Graphics.DrawLine(New Pen(clr), New Point(1, 1), New Point(9, 9))
            Case 11
                pea.Graphics.DrawLine(New Pen(clr), New Point(3, 3), New Point(7, 7))
            Case 12
                pea.Graphics.DrawLine(New Pen(clr), New Point(4, 4), New Point(6, 6))
        End Select
    End Sub
    Private Sub BrushSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        BrushType = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Sub LineSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor(PenWidth, CInt(lblPaint.Name))

        Dim pn As New Pen(clr, CInt(lblPaint.Name))
        Dim sy As Integer = lblPaint.Size.Height \ 2 - (pn.Width \ 2) + 1  '計算座標偏移
        pea.Graphics.DrawLine(pn, 4, sy, 28, sy)
    End Sub
    Private Sub LineSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        PenWidth = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Sub RectSelectOnPaint(ByVal obj As Object, ByVal pea As PaintEventArgs)
        Dim lblPaint As Label = DirectCast(obj, Label)

        Dim clr As Color = SetBackColor(UserType, CInt(lblPaint.Name))

        Select Case CInt(lblPaint.Name)
            Case 1
                pea.Graphics.DrawRectangle(New Pen(clr), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
            Case 2
                pea.Graphics.DrawRectangle(New Pen(clr), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
                pea.Graphics.FillRectangle(New SolidBrush(Color.Gray), 5, 5, lblPaint.Width - 11, lblPaint.Height - 11)
            Case 3
                pea.Graphics.FillRectangle(New SolidBrush(Color.Gray), 4, 4, lblPaint.Width - 10, lblPaint.Height - 10)
        End Select
    End Sub
    Private Sub RectSelectOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim lblSelect As Label = DirectCast(obj, Label)
        UserType = CInt(lblSelect.Name)

        ToolbarGrp.Refresh()
    End Sub

    Private Function SetBackColor(ByVal UserSelect As Integer, ByVal objName As Integer) As Color
        '這裏以 BackColor 來製作使用者選擇物件後的光棒,並以 clr 模擬反相色的效果
        Dim clr As Color
        If UserSelect = objName Then
            lblToolbarSub(objName - 1).BackColor = Color.DarkBlue
            clr = Color.White
        Else
            lblToolbarSub(objName - 1).BackColor = ToolbarGrp.BackColor
            clr = Color.Black
        End If
        Return clr
    End Function
    Private Sub DisposelblToolbarSub()
        '釋放所有附屬元件
        Dim i As Integer
        If MaxPixel = 0 Then Return
        For i = 0 To MaxPixel
            lblToolbarSub(i).Dispose()
        Next i
        MaxPixel = 0
    End Sub

    Protected Sub MyDrawEllipse(ByVal grfx As Graphics, ByVal center As Point, _
                               ByVal radius As Integer, ByVal clr As Color)
        Dim i As Double
        Dim last As Point        
        Dim pt(360) As Point
        Dim Count As Integer = 0
        For i = 0 To 2 * Math.PI Step 2 * Math.PI / 360
            last.X = center.X + CInt(radius * Math.Cos(i))
            last.Y = center.Y + CInt(radius * Math.Sin(i))
            pt(Count) = last           
            Count += 1
        Next i
        grfx.FillClosedCurve(New SolidBrush(clr), pt)
        '為什麼要自己用點畫圓呢?,這是因為有個很奇怪的理由
        'Graphics.DrawEllipse 這個命令,不知為何,在給予的矩形空間很小時,畫出來的圖
        '形,一點都不圓...><至於為什麼不直接在迴圈中畫線即可,主要是對Graphics 的動作
        '有些效能上的疑慮,不希望多次呼叫 Graphics
    End Sub
    
End Class
