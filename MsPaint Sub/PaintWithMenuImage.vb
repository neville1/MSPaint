Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Public Class PaintWithMenuImage
    '本類別定義主功能表影像類及部份內容功能表外觀及實作
    Inherits PaintWithMenuView

    Private miImageTransParent, miImageClear As MenuItemHelp
    Sub New()
        '主功能表影像類建構元
        Dim mi As New MenuItem("影像(&I)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuImageOnPopup

        Dim miImageTurn As New MenuItemHelp("翻轉/旋轉(&F)...")
        miImageTurn.Shortcut = Shortcut.CtrlR
        miImageTurn.HelpPanel = sbpMenu
        miImageTurn.HelpText = "翻轉或旋轉圖片或選擇的部份。"
        Menu.MenuItems(index).MenuItems.Add(miImageTurn)
        AddHandler miImageTurn.Click, AddressOf MenuImageTurnOnClick

        Dim miImageExtend As New MenuItemHelp("延伸/扭曲(&S)...")
        AddHandler miImageExtend.Click, AddressOf MenuImageExtendOnClick
        miImageExtend.Shortcut = Shortcut.CtrlW
        miImageExtend.HelpPanel = sbpMenu
        miImageExtend.HelpText = "拉長或扭曲圖片或選擇的部份。"
        Menu.MenuItems(index).MenuItems.Add(miImageExtend)

        Dim miImageColorXor As New MenuItemHelp("色彩對換(&I)")
        AddHandler miImageColorXor.Click, AddressOf MenuImageColorNotOnClick
        miImageColorXor.Shortcut = Shortcut.CtrlI
        miImageColorXor.HelpPanel = sbpMenu
        miImageColorXor.HelpText = "將圖片或選擇的部份顏色對調。"
        Menu.MenuItems(index).MenuItems.Add(miImageColorXor)

        Dim miImageProperty As New MenuItemHelp("屬性(&A)...")
        AddHandler miImageProperty.Click, AddressOf MenuImagePropertyOnClick
        miImageProperty.Shortcut = Shortcut.CtrlE
        miImageProperty.HelpPanel = sbpMenu
        miImageProperty.HelpText = "變更圖片的屬性。"
        Menu.MenuItems(index).MenuItems.Add(miImageProperty)

        miImageClear = New MenuItemHelp("清除影像(&C)")
        AddHandler miImageClear.Click, AddressOf MenuImageClearOnClick
        miImageClear.Shortcut = Shortcut.CtrlShiftN
        miImageClear.HelpPanel = sbpMenu
        miImageClear.HelpText = "清除圖片或所選的部份。"
        Menu.MenuItems(index).MenuItems.Add(miImageClear)

        miImageTransParent = New MenuItemHelp("不透明處理(&D)")
        AddHandler miImageTransParent.Click, AddressOf MenuImageTransParentOnClick
        With miImageTransParent
            .HelpPanel = sbpMenu
            .HelpText = "將目前選擇的部份變成透明或不透明。"
        End With
        Menu.MenuItems(index).MenuItems.Add(miImageTransParent)

        CreateContextMenu()             '建構內容功能表

    End Sub
    Private Sub CreateContextMenu()
        '內容功能表建構元
        Dim ConImageTurn As New MenuItemHelp("翻轉/旋轉(&F)...", sbpMenu)
        ConImageTurn.HelpText = "翻轉或旋轉圖片或選擇的部份。"
        ConMenuEdit.MenuItems.Add(ConImageTurn)
        AddHandler ConImageTurn.Click, AddressOf MenuImageTurnOnClick

        Dim ConImageExtend As New MenuItemHelp("延伸/扭曲(&S)...", sbpMenu)
        ConImageExtend.HelpText = "拉長或扭曲圖片或選擇的部份。"
        ConMenuEdit.MenuItems.Add(ConImageExtend)
        AddHandler ConImageExtend.Click, AddressOf MenuImageExtendOnClick

        Dim ConImageColorXor As New MenuItemHelp("色彩對換(&I)", sbpMenu)
        ConImageColorXor.HelpText = "將圖片或選擇的部份顏色對調。"
        ConMenuEdit.MenuItems.Add(ConImageColorXor)
        AddHandler ConImageColorXor.Click, AddressOf MenuImageColorNotOnClick
    End Sub
    Private Sub MenuImageOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        miImageTransParent.Checked = Not bTransParent

        miImageClear.Enabled = Not bEdit
    End Sub
    Private Sub MenuImageTurnOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '翻轉
        If bText Then TextAreaComplete()
        Dim dlg As New BitmapTurnDialog

        If dlg.ShowDialog() <> DialogResult.OK Then Return

        Dim bmp As Bitmap
        If bEdit Then
            bmp = pnEditarea.BmpData
        Else
            If bText Then TextAreaComplete()
            bmp = WorkareaChangeBmp()
        End If

        Select Case dlg.SelectCommand
            Case 1
                bmp.RotateFlip(RotateFlipType.RotateNoneFlipX)
            Case 2
                bmp.RotateFlip(RotateFlipType.RotateNoneFlipY)
            Case 3, 4
                bmp.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case 5
                bmp.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case 6
                bmp.RotateFlip(RotateFlipType.Rotate270FlipNone)
        End Select

        If bEdit Then
            Dim ptSt As Point = pnEditarea.Location
            Dim bTrans As Boolean = pnEditarea.FormTransParent
            pnEditarea.Dispose()
            CreateEditPanelWithBmp(bmp, bTrans)
            pnEditarea.Location = ptSt
        Else            
            WorkareaAddGraphicsItem(bmp)
        End If
        dlg.Dispose()
        '關於翻轉,延伸及色彩轉換都有一個很大的問題,理想的狀況下,程式應該要能視編輯區域及
        '工作區為相同的處理方式.可見我在設計理念上,還有許多需要知道的知識.
    End Sub
    Private Sub MenuImageExtendOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '延伸扭曲
        If bText Then TextAreaComplete()

        Dim dlg As New BitmapExtendDialog

        If dlg.ShowDialog <> DialogResult.OK Then Return

        If dlg.IsExtend Then
            If bEdit Then
                pnEditarea.Size = New Size(pnEditarea.Width * dlg.ExtendX \ 100, _
                                           pnEditarea.Height * dlg.ExtendY \ 100)
            Else
                If bText Then TextAreaComplete()
                Dim nsz As New Size(pnWorkarea.Width * dlg.ExtendX \ 100, _
                                    pnWorkarea.Height * dlg.ExtendY \ 100)
                Dim bmpExtEnd As New Bitmap(nsz.Width, nsz.Height)
                Dim grfxExtEnd As Graphics = Graphics.FromImage(bmpExtEnd)
                Dim rectSrc As New Rectangle(0, 0, szBmp.Width, szBmp.Height)
                Dim rectDst As New Rectangle(0, 0, nsz.Width, nsz.Height)
                grfxExtEnd.DrawImage(WorkareaChangeBmp(), rectDst, rectSrc, GraphicsUnit.Pixel)
                grfxExtEnd.Dispose()

                WorkareaAddGraphicsItem(bmpExtEnd)
            End If
        End If

        If dlg.IsDistort Then
            If bEdit Then
                Dim bmpback = BitmapDistort(pnEditarea.BmpData, New Point(dlg.DistortX, dlg.DistortY))
                Dim ptSt As Point = pnEditarea.Location
                Dim bTrans As Boolean = pnEditarea.FormTransParent
                pnEditarea.Dispose()

                CreateEditPanelWithBmp(bmpback, bTrans)
                pnEditarea.Location = ptSt
            Else
                Dim bmpback = BitmapDistort(WorkareaChangeBmp(), New Point(dlg.DistortX, dlg.DistortY))

                WorkareaAddGraphicsItem(bmpback)
            End If
        End If
    End Sub
    Private Function BitmapDistort(ByVal bmp As Bitmap, ByVal Distort As Point) As Bitmap
        Dim WorkBmp As Bitmap = bmp
        If Distort.X <> 0 Then
            Dim rDistortX As Double = (Distort.X) * (Math.PI / 180)
            Dim ExtX As Double = WorkBmp.Height * Math.Abs(Math.Tan(rDistortX))
            Dim nsz As New Size(WorkBmp.Width + CInt(ExtX), WorkBmp.Height)

            Dim bmpBack As New Bitmap(nsz.Width, nsz.Height)
            Dim grfxback As Graphics = Graphics.FromImage(bmpback)
            grfxback.Clear(clrBack)
            Dim Marx As New Matrix(1, 0, -(Math.Tan(rDistortX)), 1, 0, 0)
            grfxback.Transform = Marx

            Dim startX As Integer
            If Distort.X > 0 Then startX = CInt(ExtX) Else startX = 0

            grfxback.DrawImage(WorkBmp, startX, 0, WorkBmp.Width, WorkBmp.Height)
            grfxback.Dispose()
            WorkBmp = bmpBack
        End If

        If Distort.Y <> 0 Then
            Dim rDistortY As Double = (Distort.Y) * (Math.PI / 180)
            Dim ExtY As Double = WorkBmp.Width * Math.Abs(Math.Tan(rDistortY))
            Dim nsz As New Size(WorkBmp.Width, WorkBmp.Height + CInt(ExtY))

            Dim bmpBack As New Bitmap(nsz.Width, nsz.Height)
            Dim grfxback As Graphics = Graphics.FromImage(bmpback)
            grfxback.Clear(clrBack)
            Dim Marx As New Matrix(1, -(Math.Tan(rDistortY)), 0, 1, 0, 0)
            grfxback.Transform = Marx

            Dim startY As Integer
            If Distort.Y > 0 Then startY = CInt(ExtY) Else startY = 0
            grfxback.DrawImage(WorkBmp, 0, startY, WorkBmp.Width, WorkBmp.Height)
            grfxback.Dispose()
            WorkBmp = bmpBack
        End If
        '這段圖形扭曲的程式碼,最大的問題在於我不了解當圖形X,Y軸同時扭曲時,對應的運算基礎
        '為何? 做出這段效能低落的程式碼.

        Return WorkBmp
    End Function
    Protected Overrides Sub WorkareaAddGraphicsItem(ByVal bmp As Bitmap)
        '將圖形變更結果視作繪圖行為加入
        ClearCommandRecord()

        Dim GrapMethod As New GraphicsItem
        With GrapMethod
            .DataType = "B"
            .Image = bmp
            .ImageLocation = New Point(0, 0)
            .ImageSize = New Size(bmp.Width, bmp.Height)
            .IsSizeChange = True
            .SourceSize = New Size(pnWorkarea.Width \ ZoomBase, pnWorkarea.Height \ ZoomBase)

        End With
        szBmp = New Size(bmp.Width, bmp.Height)
        pnWorkarea.Size = New Size(szbmp.Width * ZoomBase, szbmp.Height * ZoomBase)
        GraphicsData.Add(GrapMethod)
        pnWorkarea.Refresh()
        '為了建立均一性的 GrapHicsData.及復原的考量,因而損失了大量的記憶體.
        '應該會有更好的方法.
    End Sub
    Private Sub MenuImageColorNotOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '色彩對換
        If bText Then TextAreaComplete()

        Dim bmpback As Bitmap
        If bEdit Then
            bmpback = pnEditarea.BmpData
        Else
            bmpback = WorkareaChangeBmp()
        End If

        Dim clrTmp As Color
        Dim x, y As Integer
        Dim a, r, b, g As Byte
        For y = 0 To bmpback.Height - 1
            For x = 0 To bmpback.Width - 1
                clrTmp = bmpback.GetPixel(x, y)
                r = Not clrTmp.R
                b = Not clrTmp.B
                g = Not clrTmp.G
                bmpback.SetPixel(x, y, clrTmp.FromArgb(clrTmp.A, r, b, g))
            Next x
        Next y

        If bEdit Then
            Dim ptSt As Point = pnEditarea.Location
            Dim bTrans As Boolean = pnEditarea.FormTransParent
            pnEditarea.Dispose()


            CreateEditPanelWithBmp(bmpback, bTrans)
            pnEditarea.Location = ptSt
        Else
            WorkareaAddGraphicsItem(bmpback)
        End If

        '這段程式碼相當簡單,卻存在相當大的問題
        '以Pixel 為單位,一個一個的改色,效率已經差到不能接受的地步.
    End Sub
    Private Sub MenuImagePropertyOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '圖形屬性
        Dim dlg As New PropertyDialog(strFileName, SzBmp)
        
        If dlg.ShowDialog <> DialogResult.OK Then Return
        If Not dlg.BmpSize.Equals(szBmp) Then
            szBmp = dlg.BmpSize
            ZoomBase = 1
            Dim mea As MouseEventArgs
            WorkAreaSizeChanged(obj, mea)
        End If
    End Sub
    Private Sub MenuImageClearOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '清除圖形
        Dim GrapMethod As New GraphicsItem
        Dim Rect As New Rectangle(0, 0, pnWorkarea.Width, pnWorkarea.Height)

        With GrapMethod
            .DataType = "F"
            GrapMethod.FillColor = clrBack
            GrapMethod.Data.AddRectangle(Rect)
        End With
        GraphicsData.Add(GrapMethod)

        pnWorkarea.Refresh()
        '這裏有個奇怪的現象,在小畫家原著中,只有在沒有編輯區的情況下,這個指令才能
        '工作.不知道原著為何如此設計,就如此吧.
    End Sub
    Private Sub MenuImageTransParentOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '使用透明色
        miImageTransParent.Checked = Not miImageTransParent.Checked
        bTransParent = Not bTransParent
        pnToolbar.Refresh()

        If bEdit Then pnEditarea.TransParent = bTransParent
    End Sub
    Protected Overrides Sub WorkareaSizeChanged(ByVal obj As Object, ByVal mea As MouseEventArgs)
        '繪圖工作區大小變更
        If bText Then TextAreaComplete()

        Dim BmpTmp As New Bitmap(szBmp.Width, szBmp.Height)
        Dim grfxTmpBmp As Graphics = Graphics.FromImage(BmpTmp)
        grfxTmpBmp.Clear(clrBack)

        Dim Bmp As Bitmap = WorkareaChangeBmp()
        grfxTmpBmp.DrawImage(Bmp, 0, 0, Bmp.Width, Bmp.Height)
        grfxTmpBmp.Dispose()

        WorkareaAddGraphicsItem(BmpTmp)
        sbpRectangle.Text = ""
        pnWorkarea.Parent.Refresh()
        '這裏與延伸不同,繪圖工作區大小變更的行為,被視為一種裁切或加大.
    End Sub
End Class
