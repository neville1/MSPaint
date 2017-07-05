Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Public Class PaintWithMenuEdit
    '本類別定義主功能表編輯頁,及部份內容功能表元件的外觀及實作
    Inherits PaintWithMenuPrint

    Private miEditUndo, miEditRepeat As MenuItemHelp
    Private miEditCut, miEditCopy, miEditPaste As MenuItemHelp
    Private miEditDel, miEditSelectAll, miEditCopyTo, miEditPasteTo As MenuItemHelp
    '主功能表成員
    Private ConEditCut, ConEditCopy, ConEditPaste As MenuItemHelp
    Private ConEditDel, ConEDitSelectAll, ConEditCopyTo, ConEditPasteTo As MenuItemHelp
    '內容功能表成員

    Const strOpenFilter As String = _
  "點陣圖檔案(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
  "GIF 圖形交換格式(*.GIF)|*.gif|" & _
  "JPEG 圖形交換格式(*.JPG;,*.JPEG)|*.jpg;*.jpeg|" & _
  "所有圖片檔案|*.bmp;*.dib;*.ico;*.gif;*.jpg;*.jpeg|" & _
  "所有檔案|*.*"
    Const strSaveFilter As String = _
    "單色點陣圖(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "16 色點陣圖(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "256色點陣圖(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "24 位元點陣圖(*.BMP;,*.DIB)|*.bmp;*.dib|" & _
    "GIF 圖形交換格式(*.GIF)|*.gif|" & _
    "JPEG 圖形交換格式(*.JPG;,*.JPEG)|*.jpg;*.jpeg"

    Sub New()
        '主功能表編輯項建構元
        Dim mi As New MenuItem("編輯(&E)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1
        AddHandler mi.Popup, AddressOf MenuEditOnPopup

        miEditUndo = New MenuItemHelp("復原(&U)", sbpMenu)
        miEditUndo.Shortcut = Shortcut.CtrlZ
        miEditUndo.HelpText = "取消前次的動作。"
        Menu.MenuItems(index).MenuItems.Add(miEditUndo)
        AddHandler miEditUndo.Click, AddressOf MenuEditUndoandRepeatOnClick

        miEditRepeat = New MenuItemHelp("重複(&R)", sbpMenu)
        miEditRepeat.Shortcut = Shortcut.CtrlY
        miEditRepeat.HelpText = "重做前次取消的動作。"
        Menu.MenuItems(index).MenuItems.Add(miEditRepeat)
        AddHandler miEditRepeat.Click, AddressOf MenuEditUndoandRepeatOnClick

        Menu.MenuItems(index).MenuItems.Add("-")

        miEditCut = New MenuItemHelp("剪下(&T)", sbpMenu)
        miEditCut.Shortcut = Shortcut.CtrlX
        miEditCut.HelpText = "剪下選擇的項目，並放到剪貼簿中。"
        Menu.MenuItems(index).MenuItems.Add(miEditCut)
        AddHandler miEditCut.Click, AddressOf MenuEditCutOnClick

        miEditCopy = New MenuItemHelp("複製(&C)", sbpMenu)
        miEditCopy.Shortcut = Shortcut.CtrlC
        miEditCopy.HelpText = "複製選擇的項目，並放到剪貼簿中。"
        Menu.MenuItems(index).MenuItems.Add(miEditCopy)
        AddHandler miEditCopy.Click, AddressOf MenuEditCopyOnClick

        miEditPaste = New MenuItemHelp("貼上(&P)", sbpMenu)
        miEditPaste.Shortcut = Shortcut.CtrlV
        miEditPaste.HelpText = "插入剪貼簿的內容。"
        Menu.MenuItems(index).MenuItems.Add(miEditPaste)
        AddHandler miEditPaste.Click, AddressOf MenuEditPasteOnClick

        miEditDel = New MenuItemHelp("清除選擇的項目(&L)", sbpMenu)
        miEditDel.Shortcut = Shortcut.Del
        miEditDel.HelpText = "刪除選擇的部份"
        Menu.MenuItems(index).MenuItems.Add(miEditDel)
        AddHandler miEditDel.Click, AddressOf MenuEditDelOnClick

        miEditSelectAll = New MenuItemHelp("全選(&A)", sbpMenu)
        miEditSelectAll.Shortcut = Shortcut.CtrlA
        miEditSelectAll.HelpText = "選擇全部"
        Menu.MenuItems(index).MenuItems.Add(miEditSelectAll)
        AddHandler miEditSelectAll.Click, AddressOf MenuEditSelectAllOnClick

        Menu.MenuItems(index).MenuItems.Add("-")

        miEditCopyTo = New MenuItemHelp("複製到(&O)...", sbpMenu)
        miEditCopyTo.HelpText = "將所選的部份複製到檔案"
        Menu.MenuItems(index).MenuItems.Add(miEditCopyTo)
        AddHandler miEditCopyTo.Click, AddressOf MenuEditCopyToOnClick

        miEditPasteTo = New MenuItemHelp("貼上來源(&F)...", sbpMenu)
        miEditPasteTo.HelpText = "將檔案貼到所選的部份"
        Menu.MenuItems(index).MenuItems.Add(miEditPasteTo)
        AddHandler miEditPasteTo.Click, AddressOf MenuEditPasteToOnClick

        CreateContextMenu()

    End Sub
    Private Sub CreateContextMenu()
        '內容功能表建構元

        AddHandler ConMenuEdit.Popup, AddressOf ConMenuOnPopup

        ConEditCut = New MenuItemHelp("剪下(&T)", sbpMenu)
        ConEditCut.HelpText = "剪下選擇的項目，並放到剪貼簿中。"
        ConMenuEdit.MenuItems.Add(ConEditCut)
        AddHandler ConEditCut.Click, AddressOf MenuEditCutOnClick

        ConEditCopy = New MenuItemHelp("複製(&C)", sbpMenu)
        ConEditCopy.HelpText = "複製選擇的項目，並放到剪貼簿中。"
        ConMenuEdit.MenuItems.Add(ConEditCopy)
        AddHandler ConEditCopy.Click, AddressOf MenuEditCopyOnClick

        ConEditPaste = New MenuItemHelp("貼上(&P)", sbpMenu)
        ConEditPaste.HelpText = "插入剪貼簿的內容。"
        ConMenuEdit.MenuItems.Add(ConEditPaste)
        AddHandler ConEditPaste.Click, AddressOf MenuEditPasteOnClick

        ConEditDel = New MenuItemHelp("清除選擇部份(&L)", sbpMenu)
        ConEditDel.HelpText = "刪除選擇的部份"
        ConMenuEdit.MenuItems.Add(ConEditDel)
        AddHandler ConEditDel.Click, AddressOf MenuEditDelOnClick

        ConEDitSelectAll = New MenuItemHelp("全選(&A)", sbpMenu)
        ConEDitSelectAll.HelpText = "選擇全部"
        ConMenuEdit.MenuItems.Add(ConEDitSelectAll)
        AddHandler ConEDitSelectAll.Click, AddressOf MenuEditSelectAllOnClick

        ConMenuEdit.MenuItems.Add("-")

        ConEditCopyTo = New MenuItemHelp("複製到(&O)...", sbpMenu)
        ConEditCopyTo.HelpText = "將所選的部份複製到檔案"
        ConMenuEdit.MenuItems.Add(ConEditCopyTo)
        AddHandler ConEditCopyTo.Click, AddressOf MenuEditCopyToOnClick

        ConEditPasteTo = New MenuItemHelp("貼上來源(&F)...", sbpMenu)
        ConEditPasteTo.HelpText = "將檔案貼到所選的部份"
        ConMenuEdit.MenuItems.Add(ConEditPasteTo)
        AddHandler ConEditPasteTo.Click, AddressOf MenuEditPasteToOnClick

        ConMenuEdit.MenuItems.Add("-")
    End Sub
    Private Sub MenuEditOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        '主功能表編輯項,項目作用狀態
        If GraphicsData.Count = 0 OrElse GraphicsData.Count - UndoCount = 0 Then
            miEditUndo.Enabled = False
        Else
            miEditUndo.Enabled = True
        End If

        If UndoCount = 0 Then
            miEditRepeat.Enabled = False
        Else
            miEditRepeat.Enabled = True
        End If
        '其實這裏可以寫 miEditRepeat.Enabled = CBool(UndoCount)
        '但是個人還是認為這樣的寫法,看起來比較清楚明白

        miEditCut.Enabled = bEdit
        miEditCopy.Enabled = bEdit
        miEditDel.Enabled = bEdit
        miEditCopyTo.Enabled = bEdit

        Dim data As IDataObject = Clipboard.GetDataObject()
        miEditPaste.Enabled = data.GetDataPresent(GetType(Bitmap))
    End Sub
    Private Sub ConMenuOnPopup(ByVal obj As Object, ByVal ea As EventArgs)
        '內容功能表,項目作用狀態
        ConEditCut.Enabled = bEdit
        ConEditCopy.Enabled = bEdit
        ConEditDel.Enabled = bEdit
        ConEditCopyTo.Enabled = bEdit

        Dim data As IDataObject = Clipboard.GetDataObject()
        ConEditPaste.Enabled = data.GetDataPresent(GetType(Bitmap))
    End Sub

    Private Sub MenuEditUndoandRepeatOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '復原/重覆行為
        Dim MenuHelp As MenuItemHelp = DirectCast(obj, MenuItemHelp)
        Dim EumShift() As Integer = {1, -1}
        Dim Shift As Integer = EumShift(MenuHelp.Index())

        Dim CurrentCount As Integer
        If Shift > 0 Then
            CurrentCount = Math.Max(GraphicsData.Count - UndoCount - 1, 0)
        Else
            CurrentCount = Math.Min(GraphicsData.Count - UndoCount + 1, GraphicsData.Count - 1)
        End If
        Dim GrapItem As GraphicsItem = DirectCast(GraphicsData(CurrentCount), GraphicsItem)
        If GrapItem.DataType = "B" AndAlso Not GrapItem.IsSizeChange AndAlso Not GrapItem.IsFont Then
            UndoCount += (Shift * 2)
        Else
            UndoCount += Shift
        End If
        
        If GrapItem.IsSizeChange Then
            pnWorkarea.Size = GrapItem.SourceSize
            szBmp = New Size(GrapItem.SourceSize.Width / ZoomBase, GrapItem.SourceSize.Height / ZoomBase)
        End If

        pnWorkarea.Refresh()
        '這裏使用了一個 UndoCount +=(Shift * 2),當然是一個很差的作法
        '不過原著小畫家中則是,當復原一個由 EditPanel 所產生的更動時        
    End Sub

    Private Sub MenuEditCutOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Clipboard.SetDataObject(pnEditarea.BmpData, True)

        MenuEditDelOnClick(obj, ea)
    End Sub
    Private Sub MenuEditCopyOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Clipboard.SetDataObject(pnEditarea.BmpData, True)
    End Sub
    Private Sub MenuEditPasteOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        radbtnToolbar(1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(1), ea)         '移動工具列選項

        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        Dim data As IDataObject = Clipboard.GetDataObject()
        Dim bmp As Bitmap = DirectCast(data.GetData(GetType(Bitmap)), Bitmap)

        CheckDataSize(bmp)
    End Sub
    Private Sub MenuEditDelOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        pnEditarea.Dispose()
        bEdit = False
    End Sub
    Private Sub MenuEditSelectAllOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        radbtnToolbar(1).Checked = True
        ToolbarRadbtnSelect(radbtnToolbar(1), ea)          '移動工具列選項
        Dim rect As New Rectangle(0, 0, pnWorkarea.Width, pnWorkarea.Height)
        Dim path As New GraphicsPath
        path.AddRectangle(rect)

        CreateEditPanel(rect, path, False)
    End Sub
    Private Sub MenuEditCopyToOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Dim sfd As New SaveFileDialog
        sfd.Filter = strSaveFilter

        sfd.FilterIndex = 4
        If sfd.ShowDialog() = DialogResult.OK Then
            Dim FileName As String = sfd.FileName
            Select Case sfd.FilterIndex
                Case 0
                    pnEditarea.BmpData.Save(FileName)
                Case 1, 2, 3, 4
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Bmp)
                Case 5
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Gif)
                Case 6
                    pnEditarea.BmpData.Save(FileName, ImageFormat.Jpeg)
            End Select
        End If
    End Sub
    Private Sub MenuEditPasteToOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        Dim ofd As New OpenFileDialog
        ofd.Filter = strOpenFilter
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim bmp As New Bitmap(ofd.FileName)
            CheckDataSize(bmp)
        End If
    End Sub
    Private Sub CheckDataSize(ByVal bmp As Bitmap)
        '依 bitmap 大小對應處理方式
        Dim Rect As New Rectangle(0, 0, bmp.Width * ZoomBase, bmp.Height * ZoomBase)

        If Rect.Width >= pnWorkarea.Width OrElse Rect.Height >= pnWorkarea.Height Then
            Dim dr As DialogResult = _
            MessageBox.Show("[剪貼簿]的影像大於點陣圖。" & vbLf & "你要放大點陣圖嗎?", _
                            strProgName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case dr
                Case DialogResult.Yes
                    szBmp = bmp.Size
                    Dim tmp As New ControlDot
                    Dim mea As MouseEventArgs
                    WorkareaSizeChanged(tmp, mea)
                    tmp.Dispose()

                    CreateEditPanelWithBmp(bmp, False)

                Case DialogResult.No
                    Dim TmpBmp As New Bitmap(szBmp.Width, szBmp.Height)
                    Dim grfx As Graphics = Graphics.FromImage(TmpBmp)
                    grfx.DrawImage(bmp, Rect, Rect, GraphicsUnit.Pixel)
                    grfx.Dispose()

                    CreateEditPanelWithBmp(bmp, False)
                Case DialogResult.Cancel

            End Select
        Else
            CreateEditPanelWithBmp(bmp, False)
        End If
    End Sub
End Class
