Imports Microsoft.Win32
Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

Public Class PaintWithMenuFile
    Inherits PaintCommandWithOther
    '本類別定義關於主功能表中關於檔案新增,開啟,存檔及附屬功能外觀建構及方法處理

    Protected strFileName As String = "未命名"

    Private fileColorByte As Integer
    Private strTempFileName(3) As String
    Private miFileTmpName(3) As MenuItemHelp
    Private TempFileIndex As Integer = 0
    Private miFileTmpline As MenuItem    
    Const strFilePath = "TempFilePath"
    Const strTempFileIndex = "TempFileIndex"
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
        '主功能表檔案項建構元
        Text = strProgName
        Menu = New MainMenu

        Dim mi As New MenuItem("檔案(&F)")
        Menu.MenuItems.Add(mi)
        Dim index As Integer = Menu.MenuItems.Count - 1

        Dim miFileNew As New MenuItemHelp("開新檔案(&N)")
        AddHandler miFileNew.Click, AddressOf MenuFileNewOnClick
        miFileNew.Shortcut = Shortcut.CtrlN
        miFileNew.HelpPanel = sbpMenu
        miFileNew.HelpText = "建立新的文件。"
        Menu.MenuItems(index).MenuItems.Add(miFileNew)

        Dim miFileOpen As New MenuItemHelp("開啟舊檔(&O)...")
        AddHandler miFileOpen.Click, AddressOf MenuFileOpenOnClick
        miFileOpen.Shortcut = Shortcut.CtrlO
        miFileOpen.HelpPanel = sbpMenu
        miFileOpen.HelpText = "開啟現有的文件。"
        Menu.MenuItems(index).MenuItems.Add(miFileOpen)

        Dim miFileSave As New MenuItemHelp("存檔(&S)")
        AddHandler miFileSave.Click, AddressOf MenuFileSaveOnClick
        miFileSave.Shortcut = Shortcut.CtrlS
        miFileSave.HelpPanel = sbpMenu
        miFileSave.HelpText = "儲存使用中的文件。"
        Menu.MenuItems(index).MenuItems.Add(miFileSave)

        Dim miFileSaveAs As New MenuItemHelp("另存新檔(&A)...")
        AddHandler miFileSaveAs.Click, AddressOf MenuFileSaveAsOnClick
        miFileSaveAs.HelpPanel = sbpMenu
        miFileSaveAs.HelpText = "以新名稱儲存使用中的文件。"
        Menu.MenuItems(index).MenuItems.Add(miFileSaveAs)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim miFileReviewPrint As New MenuItemHelp("預覽列印(&V)")
        AddHandler miFileReviewPrint.Click, AddressOf MenuFileReviewPrintOnClick
        miFileReviewPrint.HelpPanel = sbpmenu
        miFileReviewPrint.HelpText = "顯示全頁。"
        Menu.MenuItems(index).MenuItems.Add(miFileReviewPrint)

        Dim miFileSetPrint As New MenuItemHelp("設定列印格式(&U)...")
        AddHandler miFileSetPrint.Click, AddressOf MenuFileSetPrintOnClick
        miFileSetPrint.HelpPanel = sbpmenu
        miFileSetPrint.HelpText = "變更版面配置。"
        Menu.MenuItems(index).MenuItems.Add(miFileSetPrint)

        Dim miFilePrint As New MenuItemHelp("列印(&P)...")
        AddHandler miFilePrint.Click, AddressOf MenuFilePrintOnClick
        miFilePrint.Shortcut = Shortcut.CtrlP
        miFilePrint.HelpPanel = sbpmenu
        miFilePrint.HelpText = "列印使用中的文件並設定列印選項。"
        Menu.MenuItems(index).MenuItems.Add(miFilePrint)

        Menu.MenuItems(index).MenuItems.Add("-")

        Dim i As Integer = 3
        For i = 0 To 3
            miFileTmpName(i) = New MenuItemHelp
            miFileTmpName(i).HelpPanel = sbpmenu
            miFileTmpName(i).HelpText = "開啟這件文件"
            miFileTmpName(i).Visible = False
            Menu.MenuItems(index).MenuItems.Add(miFileTmpName(i))
            AddHandler miFileTmpName(i).Click, AddressOf MenuFileOpenTempOnClick
        Next i

        miFileTmpline = New MenuItem("-")
        miFileTmpline.Visible = False
        Menu.MenuItems(index).MenuItems.Add(miFileTmpline)

        Dim miFileExit As New MenuItemHelp("結束(&X)")
        AddHandler miFileExit.Click, AddressOf MenuFileExitOnClick
        miFileExit.Shortcut = Shortcut.AltF4
        miFileExit.HelpPanel = sbpmenu
        miFileExit.HelpText = "結束小畫家重製版。"
        Menu.MenuItems(index).MenuItems.Add(miFileExit)

        MakeCaption()
        AddHandler Me.MenuComplete, AddressOf LoseFocusMenu
    End Sub
    Private Sub MenuFileNewOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '開新檔案行為實作
        If Not OktoTrash() Then Return

        ClearGraphicsData()

        bmpPic = New Bitmap(szBmp.Width, szBmp.Height)
        Dim grfxPic As Graphics = Graphics.FromImage(bmpPic)
        grfxPic.Clear(ClrBack)
        grfxPic.Dispose()

        pnWorkarea.Size = bmpPic.Size       
        strFileName = "未命名"
        MakeCaption()
        pnWorkarea.Refresh()
        '色彩定義
    End Sub
    Private Sub MenuFileOpenOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If Not OktoTrash() Then Return

        Dim ofd As New OpenFileDialog
        ofd.Filter = strOpenFilter
        If ofd.ShowDialog() = DialogResult.OK Then LoadFile(ofd.FileName)
    End Sub
    Private Sub MenuFileSaveOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        If strFileName = "未命名" OrElse strFileName.Length = 0 Then
            SaveFileDlg()
        Else
            SaveFile(0)
        End If
    End Sub
    Private Sub MenuFileSaveAsOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        SaveFileDlg()
    End Sub
    Protected Overridable Sub MenuFileReviewPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '由 PaintWithMenuPrint.vb 覆寫
    End Sub
    Protected Overridable Sub MenuFileSetPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
    End Sub
    Protected Overridable Sub MenuFilePrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
    End Sub
    Private Sub MenuFileExitOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        Me.Close()
    End Sub
    Private Sub MenuFileOpenTempOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '開啟最近三個檔案
        If Not OktoTrash() Then Return

        Dim MenuFileIndex As MenuItem = DirectCast(obj, MenuItem)
        Dim FileIndex As Integer
        If MenuFileIndex.Text = Nothing Then Return
        FileIndex = CInt(MenuFileIndex.Text.Substring(1, 1)) - 1
        LoadFile(strTempFileName(FileIndex))
    End Sub
    Protected Overrides Sub Onload(ByVal ea As EventArgs)
        '以命令列方式讀入檔案
        MyBase.OnLoad(ea)

        Dim astrArgs As String() = Environment.GetCommandLineArgs()
        If astrArgs.Length > 1 Then
            If File.Exists(astrArgs(1)) Then
                LoadFile(astrArgs(1))
            Else
                Dim dr As DialogResult = _
                MessageBox.Show("找不到檔案: " & Path.GetFileName(astrArgs(1)) & vbLf _
                & vbLf & "要建立新檔案?", Text, MessageBoxButtons.YesNoCancel)
                Select Case dr
                    Case DialogResult.Yes
                        strFileName = astrArgs(1)
                        File.Create(strFileName).Close()
                        MakeCaption()
                    Case DialogResult.No

                    Case DialogResult.Cancel
                        Close()
                End Select
            End If
        Else

        End If
    End Sub
    Protected Sub LoadFile(ByVal strFileName As String)
        '由檔案讀取點陣圖
        Dim img As Image
        Try
            img = Image.FromFile(strFileName)
        Catch ex As OutOfMemoryException
            MessageBox.Show(strFileName & vbLf & "[小畫家重製版]無法讀取這個檔案。" & vbLf & _
            "這不是正確的點陣圖檔案，或不已經不再支援這種格式。", Text, MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
            Return
        End Try

        ClearGraphicsData()

        bmpPic = New Bitmap(img)
        SzBmp = bmpPic.Size
        pnWorkarea.Size = SzBmp   
        img.Dispose()
        Me.strFileName = strFileName
        ShowTempFilesMenu()
        MakeCaption()

        pnWorkarea.Refresh()
    End Sub
    Private Sub SaveFile(ByVal SaveFormat As Integer)
        '寫入點陣圖檔案
        bmpPic = New Bitmap(WorkareaChangeBmp())

        Select Case SaveFormat
            Case 0
                bmpPic.Save(strFileName)
            Case 1, 2, 3, 4
                bmpPic.Save(strFileName, ImageFormat.Bmp)
            Case 5
                bmpPic.Save(strFileName, ImageFormat.Gif)
            Case 6
                bmpPic.Save(strFileName, ImageFormat.Jpeg)
        End Select

        GraphicsData.Clear()
        UndoCount = 0
        pnWorkarea.Refresh()
    End Sub
    Private Function SaveFileDlg() As Boolean
        'SaveFile 對話方塊處理
        Dim sfd As New SaveFileDialog
        If strFileName.Length > 1 Then
            sfd.FileName = Path.GetFileNameWithoutExtension(strFileName)
        Else
            sfd.FileName = ""
        End If

        sfd.Filter = strSaveFilter
        Select Case Path.GetExtension(strFileName).ToLower()
            Case ".bmp", ".dib"
                Select Case fileColorByte
                    Case 1
                        sfd.FilterIndex = 1
                    Case 4
                        sfd.FilterIndex = 2
                    Case 8
                        sfd.FilterIndex = 3
                    Case 24
                        sfd.FilterIndex = 4
                End Select
            Case ".gif"
                sfd.FilterIndex = 5
            Case ".jpeg", ".jpg"
                sfd.FilterIndex = 6
            Case Else
                sfd.FilterIndex = 4
        End Select

        If sfd.ShowDialog() = DialogResult.OK Then
            strFileName = sfd.FileName
            SaveFile(sfd.FilterIndex)
            MakeCaption()
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub MakeCaption()
        '設定 Form 標題
        Text = FileTitle() & " - " & strProgName
    End Sub
    Protected Function FileTitle() As String
        Return Path.GetFileName(strFileName)
    End Function
    Private Sub ShowTempFilesMenu()
        Dim i As Integer
        If TempFileIndex > 0 Then
            For i = Math.Min(TempFileIndex, 3) To 1 Step -1
                strTempFileName(i) = strTempFileName(i - 1)
            Next i
        End If
        TempFileIndex = Math.Min(TempFileIndex + 1, 4)
        strTempFileName(0) = strFileName
        miFileTmpline.Visible = True
        For i = 0 To TempFileIndex - 1
            miFileTmpName(i).Text = "&" & (i + 1).ToString & " " & Path.GetFileName(strTempFileName(i))
            miFileTmpName(i).Visible = True
        Next i
        '這裏模擬表示,最近三個開啟的檔案,這應該與標準作法不同
    End Sub
    Private Function OktoTrash() As Boolean
        '判斷圖形是否修改
        If bEdit Then DisposeEditPanel()
        If bText Then TextAreaComplete()

        If GraphicsData.Count = 0 Then Return True

        Dim dr As DialogResult = MessageBox.Show("儲存修改到 " & FileTitle() & "?" _
        , strProgName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)

        Select Case dr
            Case DialogResult.Yes : Return SaveFileDlg()
            Case DialogResult.No : Return True
            Case DialogResult.Cancel : Return False
        End Select
        Return False
    End Function
    Private Sub ClearGraphicsData()
        '清除繪圖動作,設定為起始狀態
        GraphicsData.Clear()
        UndoCount = 0
        ZoomBase = 1

        clrFore = Color.Black
        clrBack = Color.White
        btnColorFore.BackColor = clrFore
        btnColorBack.BackColor = clrBack
    End Sub
    Protected Overrides Sub SaveRegistryInfo(ByVal regkey As RegistryKey)
        '將最近開啟的檔案位置,寫入 Registry
        MyBase.SaveRegistryInfo(regkey)

        Dim i As Integer
        regkey.SetValue(strTempFileIndex, TempFileIndex)

        For i = 1 To TempFileIndex
            regkey.SetValue(strFilePath & i.ToString, strTempFileName(i - 1))
        Next i
    End Sub
    Protected Overrides Sub LoadRegistryInfo(ByVal regkey As RegistryKey)
        '由 Registry 讀取最近開啟的檔案資料
        MyBase.LoadRegistryInfo(regkey)

        TempFileIndex = DirectCast(regkey.GetValue(strTempFileIndex, 0), Integer)

        Dim i As Integer
        For i = 1 To TempFileIndex
            strTempFileName(i - 1) = DirectCast(regkey.GetValue(strFilePath & i.ToString, _
            Nothing), String)
        Next i

        If TempFileIndex > 0 Then
            For i = 0 To TempFileIndex - 1
                miFileTmpName(i).Text = "&" & (i + 1).ToString & " " & Path.GetFileName(strTempFileName(i))
                miFileTmpName(i).Visible = True
            Next i
            miFileTmpline.Visible = True
        End If
    End Sub
    Private Sub LoseFocusMenu(ByVal obj As Object, ByVal ea As EventArgs)
        '當滑鼠離開主功能表時
        sbpMenu.Text = "如需說明，請按一下[說明]功能表中的[說明主題]。"
    End Sub
End Class
