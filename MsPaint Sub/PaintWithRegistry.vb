Imports Microsoft.Win32
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithRegistry
    '本類別定義,關於本程式基本狀態,存取 Registry 行為.
    Inherits PaintMainArea

    Private rectNormal As Rectangle
    Private strProgNameReg As String = "MsPaintSub"                 '程式名稱 by Registry
    Private strRegKey As String = "Software\Programming\VBdotNet\"  '存取 Registry 路徑 

    Const strWinState As String = "WindowState"   ' Form 狀態(最大化,最小化)等
    Const strLocationX As String = "LocationX"
    Const strLocationY As String = "LocationY"    '起始座標
    Const strWidth As String = "Width"
    Const strHeight As String = "Height"          '起始 Form 大小
    Const strShowToolbar As String = "ShowToolbar"                '工具列元件配置  
    Const strShowColorbar As String = "ShowColorbar"              '色塊區元件配置
    Const strShowStatusbar As String = "ShowStatusbar"            '狀態列元件配置
    Const strShowTextFrom As String = "ShowTextFrom"              '文字工具列元件配置. 
    Const strDefaultWidth As String = "DefaultImageSizeWidth"
    Const strDefaultHeight As String = "DefaultImageSizeHeight"   '預設工作圖形大小

    Sub New()
        rectNormal = DesktopBounds
    End Sub
    Protected Overrides Sub OnMove(ByVal ea As EventArgs)
        '當 Form 移動時
        MyBase.OnMove(ea)

        If WindowState = FormWindowState.Normal Then
            rectNormal = DesktopBounds
        End If
    End Sub
    Protected Overrides Sub OnResize(ByVal ea As EventArgs)
        '當 Form 大小變更時
        MyBase.OnResize(ea)

        If WindowState = FormWindowState.Normal Then
            rectNormal = DesktopBounds
        End If
    End Sub
    Protected Overrides Sub OnLoad(ByVal ea As EventArgs)
        '在程式啟動時,載入 Windows Registy 定義起始資訊
        MyBase.OnLoad(ea)

        strRegKey = strRegKey & strProgNameReg
        Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey(strRegKey)
        If Not regkey Is Nothing Then
            LoadRegistryInfo(regkey)
            regkey.Close()
        End If
    End Sub
    Protected Overrides Sub OnClosed(ByVal ea As EventArgs)
        '在程式即將關閉前,寫入 Windows Registry 儲存最後狀態
        MyBase.OnClosed(ea)

        Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey(strRegKey, True)
        If regkey Is Nothing Then
            regkey = Registry.CurrentUser.CreateSubKey(strRegKey)
        End If
        SaveRegistryInfo(regkey)
        regkey.Close()
    End Sub
    Protected Overridable Sub SaveRegistryInfo(ByVal regkey As RegistryKey)
        '寫入 Windows Registry 資料
        regkey.SetValue(strWinState, CInt(WindowState))
        regkey.SetValue(strLocationX, rectNormal.X)
        regkey.SetValue(strLocationY, rectNormal.Y)
        regkey.SetValue(strWidth, rectNormal.Width)
        regkey.SetValue(strHeight, rectNormal.Height)
        regkey.SetValue(strShowToolbar, bToolbar)
        regkey.SetValue(strShowColorbar, bColorbar)
        regkey.SetValue(strShowStatusbar, bStatusbar)
        regkey.SetValue(strShowTextFrom, bTextFrom)
        regkey.SetValue(strDefaultWidth, szBmp.Width)
        regkey.SetValue(strDefaultHeight, szBmp.Height)
    End Sub
    Protected Overridable Sub LoadRegistryInfo(ByVal regkey As RegistryKey)
        '讀取 Windows Registry 資料
        Dim x As Integer = DirectCast(regkey.GetValue(strLocationX, 100), Integer)
        Dim y As Integer = DirectCast(regkey.GetValue(strLocationY, 100), Integer)
        Dim cx As Integer = DirectCast(regkey.GetValue(strWidth, 275), Integer)
        Dim cy As Integer = DirectCast(regkey.GetValue(strHeight, 400), Integer)

        Dim strbT As String = DirectCast(regkey.GetValue(strShowToolbar, "True"), String)
        If strbT = "True" Then bToolbar = True Else bToolbar = False
        Dim strbC As String = DirectCast(regkey.GetValue(strShowColorbar, "True"), String)
        If strbC = "True" Then bColorbar = True Else bColorbar = False
        Dim strbS As String = DirectCast(regkey.GetValue(strShowStatusbar, "True"), String)
        If strbS = "True" Then bStatusbar = True Else bStatusbar = False
        Dim strbTF As String = DirectCast(regkey.GetValue(strShowTextFrom, "True"), String)
        If strbTF = "True" Then bTextFrom = True Else bTextFrom = False

        Dim width As Integer = DirectCast(regkey.GetValue(strDefaultWidth, 512), Integer)
        Dim height As Integer = DirectCast(regkey.GetValue(strDefaultHeight, 384), Integer)

        rectNormal = New Rectangle(x, y, cx, cy)
        Dim rectDesk As Rectangle = SystemInformation.WorkingArea
        rectNormal.X -= Math.Min(rectNormal.Right - rectDesk.Right, 0)
        rectNormal.Y -= Math.Min(rectNormal.Bottom - rectDesk.Bottom, 0)

        DesktopBounds = rectNormal
        WindowState = CType(DirectCast(regkey.GetValue(strWinState, 0), Integer), _
        FormWindowState)

        szBmp = New Size(width, height)
    End Sub
End Class
