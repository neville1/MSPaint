Imports Microsoft.Win32
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintWithRegistry
    '�����O�w�q,���󥻵{���򥻪��A,�s�� Registry �欰.
    Inherits PaintMainArea

    Private rectNormal As Rectangle
    Private strProgNameReg As String = "MsPaintSub"                 '�{���W�� by Registry
    Private strRegKey As String = "Software\Programming\VBdotNet\"  '�s�� Registry ���| 

    Const strWinState As String = "WindowState"   ' Form ���A(�̤j��,�̤p��)��
    Const strLocationX As String = "LocationX"
    Const strLocationY As String = "LocationY"    '�_�l�y��
    Const strWidth As String = "Width"
    Const strHeight As String = "Height"          '�_�l Form �j�p
    Const strShowToolbar As String = "ShowToolbar"                '�u��C����t�m  
    Const strShowColorbar As String = "ShowColorbar"              '����Ϥ���t�m
    Const strShowStatusbar As String = "ShowStatusbar"            '���A�C����t�m
    Const strShowTextFrom As String = "ShowTextFrom"              '��r�u��C����t�m. 
    Const strDefaultWidth As String = "DefaultImageSizeWidth"
    Const strDefaultHeight As String = "DefaultImageSizeHeight"   '�w�]�u�@�ϧΤj�p

    Sub New()
        rectNormal = DesktopBounds
    End Sub
    Protected Overrides Sub OnMove(ByVal ea As EventArgs)
        '�� Form ���ʮ�
        MyBase.OnMove(ea)

        If WindowState = FormWindowState.Normal Then
            rectNormal = DesktopBounds
        End If
    End Sub
    Protected Overrides Sub OnResize(ByVal ea As EventArgs)
        '�� Form �j�p�ܧ��
        MyBase.OnResize(ea)

        If WindowState = FormWindowState.Normal Then
            rectNormal = DesktopBounds
        End If
    End Sub
    Protected Overrides Sub OnLoad(ByVal ea As EventArgs)
        '�b�{���Ұʮ�,���J Windows Registy �w�q�_�l��T
        MyBase.OnLoad(ea)

        strRegKey = strRegKey & strProgNameReg
        Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey(strRegKey)
        If Not regkey Is Nothing Then
            LoadRegistryInfo(regkey)
            regkey.Close()
        End If
    End Sub
    Protected Overrides Sub OnClosed(ByVal ea As EventArgs)
        '�b�{���Y�N�����e,�g�J Windows Registry �x�s�̫᪬�A
        MyBase.OnClosed(ea)

        Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey(strRegKey, True)
        If regkey Is Nothing Then
            regkey = Registry.CurrentUser.CreateSubKey(strRegKey)
        End If
        SaveRegistryInfo(regkey)
        regkey.Close()
    End Sub
    Protected Overridable Sub SaveRegistryInfo(ByVal regkey As RegistryKey)
        '�g�J Windows Registry ���
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
        'Ū�� Windows Registry ���
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
