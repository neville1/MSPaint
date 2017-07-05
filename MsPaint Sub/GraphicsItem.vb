Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class GraphicsItem

    Private Type As Char

    Private GrapPath As GraphicsPath
    Private clr As Color
    Private brClr As Color
    Private penWidth As Integer = 0

    Private EditImage As Image
    Private Editlocation As Point
    Private EditSize As Size

    Private bSizeChange As Boolean = False
    Private bFont As Boolean = False
    Private SourSize As Size
    '此處用來儲存各種繪圖動作,與其動作所需的資料,無任何實作行為.
    '用來儲存其動作所需資料的方法,還需要重新整理
    'DataType 意義
    ' P = Path, D = Dot 配合Point(), I= Image 配合Point()
    ' F = Fill 配合路徑 ,R = Round 配合路徑, B = Bitmap 配合座標

    Property DataType() As Char
        Get
            Return Type
        End Get
        Set(ByVal Value As Char)
            Type = Value
        End Set
    End Property
    Property Data() As GraphicsPath
        Get
            Return GrapPath
        End Get
        Set(ByVal Value As GraphicsPath)
            GrapPath = Value
        End Set
    End Property
    Property Color() As Color
        Get
            Return clr
        End Get
        Set(ByVal Value As Color)
            clr = Value
        End Set
    End Property
    Property pWidth() As Integer
        Get
            Return penWidth
        End Get
        Set(ByVal Value As Integer)
            penWidth = Value
        End Set
    End Property
    Property FillColor() As Color
        Get
            Return brClr
        End Get
        Set(ByVal Value As Color)
            brClr = Value
        End Set
    End Property

    Property Image() As Image
        Get
            Return EditImage
        End Get
        Set(ByVal Value As Image)
            EditImage = Value
        End Set
    End Property
    Property ImageLocation() As Point
        Get
            Return Editlocation
        End Get
        Set(ByVal Value As Point)
            Editlocation = Value
        End Set
    End Property
    Property ImageSize() As Size
        Get
            Return EditSize
        End Get
        Set(ByVal Value As Size)
            EditSize = Value
        End Set
    End Property

    Property IsSizeChange() As Boolean
        Get
            Return bSizeChange
        End Get
        Set(ByVal Value As Boolean)
            bSizeChange = Value
        End Set
    End Property
    Property IsFont() As Boolean
        Get
            Return bFont
        End Get
        Set(ByVal Value As Boolean)
            bFont = Value
        End Set
    End Property
    Property SourceSize() As Size
        Get
            Return SourSize
        End Get
        Set(ByVal Value As Size)
            SourSize = Value
        End Set
    End Property

    Sub New()
        GrapPath = New GraphicsPath
        clr = New Color
        brClr = New Color

    End Sub

End Class
