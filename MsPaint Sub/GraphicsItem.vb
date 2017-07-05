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
    '���B�Ψ��x�s�U��ø�ϰʧ@,�P��ʧ@�һݪ����,�L�����@�欰.
    '�Ψ��x�s��ʧ@�һݸ�ƪ���k,�ٻݭn���s��z
    'DataType �N�q
    ' P = Path, D = Dot �t�XPoint(), I= Image �t�XPoint()
    ' F = Fill �t�X���| ,R = Round �t�X���|, B = Bitmap �t�X�y��

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
