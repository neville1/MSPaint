Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PaintStartPoint
    Inherits PaintWithMenuAbout

    Shared Shadows Sub Main()
        Application.Run(New PaintStartPoint)
    End Sub

    Sub New()
        ' szBmp = New Size(szDef.Width, szDef.Height)

        'bmpPic = New Bitmap(szBmp.Width, szBmp.Height)
        'Dim grfxPic As Graphics = Graphics.FromImage(bmpPic)
        'grfxPic.Clear(Color.White)
        'grfxPic.Dispose()

        '�إ߰_�l�ϧ�
    End Sub
    Protected Overrides Sub OnLoad(ByVal ea As EventArgs)
        MyBase.OnLoad(ea)

        pnWorkarea.Size = szBmp
        bmpPic = New Bitmap(szBmp.Width, szBmp.Height)
        Dim grfxPic As Graphics = Graphics.FromImage(bmpPic)
        grfxPic.Clear(Color.White)
        grfxPic.Dispose()
    End Sub
End Class
