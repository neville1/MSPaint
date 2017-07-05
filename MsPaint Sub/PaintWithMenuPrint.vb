Imports System
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms
Public Class PaintWithMenuPrint
    '本類別處理功能表中,實作關於列印的行為.
    Inherits PaintWithMenuFile

    Private prndoc As New PrintDocument
    Private setdlg As New PageSetupDialog
    Private prndlg As New PrintDialog
    Private predlg As New PrintPreviewDialog
    Sub New()
        '列印相關物件建構     
        prndoc.DocumentName = "小畫家重製版文件"
        AddHandler prndoc.PrintPage, AddressOf OnPrintPage
        setdlg.Document = prndoc
        predlg.Document = prndoc
        prndlg.Document = prndoc

        prndlg.AllowPrintToFile = False
    End Sub
    '這裏的列印行為,明顯與小畫家不同.一方面這是因為小畫家定義的列印行為
    '以像素為主,依序填入各頁.對於這種方式的懷疑.
    '另一方面則是因為自己把它定位為一個學習工具,不想在列印的方面多加表述
    '這可以說是在 .Net中,最單純的列印方式了.
    Protected Overrides Sub MenuFileReviewPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '預覽列印
        predlg.ShowDialog()
    End Sub
    Protected Overrides Sub MenuFileSetPrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '設定列印格式
        setdlg.ShowDialog()
    End Sub
    Protected Overrides Sub MenuFilePrintOnClick(ByVal obj As Object, ByVal ea As EventArgs)
        '列印
        If prndlg.ShowDialog() = DialogResult.OK Then
            prndoc.Print()
        End If
    End Sub
    Private Sub OnPrintPage(ByVal obj As Object, ByVal ppea As PrintPageEventArgs)
        Dim bmpBack As Bitmap = WorkareaChangeBmp()
        Dim RectF = New RectangleF(ppea.MarginBounds.X, ppea.MarginBounds.Y, _
                               ppea.MarginBounds.Width, ppea.MarginBounds.Height)

        Dim szf As New SizeF(bmpBack.Width / bmpBack.HorizontalResolution, _
                             bmpBack.Height / bmpBack.VerticalResolution)
        Dim fScale As Single = Math.Min(RectF.Width / szf.Width, RectF.Height / szf.Height)
        szf.Width *= fScale
        szf.Height *= fScale

        ppea.Graphics.DrawImage(bmpBack, RectF.X + (RectF.Width - szf.Width) / 2, _
                                RectF.Y + (RectF.Height - szf.Height) / 2, szf.Width, szf.Height)
        '列印行為,圖形的長寬比為主,在不改長寬比的原則下,放大到最大的可能列印範圍,並置中輸出
    End Sub
End Class
