
Imports System.Drawing

Public Class CImageTool

    Public Sub SaveImage(ByVal bm As Bitmap, ByVal strFullFileName As String, ByVal eImageFormat As Imaging.ImageFormat)
        bm.Save(strFullFileName, eImageFormat)
    End Sub

    Public Function ResizeBitmap(ByVal bm As Bitmap, ByVal nNewWidth As Integer, ByVal nNewHeight As Integer, ByVal nRotateMode As RotateFlipType) As Bitmap 'following code resizes picture to fit   
        If nNewWidth < 1 Or nNewHeight < 1 Or bm Is Nothing Then Return bm
        bm.RotateFlip(nRotateMode)
        Dim width As Integer = nNewWidth 'image width.   
        Dim height As Integer = nNewHeight 'image height   
        Dim thumb As New Bitmap(width, height)
        Dim g As Graphics = Graphics.FromImage(thumb)

        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)
        g.Dispose()
        bm.Dispose()   'image path. 
        'thumb.Save("C:\" & dir & "\" & fileName, System.Drawing.Imaging.ImageFormat.Jpeg) 'can use any image format   
        'thumb.Dispose()
        Return thumb
    End Function

    Public Function CropImage(ByVal imgOrg As Bitmap, ByVal XStart As Integer, ByVal YStart As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Bitmap
        Dim thumb As New Bitmap(nWidth, nHeight)
        Dim g As Graphics = Graphics.FromImage(thumb)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(imgOrg, New Rectangle(0, 0, nWidth, nHeight), New Rectangle(XStart, YStart, nWidth, nHeight), GraphicsUnit.Pixel)
        g.Dispose()
        imgOrg.Dispose()
        Return thumb
    End Function

    Public Sub RotateImage(ByVal bm As Bitmap, ByVal eRotateFlipType As RotateFlipType)
        bm.RotateFlip(eRotateFlipType)
    End Sub

End Class
