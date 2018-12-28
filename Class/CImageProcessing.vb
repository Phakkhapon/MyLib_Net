Imports System.Drawing

Public Class CImageProcessing
    Private m_bmOriginal As Bitmap

    Public Sub New(ByVal strPicPath As String)
        Dim bm As New Bitmap(strPicPath)
        m_bmOriginal = bm
    End Sub

    Public Sub New(ByVal bmPic As Bitmap)
        m_bmOriginal = bmPic
    End Sub

    Public Sub SaveImage(ByVal bm As Bitmap, ByVal strFullFileName As String, ByVal eImageFormat As Imaging.ImageFormat)
        bm.Save(strFullFileName, eImageFormat)
    End Sub

    Public Function ResizeBitmap(ByVal nNewWidth As Integer, ByVal nNewHeight As Integer, ByVal nRotateMode As RotateFlipType) As Bitmap 'following code resizes picture to fit   
        If nNewWidth < 1 Or nNewHeight < 1 Or m_bmOriginal Is Nothing Then Return Nothing
        Dim width As Integer = nNewWidth 'image width.   
        Dim height As Integer = nNewHeight 'image height   
        Dim bmResize As New Bitmap(width, height)
        Dim g As Graphics = Graphics.FromImage(bmResize)

        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(bmResize, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bmResize.Width, bmResize.Height), GraphicsUnit.Pixel)
        bmResize.RotateFlip(nRotateMode)
        g.Dispose()
        Return bmResize
    End Function

    Public Function CropImage(ByVal XStart As Integer, ByVal YStart As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Bitmap
        Dim bmCrop As New Bitmap(nWidth, nHeight)
        Dim g As Graphics = Graphics.FromImage(bmCrop)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(m_bmOriginal, New Rectangle(0, 0, nWidth, nHeight), New Rectangle(XStart, YStart, nWidth, nHeight), GraphicsUnit.Pixel)
        g.Dispose()
        Return bmCrop
    End Function

    Public Function RotateImage(ByVal eRotateFlipType As RotateFlipType) As Bitmap
        Dim bmRotate As Bitmap = m_bmOriginal.Clone
        bmRotate.RotateFlip(eRotateFlipType)
        RotateImage = bmRotate
    End Function

End Class
