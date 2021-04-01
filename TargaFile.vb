'Copyright by Media Support - Didi Kunz didi@mediasupport.ch
'The same license conditions apply as CasparCG server has.
'
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

''' <summary>
''' Object to save a bitmap to a 24bit or 32bit Truevision Targa file
''' </summary>
''' <remarks></remarks>
Public Class TargaFile

   '' TGA-Header for information only (remarks in German)
   'Private Structure TgaHeader
   '   ' Größe des Identblocks in Byte, der nach dem Header (18 Byte) folgt. 
   '   ' Normalerweise 0
   '   Dim IdentSize As Byte
   '   ' Palettentyp: 0 = Keine Palette vorhanden, 1 = Palette vorhanden
   '   Dim ColorMapType As Byte
   '   ' Bildtyp: 0 = none, 1 = Indexed, 2 = RGB, 3 = Grauskale, 
   '   ' 9 = Indexed (RLE), 10 = RGB (RLE), 11 = Grauskale (RLE)
   '   Dim ImageType As Byte
   '   ' erster Eintrag in der Farbtabelle
   '   Dim ColorMapStart As Short
   '   ' Anzahl der Farben in der Farbpalette
   '   Dim ColorMapLength As Short
   '   ' Bits Per Pixel der Farbtabelle 15, 16, 24, 32
   '   Dim ColorMapBits As Byte
   '   ' X-Position des Bildes in Pixel. Normalerweise 0
   '   Dim xStart As Short
   '   ' Y-Position des Bildes in Pixel. Normalerweise 0
   '   Dim yStart As Short
   '   Dim Width As Short          ' Breite des Bildes in Pixel
   '   Dim Height As Short         ' Höhe des Bildes in Pixel
   '   Dim Bits As Byte            ' Bits Per Pixel des Bildes 8, 16, 24, 32
   '   Dim Descriptor As Byte      ' Descriptor bits des Bildes
   'End Structure

   ''' <summary>
   ''' Save the provided bitmap as a Truevision Targa file
   ''' </summary>
   ''' <param name="Filename">The filename of the Targa-file</param>
   ''' <param name="Picture">The Drawing.Bitmap object to save</param>
   Public Shared Sub SaveAsTarga(Filename As String, Picture As Bitmap)

      If Picture.PixelFormat <> Imaging.PixelFormat.Format32bppArgb Then
         Throw New Exception("Must be a 32-Bit Image")
         Exit Sub
      End If

      If Not IO.Directory.Exists(IO.Path.GetDirectoryName(Filename)) Then
         Throw New Exception("Path to save file does not exist")
         Exit Sub
      End If

      'System.Drawing.Bitmap have there pixels arranged from top to bottom, TGA's from bottom to top, so we flip the picture
      Picture.RotateFlip(RotateFlipType.RotateNoneFlipY)

      Using FS As New FileStream(Filename, FileMode.Create, FileAccess.Write)

         Dim bw As BinaryWriter = New BinaryWriter(FS)

         'Writing the Header
         Dim sh As Int16 = 0

         bw.Write(CByte(0))   'IdentSize
         bw.Write(CByte(0))   'ColorMapType
         bw.Write(CByte(2))   'ImageType

         bw.Write(sh)         'ColorMapStart
         bw.Write(sh)         'ColorMapLength
         bw.Write(CByte(0))   'ColorMapBits

         bw.Write(sh)         'xStart
         bw.Write(sh)         'yStart

         sh = CShort(Picture.Width)
         bw.Write(sh)         'Width
         sh = CShort(Picture.Height)
         bw.Write(sh)         'Height

         bw.Write(CByte(32))  'Bits Per Pixel
         bw.Write(CByte(8))   'Descriptor

         'Looking and writing the Bitmap
         Dim bmpData As BitmapData = Picture.LockBits(New Rectangle(0, 0, Picture.Width, Picture.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

         Dim ptr As IntPtr = bmpData.Scan0
         Dim bytes As Integer = bmpData.Stride * Picture.Height
         Dim rgbValues(bytes - 1) As Byte

         System.Runtime.InteropServices.Marshal.Copy(ptr, rgbValues, 0, bytes)
         bw.Write(rgbValues)

         Picture.UnlockBits(bmpData)

         'Writing the Footer
         Dim ln As Int32 = 0
         bw.Write(ln)
         bw.Write(ln)

         bw.Write("TRUEVISION-XFILE.")
         bw.Write(CByte(0))

         'Clean Up
         bw.Flush()
         FS.Flush()
         bw.Close()

      End Using

   End Sub

End Class
