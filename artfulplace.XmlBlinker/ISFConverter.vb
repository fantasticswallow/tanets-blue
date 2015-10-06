Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Ink
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

''' <summary>
''' Converter of ISF(Ink Serialized Format) to web browser available
''' </summary>
Public Class ISFConverter
    <STAThread>
    Public Shared Function IsfConvert(bs As String) As String
        bs = bs.Replace("&#xD;&#xA;", "")
        Dim data = Convert.FromBase64String(bs)
        Dim res = ""
        Using mstream As New MemoryStream(data)
            Dim stcol = New StrokeCollection(mstream)

            Dim drawingVisual As New DrawingVisual()
            Dim dc As DrawingContext = drawingVisual.RenderOpen()
            dc.DrawRectangle(Brushes.Transparent, Nothing, stcol.GetBounds)
            stcol.Draw(dc)
            dc.Close()

            Dim dimg = New DrawingImage(drawingVisual.Drawing)

            Dim bnds = stcol.GetBounds()

            Dim dbmp = New Image With {.Source = dimg}
            dbmp.Measure(New Size(bnds.Width, bnds.Height))
            dbmp.Arrange(New Rect(0, 0, bnds.Width, bnds.Height))
            dbmp.UpdateLayout()

            Dim bmp As New RenderTargetBitmap(bnds.Width, bnds.Height, 96, 96, PixelFormats.Default)
            'bmp.Render(drawingVisual)
            bmp.Render(dbmp)

            Dim enc As BitmapEncoder = New PngBitmapEncoder()


            Try

                Using ds As MemoryStream = New MemoryStream()
                    enc.Frames.Add(BitmapFrame.Create(bmp))
                    enc.Save(ds)
                    res = Convert.ToBase64String(ds.ToArray())

                End Using
            Catch ex As Exception
            End Try

            enc = Nothing
            bmp = Nothing
            dbmp = Nothing
            bnds = Nothing
            drawingVisual = Nothing
            stcol = Nothing

        End Using

        Return res
    End Function


End Class
