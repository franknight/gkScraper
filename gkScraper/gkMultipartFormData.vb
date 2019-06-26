Imports System.Net

Public Class gkMultipartFormData

    Private pboundary As String
    Private boundaryBytes() As Byte
    Private endBoundaryBytes() As Byte

    Private Data As Dictionary(Of String, Object)

    Public Sub New()

        Data = New Dictionary(Of String, Object)

        pboundary = "----------------------------" & DateTime.Now.Ticks.ToString("x")

        boundaryBytes = System.Text.Encoding.ASCII.GetBytes(vbCrLf & "--" & pboundary & vbCrLf)

        endBoundaryBytes = System.Text.Encoding.ASCII.GetBytes(vbCrLf & "--" & Boundary & "--")

    End Sub


    Public Sub Add(ByVal name As String, ByVal value As Object)
        Data.Add(name, value)

    End Sub

    Public ReadOnly Property Boundary() As String
        Get
            Return Me.pboundary
        End Get
    End Property

    Friend Function WriteTo(ByVal requestStream As IO.Stream) As Integer

        Dim totlen As Integer = 0

        For Each pair As KeyValuePair(Of String, Object) In Me.Data

            requestStream.Write(boundaryBytes, 0, boundaryBytes.Length)
            totlen += boundaryBytes.Length

            If TypeOf pair.Value Is gkFormFile Then
                Dim file As gkFormFile = pair.Value
                Dim header As String = "Content-Disposition: form-data; name=""" & pair.Key & """; filename=""" & file.Name & """" & vbCrLf & "Content-Type: " & file.ContentType & vbCrLf & vbCrLf
                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(header)
                requestStream.Write(bytes, 0, bytes.Length)
                totlen += bytes.Length

                Dim buffer(32768) As Byte
                Dim bytesRead As Integer

                If (Not file.Stm Is Nothing) Then
                    'upload from given stream
                    Do
                        bytesRead = file.Stm.Read(buffer, 0, buffer.Length)
                        requestStream.Write(buffer, 0, bytesRead)
                        totlen += bytesRead
                    Loop While bytesRead > 0
                    'file.Stm.Close()

                ElseIf (Not file.url = "") Then
                    'Get a stream from URL
                    Dim req As HttpWebRequest = HttpWebRequest.Create(file.url)
                    Dim response As HttpWebResponse = req.GetResponse()
                    Dim stream As IO.Stream = response.GetResponseStream()
                    Do
                        bytesRead = stream.Read(buffer, 0, buffer.Length)
                        requestStream.Write(buffer, 0, bytesRead)
                        totlen += bytesRead
                    Loop While bytesRead > 0
                    stream.Close()
                Else
                    ' upload from local file
                    Dim fileStream As IO.FileStream = IO.File.OpenRead(file.FilePath)
                    Do
                        bytesRead = fileStream.Read(buffer, 0, buffer.Length)
                        requestStream.Write(buffer, 0, bytesRead)
                        totlen += bytesRead
                    Loop While bytesRead > 0

                    fileStream.Close()

                End If

            Else
                Dim data As String = "Content-Disposition: form-data; name=""" & pair.Key & """" & vbCrLf & vbCrLf & pair.Value
                Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(data)
                requestStream.Write(bytes, 0, bytes.Length)
                totlen += bytes.Length
            End If

        Next

        requestStream.Write(endBoundaryBytes, 0, endBoundaryBytes.Length)
        totlen += endBoundaryBytes.Length

        Return totlen

    End Function

End Class


Public Class gkFormFile
    Public Name As String
    Public ContentType As String
    Public FilePath As String
    Public Stm As IO.Stream
    Public url As String
End Class

