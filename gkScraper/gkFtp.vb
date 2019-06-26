Imports System.Net
Imports System.IO
Imports System.Text


Public Class gkFtp

    Dim host As String
    Dim user As String
    Dim pass As String
    Dim ftpRequest As FtpWebRequest
    Dim ftpResponse As FtpWebResponse
    Dim ftpStream As Stream
    Dim bufferSize As Integer = 2048

    ' Construct Object */
    Public Sub New(ByVal hostIP As String, ByVal userName As String, ByVal password As String)
        host = hostIP
        user = userName
        pass = password
    End Sub

    ' Download File */
    Public Sub download(ByVal remoteFile As String, ByVal localFile As String)

        Try

            ' Create an FTP Request */
            Dim ftpRequest As FtpWebRequest = FtpWebRequest.Create(host + "/" + remoteFile)
            ' Log in to the FTP Server with the User Name and Password Provided */
            ftpRequest.Credentials = New NetworkCredential(user, pass)
            ' When in doubt, use these options */
            ftpRequest.UseBinary = True
            ftpRequest.UsePassive = True
            ftpRequest.KeepAlive = True
            ' Specify the Type of FTP Request */
            ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile
            ' Establish Return Communication with the FTP Server */
            Dim ftpResponse As FtpWebResponse = ftpRequest.GetResponse()
            ' Get the FTP Server's Response Stream */
            ftpStream = ftpResponse.GetResponseStream()
            ' Open a File Stream to Write the Downloaded File */
            Dim localFileStream As FileStream = New FileStream(localFile, FileMode.Create)
            ' Buffer for the Downloaded Data */
            Dim byteBuffer(bufferSize) As Byte
            Dim bytesRead As Integer = ftpStream.Read(byteBuffer, 0, bufferSize)
            ' Download the File by Writing the Buffered Data Until the Transfer is Complete */
            Try

                While (bytesRead > 0)
                    localFileStream.Write(byteBuffer, 0, bytesRead)
                    bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize)
                End While

            Catch ex As Exception
                Console.WriteLine(ex.ToString())
            End Try

            ' Resource Cleanup */
            localFileStream.Close()
            ftpStream.Close()
            ftpResponse.Close()

        Catch ex As Exception
            Console.WriteLine(ex.ToString())

        End Try

        Return

    End Sub

    ' Upload File */
    Public Sub upload(ByVal localFile As String, ByVal remoteFile As String)

        Try


            ' Create an FTP Request */
            Dim ftpRequest As FtpWebRequest = FtpWebRequest.Create("ftp://" & host & remoteFile)
            ' Log in to the FTP Server with the User Name and Password Provided */
            ftpRequest.Credentials = New NetworkCredential(user, pass)
            ' When in doubt, use these options */
            'ftpRequest.UseBinary = True
            'ftpRequest.UsePassive = True
            'ftpRequest.KeepAlive = True
            ' Specify the Type of FTP Request */
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile

            'Copy the contents of the file to the request stream.
            Dim sourceStream As StreamReader = New StreamReader(localFile)
            Dim fileContents() As Byte = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd())
            sourceStream.Close()
            ftpRequest.ContentLength = fileContents.Length

            ' Establish Return Communication with the FTP Server */
            ftpStream = ftpRequest.GetRequestStream()

            'Dim requestStream As Stream = ftpRequest.GetRequestStream()
            ftpStream.Write(fileContents, 0, fileContents.Length)
            ftpStream.Close()

            Dim response As FtpWebResponse = ftpRequest.GetResponse()
            response.Close()

            '' Open a File Stream to Read the File for Upload */
            'Dim localFileStream As FileStream = New FileStream(localFile, FileMode.Create)
            '' Buffer for the Downloaded Data */
            'Dim byteBuffer(bufferSize) As Byte
            'Dim bytesSent As Integer = localFileStream.Read(byteBuffer, 0, bufferSize)
            '' Upload the File by Sending the Buffered Data Until the Transfer is Complete */
            'Try

            '    While (bytesSent <> 0)
            '        ftpStream.Write(byteBuffer, 0, bytesSent)
            '        bytesSent = localFileStream.Read(byteBuffer, 0, bufferSize)
            '    End While

            'Catch ex As Exception
            '    Throw ex
            'End Try

            '' Resource Cleanup */
            'localFileStream.Close()
            'ftpStream.Close()

        Catch e As WebException
            Dim w As FtpWebResponse = e.Response
            Debug.WriteLine(w.StatusCode & ", " & w.StatusDescription)
            'Throw e

        End Try

        Return

    End Sub

    Public Function directoryListSimple(ByVal directory As String) As String()

        Try

            ' Create an FTP Request */
            ftpRequest = FtpWebRequest.Create("ftp://" & host & "/" & directory)
            ' Log in to the FTP Server with the User Name and Password Provided */
            ftpRequest.Credentials = New NetworkCredential(user, pass)
            ' When in doubt, use these options */
            ftpRequest.UseBinary = True
            ftpRequest.UsePassive = True
            ftpRequest.KeepAlive = True
            ' Specify the Type of FTP Request */
            ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory
            ' Establish Return Communication with the FTP Server */
            ftpResponse = ftpRequest.GetResponse
            ' Establish Return Communication with the FTP Server */
            ftpStream = ftpResponse.GetResponseStream()
            ' Get the FTP Server's Response Stream */
            Dim ftpReader As StreamReader = New StreamReader(ftpStream)
            ' Store the Raw Response */
            Dim directoryRaw As String = String.Empty
            ' Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
            Try
                While (ftpReader.Peek() <> -1)
                    directoryRaw &= ftpReader.ReadLine() & "|"
                End While
            Catch ex As Exception
                Console.WriteLine(ex.ToString())
            End Try

            ' Resource Cleanup */
            ftpReader.Close()
            ftpStream.Close()
            ftpResponse.Close()

            ' Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
            Try
                Dim directoryList() As String = directoryRaw.Split("|".ToCharArray())
                Return directoryList
            Catch ex As Exception
                Console.WriteLine(ex.ToString())
            End Try

        Catch wex As WebException
            Console.WriteLine(wex.ToString())
        End Try

            ' Return an Empty string Array if an Exception Occurs */
            Return New String() {""}
    End Function


End Class
