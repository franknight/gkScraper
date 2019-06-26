Imports System.Threading
Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Text
Imports HtmlParserSharp
Imports System.Runtime.Serialization.Json
Imports System.Configuration
Imports NLog
Imports NLog.Config

Public Class gkParser

    Private p_Doc As XmlDocument
    Private p_Cookies As CookieContainer
    Private p_Parser As SimpleHtmlParser
    Private p_contentType As String
    Private p_tag2ExcInstruction As String
    Private p_tag2IncInstruction As String
    Private p_tag2exclude As ArrayList
    Private p_tag2include As Hashtable
    Private p_headers As gkHeaderCollection
    Private LogMgr As NLog.Logger

    Public Sub New()
        Dim conf As Config.LoggingConfiguration = New LoggingConfiguration
        LogManager.Configuration = conf
        LogManager.Configuration.AddTarget("gkscraper_log_file", New Targets.FileTarget With {.FileName = "gkscraper.log"})
        LogManager.Configuration.AddTarget("gkscraper_console", New Targets.ConsoleTarget With {.Layout = "${longdate} ${callsite} ${level} ${message}"})
        LogManager.Configuration.LoggingRules.Add(New LoggingRule("gkscraper*", LogLevel.Debug, LogManager.Configuration.FindTargetByName("gkscraper_log_file")))
        LogManager.Configuration.LoggingRules.Add(New LoggingRule("gkscraper*", LogLevel.Debug, LogManager.Configuration.FindTargetByName("gkscraper_console")))
        LogManager.ReconfigExistingLoggers()

        Dim l As Logger = LogManager.GetLogger("gkscraper")

        internal_new(l)
    End Sub
    Friend Sub New(ByVal LogMgr As NLog.Logger)
        internal_new(LogMgr)
    End Sub

    Private Sub internal_new(ByVal LogMgr As NLog.Logger)
        Me.LogMgr = LogMgr

        p_Parser = New SimpleHtmlParser
        LogMgr.Trace("SimpleHtmlParser created.")

        p_tag2ExcInstruction = "SCRIPT|META|LINK|STYLE"
        p_tag2IncInstruction = "A:href|IMG:src,alt,big|INPUT:type,value,name"
        LoadPreferences()

    End Sub
    Public Property ExcludeInstructions() As String
        Get
            Return p_tag2ExcInstruction
        End Get
        Set(ByVal value As String)
            p_tag2ExcInstruction = value.ToUpper
            LoadPreferences()
        End Set
    End Property
    Public Property IncludeInstructions() As String
        Get
            Return p_tag2IncInstruction
        End Get
        Set(ByVal value As String)
            p_tag2IncInstruction = value '.ToUpper
            LoadPreferences()
        End Set
    End Property

    Public Property CustomHeaders() As gkHeaderCollection
        Get
            Return p_headers
        End Get
        Set(ByVal value As gkHeaderCollection)
            p_headers = value
        End Set
    End Property

    Private Sub LoadPreferences()
        Dim tmp As String
        Dim tmpa As String()

        Try

            p_tag2exclude = New ArrayList

            tmp = p_tag2ExcInstruction
            tmpa = tmp.Split("|")
            p_tag2exclude.AddRange(tmpa)

            p_tag2include = New Hashtable
            tmp = p_tag2IncInstruction
            tmpa = tmp.Split("|")
            For Each tag_string As String In tmpa
                Dim atts As ArrayList
                Dim tmparr As String()
                Dim tag As String
                atts = New ArrayList
                tmparr = tag_string.Split(":")
                tag = tmparr(0).ToUpper
                atts.AddRange(tmparr(1).Split(","))
                p_tag2include.Add(tag, atts)
            Next

        Catch ex As Exception
            LogMgr.ErrorException("Error loading parser preferences.", ex)
        End Try

    End Sub

    Public ReadOnly Property Document() As XmlDocument
        Get
            Return p_Doc
        End Get
    End Property
    Public ReadOnly Property ContentType() As String
        Get
            Return p_contentType
        End Get
    End Property

    '
    'Navigate corrisponde al GET
    '
    Public Function Navigate(ByVal url As String) As String
        Return Navigate(url, Nothing, False)
    End Function
    Public Function Navigate(ByVal url As String, ByVal dontParse As Boolean) As String
        Return Navigate(url, Nothing, dontParse)
    End Function
    Public Function Navigate(ByVal url As String, ByVal Parameters As String) As String
        Return Navigate(url, Parameters, False)
    End Function
    Public Function Navigate(ByVal url As String, ByVal Parameters As String, ByVal dontParse As Boolean) As String

        Dim text As String = String.Empty

        If (p_Cookies Is Nothing) Then
            p_Cookies = New CookieContainer()
        End If

        'text = GetData(type, url, Parameters, p_Cookies)
        Dim dataStream As Stream
        Dim response_Cookies As New CookieCollection

        ' Create a request using a URL that can receive a post. 
        ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault

        Dim request As HttpWebRequest = WebRequest.Create(url)
        request.CookieContainer = p_Cookies
        request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; rv:67.0) Gecko/20100101"
        request.Headers.Add("Accept-Language", "it-IT,it;q=0.8,en-US;q=0.5,en;q=0.3")
        'request.Headers.Add("Accept-Encoding", "gzip, deflate, br")
        'request.Connection = "keep-alive"
        'request.Headers.Add("Upgrade-Insecure-Requests", "1")
        'request.Headers.Add("Pragma", "no-cache")
        'request.Headers.Add("Cache-Control", "no-cache")


        request.AllowAutoRedirect = False
        request.Method = "GET"

        LogMgr.Info("-----------")
        LogMgr.Info("GET:{0} ? {1}", url, Parameters)
        LogMgr.Info("-----------")

        Dim response As HttpWebResponse
        Try
            response = request.GetResponse()
        Catch wex As Net.WebException
            LogMgr.Error("Errore navigando su " & url, wex)
            'If wex.Status = WebExceptionStatus.ProtocolError Then
            'End If
            Throw wex
        End Try

        'Get set-cookie header 
        For i As Integer = 0 To response.Headers.Count - 1
            Dim name As String = response.Headers.GetKey(i)
            If (name.ToLower <> "set-cookie") Then
                Continue For
            End If
            Dim value As String = response.Headers.Get(i)
            Dim var As CookieCollection = ParseCookieString(value, request.RequestUri.Host.Split(":")(0))
            response_Cookies.Add(var)
        Next
        p_Cookies.Add(response_Cookies)

        Dim status As Integer = response.StatusCode
        Dim new_location As String
        new_location = String.Empty
        If status > 300 And status < 400 Then
            new_location = response.Headers("location")
        End If

        ' Display the status.
        LogMgr.Info("Response status: {0}", response.StatusDescription)
        LogMgr.Trace("Headers:")
        For i As Integer = 0 To response.Headers.Count - 1
            If response.Headers.AllKeys(i) = "Content-Type" Then
                Me.p_contentType = response.Headers(i)
            End If
            LogMgr.Trace(response.Headers.AllKeys(i) & "=" & response.Headers(i))
        Next

        LogMgr.Trace("Cookies received: {0}", response_Cookies.Count)
        For Each cook As Cookie In response_Cookies
            LogMgr.Trace("Cookie:")
            LogMgr.Trace(" {0} = {1}", cook.Name, cook.Value)
            LogMgr.Trace(" Domain: {0}", cook.Domain)
            LogMgr.Trace(" Path: {0}", cook.Path)
            LogMgr.Trace(" Port: {0}", cook.Port)
            LogMgr.Trace(" Secure: {0}", cook.Secure)
            LogMgr.Trace(" When issued: {0}", cook.TimeStamp)
            LogMgr.Trace(" Expires: {0} (expired? {1})", cook.Expires, cook.Expired)
            LogMgr.Trace(" Don't save: {0}", cook.Discard)
            LogMgr.Trace(" Comment: {0}", cook.Comment)
            LogMgr.Trace(" Uri for comments: {0}", cook.CommentUri)
            LogMgr.Trace(" Version: RFC {0}", IIf(cook.Version = 1, "2109", "2965"))
            LogMgr.Trace(" String: {0}", cook.ToString())
        Next

        dataStream = response.GetResponseStream()
        Dim reader As StreamReader = New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()

        reader.Close()
        dataStream.Close()
        response.Close()

        'Manual redirection management 
        If status > 300 And status < 400 Then
            responseFromServer = Navigate(new_location) 'GetData("GET", new_location, "", p_Cookies)
        End If

        text = responseFromServer


        If Me.p_contentType = "application/json" Then
            StartParsingJson(text)
        Else

            Dim f As FileStream = New FileStream("debug.html", FileMode.Create)
            Dim sw As StreamWriter = New StreamWriter(f)
            sw.Write(text)
            sw.Close()

            StartParsingHtml(text)

        End If

        Return text

    End Function

    '
    'POST
    '
    Public Function Post(ByVal url As String, ByVal Data As String) As String
        Return Post(url, Data, False)
    End Function
    Public Function Post(ByVal url As String, ByVal Data As String, ByVal dontParse As Boolean) As String
        Return internal_post(url, Data, dontParse)
    End Function
    Public Function Post(ByVal url As String, ByVal Data As gkMultipartFormData) As String
        Return Post(url, Data, False)
    End Function
    Public Function Post(ByVal url As String, ByVal Data As gkMultipartFormData, ByVal dontParse As Boolean) As String
        Return internal_post(url, Data, dontParse)
    End Function

    Private Function internal_post(ByVal url As String, ByVal data As Object, ByVal dontparse As Boolean) As String

        Dim text As String = String.Empty

        If (p_Cookies Is Nothing) Then
            p_Cookies = New CookieContainer()
        End If

        Dim dataStream As Stream
        Dim response_Cookies As New CookieCollection
        Dim byteArray As Byte() = {}

        ' Create a request using a URL that can receive a post. 
        Dim request As HttpWebRequest = WebRequest.Create(url)
        request.CookieContainer = p_Cookies
        request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:67.0) Gecko/20100101 Firefox/67.0"
        request.AllowAutoRedirect = False
        request.Method = "POST"

        If Not p_headers Is Nothing Then
            For Each key As String In p_headers.AllKeys
                If key.ToLower = "accept" Then
                    request.Accept = p_headers(key)
                Else
                    request.Headers.Add(key, p_headers(key))
                End If
            Next
        End If
        'request.Headers.Add("X-Requested-With", "XMLHttpRequest")


        'Controllo che tipo di dati di post sono stati passati
        If TypeOf data Is String Then
            Dim postData As String = data
            byteArray = Encoding.UTF8.GetBytes(postData)
            request.ContentType = "application/x-www-form-urlencoded; charset=UTF-8" '"application/x-www-form-urlencoded" '
            request.ContentLength = byteArray.Length

            dataStream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

        ElseIf TypeOf data Is gkMultipartFormData Then
            'TODO
            Dim postdata As gkMultipartFormData
            postdata = data

            request.ContentType = "multipart/form-data; boundary=" & postdata.Boundary

            dataStream = request.GetRequestStream()
            Dim len As Integer = postdata.WriteTo(dataStream)
            dataStream.Close()

            'request.ContentLength = len


        End If


        

        LogMgr.Info("-----------")
        LogMgr.Info("POST :{0}?{1}", url, data.ToString)
        LogMgr.Info("-----------")

        LogMgr.Trace("Request Headers")
        For Each el As Object In request.Headers
            LogMgr.Trace(" {0}: {1}", el, request.Headers(el))
        Next


        Dim response As HttpWebResponse = request.GetResponse()

        'Get set-cookie header 
        For i As Integer = 0 To response.Headers.Count - 1

            Dim name As String = response.Headers.GetKey(i)
            If (name.ToLower <> "set-cookie") Then
                Continue For
            End If
            Dim value As String = response.Headers.Get(i)
            Dim var As CookieCollection = ParseCookieString(value, request.RequestUri.Host.Split(":")(0))
            response_Cookies.Add(var)
        Next
        p_Cookies.Add(response_Cookies)


        Dim status As Integer = response.StatusCode
        Dim new_location As String
        new_location = String.Empty
        If status > 300 And status < 400 Then
            new_location = response.Headers("location")
        End If

        ' Display the status.
        LogMgr.Info("Response status: {0}", response.StatusDescription)
        LogMgr.Trace("Headers:")
        For i As Integer = 0 To response.Headers.Count - 1
            If response.Headers.AllKeys(i) = "Content-Type" Then
                Me.p_contentType = response.Headers(i)
            End If
            LogMgr.Trace(response.Headers.AllKeys(i) & "=" & response.Headers(i))
        Next

        LogMgr.Trace("Cookies received: {0}", response_Cookies.Count)
        For Each cook As Cookie In response_Cookies
            LogMgr.Trace("Cookie:")
            LogMgr.Trace(" {0} = {1}", cook.Name, cook.Value)
            LogMgr.Trace(" Domain: {0}", cook.Domain)
            LogMgr.Trace(" Path: {0}", cook.Path)
            LogMgr.Trace(" Port: {0}", cook.Port)
            LogMgr.Trace(" Secure: {0}", cook.Secure)
            LogMgr.Trace(" When issued: {0}", cook.TimeStamp)
            LogMgr.Trace(" Expires: {0} (expired? {1})", cook.Expires, cook.Expired)
            LogMgr.Trace(" Don't save: {0}", cook.Discard)
            LogMgr.Trace(" Comment: {0}", cook.Comment)
            LogMgr.Trace(" Uri for comments: {0}", cook.CommentUri)
            LogMgr.Trace(" Version: RFC {0}", IIf(cook.Version = 1, "2109", "2965"))
            LogMgr.Trace(" String: {0}", cook.ToString())
        Next

        dataStream = response.GetResponseStream()
        Dim reader As StreamReader = New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()

        reader.Close()
        dataStream.Close()
        response.Close()

        'Manage redirection
        If status > 300 And status < 400 Then
            responseFromServer = Navigate(new_location) 'GetData("GET", new_location, "", p_Cookies)
        End If

        text = responseFromServer


        If Me.p_contentType = "application/json" Then
            StartParsingJson(text)

        Else
            Dim f As FileStream = New FileStream("debug.html", FileMode.Create)
            Dim sw As StreamWriter = New StreamWriter(f)
            sw.Write(text)
            sw.Close()

            StartParsingHtml(text)

        End If

        Return text

    End Function


    Public Sub Download(ByVal url As String, ByVal fullpath As String)

        Dim request As System.Net.HttpWebRequest
        Dim a() As String
        Dim fname As String
        Dim response As HttpWebResponse = Nothing
        Dim receiveStream As Stream
        Dim br As BinaryReader
        Dim fout As FileStream
        Dim fsw As BinaryWriter

        a = Split(url, "/")
        fname = a(UBound(a))
        If fullpath <> "" Then
            fname = fullpath
            fname = Replace(fname, "/", "-")
        End If

retry:
        Try

            request = HttpWebRequest.Create(url)
            request.Timeout = 60000
            request.MaximumAutomaticRedirections = 4
            request.MaximumResponseHeadersLength = 4
            request.Credentials = CredentialCache.DefaultCredentials

            response = CType(request.GetResponse(), HttpWebResponse)

        Catch wex As WebException
            Throw wex

        Catch ex As Exception
            LogMgr.ErrorException("Error during download. (" & url & ")", ex)
        End Try

        receiveStream = response.GetResponseStream()
        br = New BinaryReader(receiveStream)

        fout = New FileStream(fullpath, FileMode.Create)
        fsw = New BinaryWriter(fout)
        Try
            fsw.Write(br.ReadBytes(response.ContentLength))
        Catch wex As WebException
            If wex.Status = WebExceptionStatus.Timeout Then
                GoTo retry
            End If
        Finally
            fout.Close()
        End Try

        fsw.Close()
        response.Close()
        br.Close()

    End Sub
    Public Function DownloadToString(ByVal url As String) As String

        Dim request As System.Net.HttpWebRequest
        Dim a() As String
        Dim fname As String
        Dim ret As String

        a = Split(url, "/")
        fname = a(UBound(a))

        request = HttpWebRequest.Create(url)
        request.Timeout = 30000
        request.MaximumAutomaticRedirections = 4
        request.MaximumResponseHeadersLength = 4
        request.Credentials = CredentialCache.DefaultCredentials

        Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
        Dim receiveStream As Stream = response.GetResponseStream()

        Dim br As New StreamReader(receiveStream)
        ret = br.ReadToEnd()

        response.Close()
        br.Close()

        Return ret

    End Function

    Private Function OBSOLETE_GetData(ByVal type As String, ByVal url As String, ByVal data As String, ByVal cookies As CookieContainer) As String

        Dim dataStream As Stream
        Dim response_Cookies As New CookieCollection


        ' Create a request using a URL that can receive a post. 
        Dim request As HttpWebRequest = WebRequest.Create(url)
        request.CookieContainer = cookies
        request.AllowAutoRedirect = False
        request.Method = type
        If (type = "POST") Then
            Dim postData As String = data
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            request.ContentType = "application/x-www-form-urlencoded" '
            request.ContentLength = byteArray.Length
            dataStream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()
        End If



        LogMgr.Info("-----------")
        LogMgr.Info("REQUEST:{0}-{1}-{2}", url, type, data)
        LogMgr.Info("-----------")
        'LogMgr.Trace("Sending {0} request.", type)

        Dim response As HttpWebResponse = request.GetResponse()


        'cookies.Add(response.Cookies)
        'Get set-cookie header 
        For i As Integer = 0 To response.Headers.Count - 1

            Dim name As String = response.Headers.GetKey(i)
            If (name.ToLower <> "set-cookie") Then
                Continue For
            End If
            Dim value As String = response.Headers.Get(i)
            Dim var As CookieCollection = ParseCookieString(value, request.RequestUri.Host.Split(":")(0))
            response_Cookies.Add(var)
        Next
        cookies.Add(response_Cookies)


        Dim status As Integer = response.StatusCode
        Dim new_location As String
        new_location = String.Empty
        If status > 300 And status < 400 Then
            new_location = response.Headers("location")
        End If

        ' Display the status.
        LogMgr.Info("Response status: {0}", response.StatusDescription)
        LogMgr.Trace("Headers:")
        For i As Integer = 0 To response.Headers.Count - 1
            If response.Headers.AllKeys(i) = "Content-Type" Then
                Me.p_contentType = response.Headers(i)
            End If
            LogMgr.Trace(response.Headers.AllKeys(i) & "=" & response.Headers(i))
        Next

        LogMgr.Trace("Cookies received: {0}", response_Cookies.Count)
        For Each cook As Cookie In response_Cookies
            LogMgr.Trace("Cookie:")
            LogMgr.Trace(" {0} = {1}", cook.Name, cook.Value)
            LogMgr.Trace(" Domain: {0}", cook.Domain)
            LogMgr.Trace(" Path: {0}", cook.Path)
            LogMgr.Trace(" Port: {0}", cook.Port)
            LogMgr.Trace(" Secure: {0}", cook.Secure)
            LogMgr.Trace(" When issued: {0}", cook.TimeStamp)
            LogMgr.Trace(" Expires: {0} (expired? {1})", cook.Expires, cook.Expired)
            LogMgr.Trace(" Don't save: {0}", cook.Discard)
            LogMgr.Trace(" Comment: {0}", cook.Comment)
            LogMgr.Trace(" Uri for comments: {0}", cook.CommentUri)
            LogMgr.Trace(" Version: RFC {0}", IIf(cook.Version = 1, "2109", "2965"))
            LogMgr.Trace(" String: {0}", cook.ToString())
        Next

        dataStream = response.GetResponseStream()
        Dim reader As StreamReader = New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        'LogMgr.Trace (responseFromServer)
        reader.Close()
        dataStream.Close()
        response.Close()

        'Manage redirection
        If status > 300 And status < 400 Then
            responseFromServer = OBSOLETE_GetData("GET", new_location, "", cookies)
        End If

        Return responseFromServer

    End Function

    Private Sub StartParsingHtml(ByVal html As String)

        Dim xmld As XmlDocument
        'Try
        xmld = p_Parser.ParseString(html)
        ''xmld.Save("gkParser_debug.xml")
        'Catch ex As Exception
        'Throw ex
        'End Try

        p_Doc = New XmlDocument()
        'Dim docel As XmlElement = p_Doc.CreateElement("DOCEL")
        'p_Doc.AppendChild(docel)
        ExploreHtml(xmld.DocumentElement, p_Doc)

    End Sub
    Private Sub ExploreHtml(ByVal docn As XmlNode, ByVal outel As XmlNode)

        Dim tag As String
        Dim outn As XmlNode

        'se è un tipo testo lo aggiungo.
        If TypeOf docn Is Xml.XmlText Then
            'Debug.Print(Asc(docn.Value))
            outn = p_Doc.CreateTextNode(docn.Value)

        Else
            Dim docel As XmlElement
            docel = docn
            tag = docel.Name.ToUpper

            ''''DEBUG
            If tag = "PRE" Then
                If docel.GetAttribute("class") = "de1" Then
                    Debug.Assert(True)
                End If
            End If
            ''''DEBUG

            'EXCLUDED TAGS
            If Me.p_tag2exclude.Contains(tag) Then
                Return
            End If

            'se non è da scartare creo un elemento
            outn = p_Doc.CreateElement(tag)


            'DEFAULT ARRIBUTES
            Dim a As XmlAttribute
            Dim outa As XmlAttribute
            a = docel.GetAttributeNode("id")
            If Not a Is Nothing Then
                outa = p_Doc.CreateAttribute("id")
                outa.Value = a.Value
                outn.Attributes.Append(outa)
            End If

            a = docel.GetAttributeNode("class")
            If Not a Is Nothing Then
                outa = p_Doc.CreateAttribute("class")
                outa.Value = a.Value
                outn.Attributes.Append(outa)
            End If

            'ATTRIBUTES
            Dim atts As ArrayList
            If Me.p_tag2include.ContainsKey(tag) Then
                atts = Me.p_tag2include(tag)
                For Each att As String In atts
                    a = docel.GetAttributeNode(att)
                    If Not a Is Nothing Then
                        outa = p_Doc.CreateAttribute(att)
                        outa.Value = a.Value
                        outn.Attributes.Append(outa)
                    End If
                Next
            End If

            'If tag = "A" Then
            '    a = docel.GetAttributeNode("href")
            '    If Not a Is Nothing Then
            '        outa = p_Doc.CreateAttribute("href")
            '        outa.Value = a.Value
            '        outn.Attributes.Append(outa)
            '    End If
            'End If
            'If tag = "IMG" Then
            '    a = docel.GetAttributeNode("src")
            '    If Not a Is Nothing Then
            '        outa = p_Doc.CreateAttribute("src")
            '        outa.Value = a.Value
            '        outn.Attributes.Append(outa)
            '    End If
            'End If

            '
            ' Continuo a navigare nei figli ricorsivamente
            '
            'Dim el As XmlElement
            If docel.ChildNodes.Count > 0 Then
                For Each n In docel.ChildNodes
                    ExploreHtml(n, outn)
                Next
            End If

        End If

        outel.AppendChild(outn)

    End Sub
    Private Sub StartParsingJson(ByVal json As String)

        Dim doc As XmlDocument = New XmlDocument()
        Dim reader As XmlDictionaryReader
        reader = JsonReaderWriterFactory.CreateJsonReader(Encoding.UTF8.GetBytes(json), XmlDictionaryReaderQuotas.Max)
        Dim xml As XElement = XElement.Load(reader)
        doc.LoadXml(xml.ToString())

        p_Doc = doc

    End Sub

    Private Function PurgeText(ByVal text As String) As String
        text = System.Text.RegularExpressions.Regex.Replace(text, "^\s+", "")
        text = System.Text.RegularExpressions.Regex.Replace(text, "\s+$", "")
        text = text.Replace(Chr(0), "")
        text = text.Replace(Chr(19), "")
        text = text.Replace(Chr(26), "")
        Return text
    End Function


    Private Function ParseCookieString(ByVal cookieString As String, ByVal host As String) As CookieCollection

        Dim cookieCollection As CookieCollection = New CookieCollection()

        'Splitto la cookiestring nei differenti cookie (cerco una virgola senza spazi per non confonderla con le virgole delle date)
        Dim a = System.Text.RegularExpressions.Regex.Matches(cookieString, "[\w/](,)[\w/]", RegularExpressions.RegexOptions.ExplicitCapture)
        Dim clist As New ArrayList
        If a.Count = 0 And cookieString.Length > 0 Then
            'Se ce n'è solo uno
            clist.Add(cookieString)
        Else
            'se ce n'è più di uno
            If a.Count = 1 Then
                'ce n'è uno prima della virgola e uno dopo la virgola
                clist.Add(cookieString.Substring(0, a(0).Index + 1))
                clist.Add(cookieString.Substring(a(0).Index + 2))
            ElseIf a.Count > 1 Then
                Dim i As Integer
                clist.Add(cookieString.Substring(0, a(0).Index + 1))
                For i = 1 To a.Count - 1
                    clist.Add(cookieString.Substring(a(i - 1).Index + 2, a(i).Index - a(i - 1).Index - 1))
                Next
                clist.Add(cookieString.Substring(a(i - 1).Index + 2))
            End If
        End If

        For Each singleCookie As String In clist

            'Dim match As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(singleCookie, "(.+?)=(.+?);")
            'If (match.Captures.Count = 0) Then
            '    Continue For
            'End If
            'cc.Add(New Cookie(match.Groups(1).ToString(), match.Groups(2).ToString(), "/", host))

            Dim secure As Boolean = False
            Dim httpOnly As Boolean = False

            Dim domainFromCookie As String = String.Empty
            Dim path As String = String.Empty
            Dim expiresString As String = String.Empty

            Dim cookiesValues As Dictionary(Of String, String) = New Dictionary(Of String, String)

            Dim cookieValuePairsStrings() As String = singleCookie.Split(New String() {"; "}, StringSplitOptions.RemoveEmptyEntries)
            For Each cookieValuePairString As String In cookieValuePairsStrings

                Dim propertyName As String
                Dim propertyValue As String
                Dim hasval As Boolean
                Dim epos As Integer = cookieValuePairString.IndexOf("=")
                If epos > 0 Then
                    hasval = True
                    propertyName = cookieValuePairString.Substring(0, epos)
                    propertyValue = cookieValuePairString.Substring(epos + 1)
                Else
                    hasval = False
                    propertyName = cookieValuePairString
                End If

                'Dim pairArr As Object = cookieValuePairString.Split("=")
                'Dim pairArrLength As Integer = pairArr.Length
                'For i As Integer = 0 To pairArrLength - 1
                'pairArr(i) = pairArr(i).Trim()
                'Next

                'Dim propertyName As String = pairArr(0)
                If Not hasval Then
                    If (propertyName.Equals("httponly", StringComparison.OrdinalIgnoreCase)) Then
                        httpOnly = True
                    ElseIf (propertyName.Equals("secure", StringComparison.OrdinalIgnoreCase)) Then
                        secure = True
                    Else
                        LogMgr.Error(String.Format("Unknown cookie property ""{0}"". All cookie is ""{1}""", propertyName, cookieString))
                    End If
                    Continue For
                End If

                'Dim propertyValue As String = pairArr(1)
                If (propertyName.Equals("expires", StringComparison.OrdinalIgnoreCase)) Then
                    expiresString = propertyValue
                ElseIf (propertyName.Equals("domain", StringComparison.OrdinalIgnoreCase)) Then
                    domainFromCookie = propertyValue
                    If domainFromCookie.StartsWith(".") Then
                        domainFromCookie = domainFromCookie.Substring(1)
                    End If
                ElseIf (propertyName.Equals("path", StringComparison.OrdinalIgnoreCase)) Then
                    path = propertyValue
                Else
                    cookiesValues.Add(propertyName, propertyValue)
                End If

            Next

            Dim expiresDateTime As DateTime
            If (expiresString <> "") Then
                Try
                    expiresDateTime = DateTime.Parse(expiresString)
                Catch ex As Exception
                    expiresDateTime = DateTime.MinValue
                End Try

            Else
                expiresDateTime = DateTime.MinValue
            End If

            If (String.IsNullOrEmpty(domainFromCookie)) Then
                domainFromCookie = host
            End If


            For Each pair As Object In cookiesValues
                Dim cookie As Cookie = New Cookie(pair.Key, pair.Value, path, domainFromCookie)
                cookie.Secure = secure
                cookie.HttpOnly = httpOnly
                cookie.Expires = expiresDateTime

                cookieCollection.Add(cookie)
            Next
        Next

        Return cookieCollection
    End Function

#Region "Obsolete"
    Private Function _ExploreHtml(ByRef docel As XmlElement)

        Dim ret As String
        Dim post_content As String
        Dim tag As String
        Dim el As XmlElement

        ret = ""
        post_content = ""
        tag = docel.Name.ToUpper

        'some tag start with "!" ?!?
        If Left(tag, 1) = "!" Then
            Debug.WriteLine("")
        End If

        'can be there some tag starting with "/" or with empty name?
        If Left(tag, 1) = "/" Or tag = "" Then
            Return Nothing
        End If

        'If Me.p_exScript And tag = "SCRIPT" Then
        ' Return Nothing
        'End If

        Try

            'create tag element
            ret = ret & "<" & tag

            'add always id attribute when exists
            Dim tmp As String
            tmp = docel.GetAttribute("id")
            If tmp <> "" Then
                ret = ret & " id=""" & tmp & """"
            End If

            'in this section save different attributes per tag name
            If tag = "IMG" Then
                tmp = docel.GetAttribute("src")
                If tmp <> "" Then
                    'old: ret = ret & " src=""" & System.Web.HttpUtility.UrlEncode(tmp) & """"
                    ret = ret & " src=""" & tmp & """"
                End If
                tmp = docel.GetAttribute("alt")
                If tmp <> "" Then
                    Dim alt As String = docel.GetAttribute("alt")
                    alt = PurgeText(alt)
                    alt = alt.Replace("""", "'")
                    ret = ret & " alt=""" & alt & """"
                End If

            ElseIf tag = "A" Then
                tmp = docel.GetAttribute("href")
                If tmp <> "" Then
                    'old: ret = ret & " href=""" & System.Web.HttpUtility.UrlEncode(tmp) & """"
                    ret = ret & " href=""" & tmp & """"
                End If

            ElseIf tag = "DIV" Then
                If Not docel.Attributes("class") Is Nothing AndAlso docel.Attributes("class").Value <> "" Then
                    ret = ret & " class=""" & docel.Attributes("class").Value & """"
                End If

            ElseIf tag = "TABLE" Then
                If Not docel.Attributes("class") Is Nothing AndAlso docel.Attributes("class").Value <> "" Then
                    ret = ret & " class=""" & docel.Attributes("class").Value & """"
                End If

            ElseIf tag = "INPUT" Then
                tmp = docel.GetAttribute("class")
                If tmp <> "" Then
                    ret = ret & " class=""" & tmp & """"
                End If
                tmp = docel.GetAttribute("type")
                If tmp <> "" Then
                    ret = ret & " type=""" & tmp & """"
                End If
                tmp = docel.GetAttribute("value")
                If tmp <> "" Then
                    ret = ret & " value=""" & tmp & """"
                End If
                tmp = docel.GetAttribute("name")
                If tmp <> "" Then
                    ret = ret & " name=""" & tmp & """"
                End If

            ElseIf tag = "TD" Then
                If docel.GetAttribute("width") <> "" Then
                    ret = ret & " width=""" & docel.GetAttribute("width") & """"
                End If
                If Not docel.Attributes("class") Is Nothing AndAlso docel.Attributes("class").Value <> "" Then
                    ret = ret & " class=""" & docel.Attributes("class").Value & """"
                End If

            ElseIf tag = "META" Then
                If docel.GetAttribute("content") <> "" Then
                    'metto il content in un sotto-tag
                    post_content = docel.GetAttribute("content")
                End If

            Else
                tmp = docel.GetAttribute("class")
                If tmp <> "" Then
                    ret = ret & " class=""" & tmp & """"
                End If

            End If

            ret = ret & ">"

            If post_content <> "" Then
                post_content = Replace(post_content, Chr(0), "")
                post_content = Replace(post_content, Chr(19), "")
                post_content = Replace(post_content, Chr(26), "")
                ret = ret & "<CONTENT>"
                ret = ret & "<![CDATA[" & post_content & "]]>"
                ret = ret & "</CONTENT>"
            End If
            '
            'Metto il testo in un tag TEXT e in un [CDATA]
            '
            'ci sono alcuni tag di cui è inutile prendere il testo
            'tag = tag
            If tag <> "TABLE" And _
               tag <> "TBODY" And _
               tag <> "TR" And _
               tag <> "UL" Then

                tmp = Trim(docel.InnerText)
                If tmp <> "" And tmp <> vbCrLf Then
                    tmp = PurgeText(tmp)
                    'tmp = Replace(tmp, Chr(0), "")
                    'tmp = Replace(tmp, Chr(19), "")
                    'tmp = Replace(tmp, Chr(26), "")
                    tmp = Replace(tmp, "]]>", ".].]>")
                    'tmp = Replace(tmp, Chr(
                    ret = ret & "<TEXT>"
                    ret = ret & "<![CDATA[" & tmp & "]]>"
                    ret = ret & "</TEXT>"
                End If
            End If
            '
            'Metto il testo HTML in un tag HTML e in un [CDATA]
            '
            'ci sono alcuni tag di cui è inutile prendere il testo
            'tag = UCase(tag)
            If tag <> "TABLE" And _
               tag <> "TBODY" And _
               tag <> "TR" And _
               tag <> "UL" Then

                tmp = Trim(docel.InnerXml)
                If tmp <> "" And tmp <> vbCrLf Then
                    tmp = PurgeText(tmp)
                    'tmp = Replace(tmp, Chr(0), "")
                    'tmp = Replace(tmp, Chr(19), "")
                    'tmp = Replace(tmp, Chr(26), "")
                    tmp = Replace(tmp, "]]>", ".].]>")

                    ret = ret & "<HTML>"
                    ret = ret & "<![CDATA[" & tmp & "]]>"
                    ret = ret & "</HTML>"
                End If
            End If


            '
            ' Continuo a navigare nei figli ricorsivamente
            '
            If docel.ChildNodes.Count > 0 Then
                For Each n As XmlNode In docel.ChildNodes

                    If Not TypeOf n Is Xml.XmlText Then
                        el = n
                        'Escludo i tag che non voglio
                        If Left(tag, 1) = "!" Then
                            'Debug.WriteLine("")
                        End If
                        If el.Name.Contains(":") Then
                            Continue For
                        End If

                        ret = ret & _ExploreHtml(el)

                    End If

                Next
            End If

            ret = ret & "</" & tag & ">"

        Catch ex As Exception
            Throw ex
        End Try

        _ExploreHtml = ret

        el = Nothing

    End Function
#End Region

End Class

