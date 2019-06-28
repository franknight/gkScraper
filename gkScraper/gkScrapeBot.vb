Imports System.Xml
Imports System.IO
Imports System.Net
Imports System.Text
Imports NLog
Imports NLog.Config

Public Class gkScrapeBot

    Private p_P As gkParser
    Private p_cXML_text As String
    Private p_cXML As XmlDocument
    Private p_debug As Boolean
    Private p_trace As Boolean
    Private p_globalWaitTime As Integer
    Private p_currentUrl As String
    Private p_exScript As Boolean
    Private p_RandomWaitEnabled As Boolean
    Private LogMgr As NLog.Logger

    Public Sub New()

        Dim conf As LoggingConfiguration
        conf = NLog.LogManager.Configuration
        If conf Is Nothing Then
            conf = New LoggingConfiguration
            NLog.LogManager.Configuration = conf
        End If
        'LogManager.Configuration = conf
        conf.AddTarget("gkscraper_log_file", New Targets.FileTarget With {.FileName = "gkscraper.log"})
        conf.AddTarget("gkscraper_console", New Targets.ConsoleTarget With {.Layout = "${date:format=yyyyMMdd HH\:mm\:ss} ${logger} ${level} ${message}"})
        conf.LoggingRules.Add(New LoggingRule("gkscraper*", LogLevel.Debug, conf.FindTargetByName("gkscraper_log_file")))
        conf.LoggingRules.Add(New LoggingRule("gkscraper*", LogLevel.Debug, conf.FindTargetByName("gkscraper_console")))
        'LogManager.ReconfigExistingLoggers()

        'Dim l As Logger = LogManager.GetLogger("gkscraper")
        'Dim conf As Config.LoggingConfiguration = New LoggingConfiguration
        'LogManager.Configuration = conf
        'LogManager.Configuration.AddTarget("f1", New Targets.FileTarget With {.FileName = "batch.log"})
        'LogManager.Configuration.AddTarget("c", New Targets.ConsoleTarget With {.Layout = "${longdate} ${callsite} ${level} ${message}"})
        'LogManager.Configuration.LoggingRules.Add(New LoggingRule("*", NLog.LogLevel.Debug, LogManager.Configuration.FindTargetByName("f1")))
        'LogManager.Configuration.LoggingRules.Add(New LoggingRule("*", NLog.LogLevel.Debug, LogManager.Configuration.FindTargetByName("c")))
        'LogManager.ReconfigExistingLoggers()
        ''LogManager.Configuration.Reload()

        Me.LogMgr = LogManager.GetLogger("gkscraper")
        Me.LogMgr.Info("instance gkBot created.")
        p_debug = False
        p_P = New gkParser(LogMgr)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        p_cXML = Nothing
    End Sub

    ''' <summary>
    ''' Wait time expressed in milliseconds.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RandomWaitEnabled() As Boolean
        Get
            Return (Me.p_RandomWaitEnabled)
        End Get
        Set(ByVal value As Boolean)
            Me.p_RandomWaitEnabled = value
        End Set
    End Property
    Public Property GlobalWaitTime() As Integer
        Get
            Return (Me.p_globalWaitTime)
        End Get
        Set(ByVal value As Integer)
            Me.p_globalWaitTime = value
        End Set
    End Property

    Public Property Debug() As Boolean
        Get
            Return p_debug
        End Get
        Set(ByVal value As Boolean)
            p_debug = value
            p_RefreshLogLevel()
        End Set
    End Property
    Public Property Trace() As Boolean
        Get
            Return p_trace
        End Get
        Set(ByVal value As Boolean)
            p_trace = value
            p_RefreshLogLevel()
        End Set
    End Property
    Private Sub p_RefreshLogLevel()
        For Each l As LoggingRule In LogManager.Configuration.LoggingRules
            l.EnableLoggingForLevel(LogLevel.Fatal)
            l.EnableLoggingForLevel(LogLevel.Error)
            l.EnableLoggingForLevel(LogLevel.Warn)
            l.EnableLoggingForLevel(LogLevel.Info)
            If p_debug Then
                l.EnableLoggingForLevel(LogLevel.Debug)
            Else
                l.DisableLoggingForLevel(LogLevel.Debug)
            End If
            If p_trace Then
                l.EnableLoggingForLevel(LogLevel.Trace)
            Else
                l.DisableLoggingForLevel(LogLevel.Trace)
            End If

        Next
        LogManager.ReconfigExistingLoggers()
    End Sub

    Public ReadOnly Property CurrentURL() As String
        Get
            Return p_currentUrl
        End Get
    End Property

    Public ReadOnly Property Parser() As gkParser
        Get
            Return p_P
        End Get
    End Property

    'GET CON PARAMETRI
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Navigate(ByVal url As String)
        p_Navigate(url, Nothing, Nothing, Nothing)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Navigate(ByVal url As String, ByVal subel As String)
        p_Navigate(url, Nothing, subel, Nothing)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Navigate(ByVal url As String, ByVal subel As String, ByVal wait As Integer)
        p_Navigate(url, Nothing, subel, wait)
    End Sub

    'POST CON DATA STRING
    Public Sub Post(ByVal url As String, ByVal postData As String, ByVal customHeaders As gkHeaderCollection)
        p_Post(url, postData, Nothing, Nothing, customHeaders)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Post(ByVal url As String, ByVal postData As String)
        p_Post(url, postData, Nothing, Nothing, Nothing)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Post(ByVal url As String, ByVal postData As String, ByVal subel As String)
        p_Post(url, postData, subel, Nothing, Nothing)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="url">destination url</param>
    ''' <remarks></remarks>
    Public Sub Post(ByVal url As String, ByVal postData As String, ByVal subel As String, ByVal wait As Integer)
        p_Post(url, postData, subel, wait, Nothing)
    End Sub

    'POST CON DATA MULTIPART
    Public Sub Post(ByVal url As String, ByVal postData As gkMultipartFormData)
        p_Post(url, postData, Nothing, Nothing, Nothing)
    End Sub
    Public Sub Post(ByVal url As String, ByVal postData As gkMultipartFormData, ByVal subel As String)
        p_Post(url, postData, subel, Nothing, Nothing)
    End Sub
    Public Sub Post(ByVal url As String, ByVal postData As gkMultipartFormData, ByVal subel As String, ByVal wait As Integer)
        p_Post(url, postData, subel, wait, Nothing)
    End Sub

    Private Sub p_Post(ByVal url As String, ByVal postData As Object, ByVal subel As String, ByVal wait As Integer, ByVal customHeaders As gkHeaderCollection)

        'Wait parameter specifies the MAXIMUM time in seconds that bot waits before navigate
        ' The real waiting time is computed randomly in a range from 0 to the value of "wait" parameter.
        ' If wait parameter is not specified, a global setting specifies a wait time. Global default is 0.
        Dim f As FileStream
        Dim fw As StreamWriter
        Dim e As XmlElement

        Pause(wait)

        p_P.CustomHeaders = customHeaders

        If TypeOf postData Is String Then
            Dim data As String
            data = postData
            p_P.Post(url, data)
        ElseIf TypeOf postData Is gkMultipartFormData Then
            Dim data As gkMultipartFormData
            data = postData
            p_P.Post(url, data)
        End If

        p_currentUrl = url

        e = p_P.Document.DocumentElement
        If e Is Nothing Then
            Err.Raise(vbObjectError, "Navigate", "DocumentElement not set.")
        End If

        If subel <> "" Then
            e = p_P.Document.GetElementById(subel)
        End If
        If e Is Nothing Then
            Err.Raise(vbObjectError, "Navigate", "Subelement not found.")
        End If

        'if text is json
        If p_P.ContentType = "application/json" Then
            p_cXML = New XmlDocument()
            p_cXML.LoadXml(e.OuterXml)

            If p_debug Then
                f = New FileStream("file.xml", FileMode.Create)
                fw = New StreamWriter(f)
                fw.WriteLine(p_cXML.OuterXml)
                fw.Flush()
                f.Close()
            End If

        Else
            'HTML
            If p_debug Then
                f = New FileStream("parser_orig.xml", FileMode.Create)
                fw = New StreamWriter(f)
                fw.Write(e.OuterXml)
                fw.Flush()
                f.Close()
            End If

            p_cXML = p_P.Document

        End If


    End Sub

    Private Sub p_Navigate(ByVal url As String, ByVal Parameters As String, ByVal subel As String, ByVal wait As Integer)
        'Wait parameter specifies the MAXIMUM time in seconds that bot waits before navigate
        ' The real waiting time is computed randomly in a range from 0 to the value of "wait" parameter.
        ' If wait parameter is not specified, a global setting specifies a wait time. Global default is 0.
        Dim f As FileStream
        Dim fw As StreamWriter
        Dim e As XmlElement

        Pause(wait)

        p_P.Navigate(url, Parameters)

        p_currentUrl = url

        e = p_P.Document.DocumentElement
        If e Is Nothing Then
            Err.Raise(vbObjectError, "Navigate", "DocumentElement not set.")
        End If

        If subel <> "" Then
            e = p_P.Document.GetElementById(subel)
        End If
        If e Is Nothing Then
            Err.Raise(vbObjectError, "Navigate", "Subelement not found.")
        End If

        'if text is json
        If p_P.ContentType = "application/json" Then
            p_cXML = New XmlDocument()
            p_cXML.LoadXml(e.OuterXml)

            If p_debug Then
                f = New FileStream("file.xml", FileMode.Create)
                fw = New StreamWriter(f)
                fw.WriteLine(p_cXML.OuterXml)
                fw.Flush()
                f.Close()
            End If

        Else
            'HTML
            If p_debug Then
                f = New FileStream("parser_orig.xml", FileMode.Create)
                fw = New StreamWriter(f)
                fw.Write(e.OuterXml)
                fw.Flush()
                f.Close()
            End If

            p_cXML = p_P.Document

        End If


    End Sub

    <DebuggerStepThroughAttribute()> _
    Public Function GetNode_byXpath(ByVal xpath As String, Optional ByRef relNode As XmlNode = Nothing, Optional ByVal Attrib As String = "") As XmlNode

        Dim n As XmlNode

        If relNode Is Nothing Then
            n = p_cXML.SelectSingleNode(xpath)
        Else
            n = relNode.SelectSingleNode(xpath)
        End If

        GetNode_byXpath = n

    End Function
    <DebuggerStepThroughAttribute()> _
    Public Function GetNodes_byXpath(ByVal xpath As String, Optional ByRef relNode As XmlNode = Nothing, Optional ByVal Attrib As String = "") As XmlNodeList

        Dim ns As XmlNodeList

        If relNode Is Nothing Then
            ns = p_cXML.SelectNodes(xpath)
        Else
            ns = relNode.SelectNodes(xpath)
        End If

        GetNodes_byXpath = ns

    End Function

    Public Function GetText_byXpath(ByVal xpath As String, Optional ByRef relNode As XmlNode = Nothing, Optional ByVal Attrib As String = "") As String

        Dim result As String
        Dim n As XmlNode
        Dim n_el As XmlElement

        result = ""

        If relNode Is Nothing Then
            n = p_cXML.SelectSingleNode(xpath)
        Else
            n = relNode.SelectSingleNode(xpath)
        End If

        If Not n Is Nothing Then
            If Attrib <> "" Then
                n_el = n
                result = n_el.GetAttribute(Attrib)
                If Attrib = "href" Or Attrib = "src" Then
                    result = System.Web.HttpUtility.UrlDecode(result)
                End If
            Else
                result = n.InnerText
            End If
        End If
        GetText_byXpath = result

        n = Nothing
        n_el = Nothing

    End Function
    Public Function GetValue_byXpath(ByVal xpath As String, Optional ByRef relNode As XmlNode = Nothing, Optional ByVal Attrib As String = "") As String

        Dim result As String
        Dim n As XmlNode
        Dim n_el As XmlElement

        result = ""

        If relNode Is Nothing Then
            n = p_cXML.SelectSingleNode(xpath)
        Else
            n = relNode.SelectSingleNode(xpath)
        End If

        If Not n Is Nothing Then
            If Attrib <> "" Then
                n_el = n
                result = n_el.GetAttribute(Attrib)
                If Attrib = "href" Or Attrib = "src" Then
                    result = System.Web.HttpUtility.UrlDecode(result)
                End If
            Else
                result = n.InnerText
            End If
        End If
        GetValue_byXpath = result

        n = Nothing
        n_el = Nothing

    End Function
    Public Function GetHtml_byXpath(ByVal xpath As String, Optional ByRef relNode As XmlNode = Nothing) As String

        Dim result As String
        Dim n As XmlNode
        Dim txt_node As XmlNode
        Dim n_el As XmlElement

        result = ""
        On Error GoTo Errore

        If relNode Is Nothing Then
            n = p_cXML.SelectSingleNode(xpath)
        Else
            n = relNode.SelectSingleNode(xpath)
        End If
        If Not n Is Nothing Then
            result = n.InnerXml
        End If
        GetHtml_byXpath = result

        n = Nothing
        n_el = Nothing

        Exit Function

Errore:
        n = Nothing
        n_el = Nothing

    End Function

    Public Function FindHtml(ByVal Find As String) As Integer
        Dim result As String
        Dim n As XmlNode

        result = ""

        n = p_cXML.DocumentElement

        result = n.InnerXml

        If result.Contains(Find) Then
            Return result.IndexOf(Find)
        End If

        Return 0

    End Function
    Public Function FindText(ByVal Find As String) As Integer
        Dim result As String
        Dim n As XmlNode

        result = ""

        n = p_cXML.DocumentElement

        result = n.InnerText

        If result.Contains(Find) Then
            Return result.IndexOf(Find)
        End If

        Return 0

    End Function

    Public Sub Download(ByVal url As String, ByVal fullpath As String)
        Me.p_P.Download(url, fullpath)
    End Sub
    Public Function DownloadToString(ByVal url As String) As String

        Return Me.p_P.DownloadToString(url)

    End Function

    Public Shared Function TextExtractSingle(ByVal input As String, ByVal pattern As String) As String
        Dim mc As System.Text.RegularExpressions.MatchCollection
        mc = System.Text.RegularExpressions.Regex.Matches(input, pattern, RegularExpressions.RegexOptions.Singleline Or RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        If mc.Count > 0 AndAlso mc(0).Groups.Count > 1 Then
            Return mc(0).Groups(1).Value
        End If
        Return String.Empty
    End Function
    Public Shared Function TextExtractAll(ByVal input As String, ByVal pattern As String) As String()
        Dim mc As System.Text.RegularExpressions.MatchCollection
        mc = System.Text.RegularExpressions.Regex.Matches(input, pattern)

        If mc.Count > 0 AndAlso mc(0).Groups.Count > 1 Then
            Dim a As New ArrayList
            For i As Integer = 1 To mc(0).Groups.Count - 1
                Dim g As System.Text.RegularExpressions.Group
                g = mc(0).Groups(i)
                a.Add(Trim(g.Value))
            Next
            Return a.ToArray(GetType(String))
        End If

        Return New String() {input}

    End Function

    Public Shared Function TextRemoveTag(ByVal text As String, ByVal tag As String) As String

        Dim ret As String
        Dim pos1 As Integer
        Dim pos2 As Integer
        Dim start_tag As String
        Dim start_tag_1 As Integer
        Dim start_tag_2 As Integer
        Dim end_tag As String
        Dim end_tag_1 As Integer
        Dim end_tag_2 As Integer

        ret = text
        'cerco l'inizio del tag
        start_tag = "<" & tag

        While start_tag_1 > -1

            pos1 = ret.IndexOf(start_tag)
            If pos1 > -1 Then
                'cerco la chiusura dello start_tag
                pos2 = ret.IndexOf(">", pos1)
                If pos2 > -1 Then
                    'trovato tutto lo start tag
                    start_tag_1 = pos1
                    start_tag_2 = pos2

                    'cerco il tag di chiusura
                    end_tag = "</" & tag & ">"
                    pos1 = ret.IndexOf(end_tag, start_tag_2)
                    If pos1 > -1 Then
                        'trovato end tag
                        end_tag_1 = pos1
                        end_tag_2 = pos1 + end_tag.Length

                        'rimuovo tutto il testo contenuto nel tag
                        ret = ret.Remove(start_tag_1, end_tag_2 - start_tag_1)
                    Else
                        'non ha chiusura
                        'rimuovo solo lo start tag
                        ret = ret.Remove(start_tag_1, start_tag_2 - start_tag_1 + 1)
                    End If

                Else
                    '?!?
                    Throw New Exception
                End If
            Else
                start_tag_1 = -1
            End If


        End While


        Return ret

    End Function

    Public Shared Function GetAlphaPart(ByVal value As String) As String
        Dim out As String
        out = ""
        Dim digit As String = "0123456789"
        'value = value.ToLower
        For Each c As Char In value.ToCharArray
            If Not digit.Contains(c) Then
                out &= c
            End If
        Next
        Return out
    End Function
    Public Shared Function GetNumberPart(ByVal value As String) As String
        'get systems decimal separator 
        Dim nfi As New System.Globalization.NumberFormatInfo
        Dim decsep = nfi.NumberDecimalSeparator
        Return GetNumberPart(value, decsep)
    End Function
    Public Shared Function GetNumberPart(ByVal value As String, ByVal decimalSep As String) As Double
        Dim out As String
        out = ""
        Dim nos As String = "0123456789" & decimalSep
        For Each c As Char In value.ToCharArray
            If nos.Contains(c) Then
                out &= c
            End If
        Next
        Dim nfi As System.Globalization.NumberFormatInfo
        nfi = System.Globalization.CultureInfo.CurrentCulture.NumberFormat
        Dim decsep = nfi.NumberDecimalSeparator
        Dim thsep = nfi.NumberGroupSeparator

        If decimalSep <> decsep Then
            out = out.Replace(decimalSep, decsep)
        End If
        'Return CDbl(out)
        Return out

    End Function
    Public Shared Function FriendLeft(ByVal text As String, ByVal lenght As Integer) As String
        Dim ret As String
        ret = String.Empty
        text = text.Trim
        If text.Length > 0 Then
            If text.Length > lenght Then
                ret = text.Substring(0, lenght - 3) & "..."
            Else
                ret = text
            End If
        End If
        Return ret
    End Function

    Public Shared Function FtpPutFile(ByVal server As String, ByVal user As String, ByVal password As String, ByVal localFile As String, ByVal remoteFile As String)
        Dim ftp As New gkFtp(server, user, password)
        ftp.upload(localFile, remoteFile)
        Return Nothing
    End Function

    Public Shared Function FtpDownloadFile()
        Return Nothing
    End Function

    Private Sub Pause(ByVal wait As Integer)
        Dim nextValue = 0
        If wait > 0 Then
            Dim rnd = New Random()
            nextValue = rnd.Next(wait * 1000)
        End If
        If p_RandomWaitEnabled Then
            Dim rnd = New Random()
            nextValue = rnd.Next(p_globalWaitTime * 2)
        End If
        System.Threading.Thread.Sleep(p_globalWaitTime + nextValue)
    End Sub

    Public Shared Function URLEncode(ByVal input As String) As String
        Return System.Web.HttpUtility.UrlEncode(input)
    End Function
    Public Shared Function URLDecode(ByVal input As String) As String
        Return System.Web.HttpUtility.UrlDecode(input)
    End Function

End Class


Public Enum gkReqType
    [Get]
    [Post]
End Enum
