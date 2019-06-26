Imports System.Xml
Imports GekoProject

Module Module1

    Sub Main()

        Dim url As String
        Dim data As String
        Dim p As gkParser
        Dim doc As XmlDocument

        p = New gkParser()

        'p.HtmlParsingDoneEvent += new HtmlParsingDoneEventHandler(p_HtmlParsingDoneEvent);

        'GET
        url = "http://www.gekoproject.com/"
        p.Navigate(url)

        Dim l As XmlNodeList

        Console.WriteLine("Links:")
        Debug.Print(p.Document.NamespaceURI)

        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(p.Document.NameTable)
        nsmgr.AddNamespace("ns", "http://www.w3.org/1999/xhtml")
        l = p.Document.SelectNodes("//ns:a", nsmgr)
        For Each n As XmlNode In l
            Console.WriteLine(" " & n.Name & ":" & n.Attributes("href").Value)
        Next

        Console.WriteLine("Images:")
        l = p.Document.SelectNodes("//ns:img", nsmgr)
        For Each n As XmlNode In l
            If Not n.Attributes("src") Is Nothing Then
                Console.WriteLine(" " & n.Attributes("src").Value)
            End If
        Next

        'POST
        url = "http://testscrape.gekoproject.com/test_post.php"
        data = "name=francesco"
        p.Navigate(url, data)
        doc = p.Document
        Console.WriteLine(doc.InnerXml)

        Console.ReadLine()


    End Sub

End Module
