Imports System.Xml
Imports Gekoproject.gkScraper

Module MainModule

    Dim db As DBAccessHelper
    Dim bot As gkScrapeBot

    Sub Main()

        db = New DBAccessHelper
        db.ConnectDatabase("Provider = Microsoft.ACE.OLEDB.16.0;Data Source=.\ScrapeDB.mdb;Persist Security Info=False;")

        If bot Is Nothing Then
            bot = New gkScrapeBot
        End If

        bot.Debug = True
        bot.Trace = True

        bot.GlobalWaitTime = 1000

        Scrape()

    End Sub

    Private Sub Scrape()

        Dim ct As Integer

        If Login() Then
            ct = GetData()
            Console.WriteLine("Retrived " & ct & " article.")
        Else
            Throw New Exception("Some error occurred during login.")
        End If

    End Sub

    Private Function Login() As Boolean

        Dim url As String
        Dim data As String
        Dim mytext As String
        Dim token1 As String
        Dim token2 As String
        Const USER = "test"
        Const PASS = "test"

        'Navigate to homepage and get cookies. 
        url = "https://testscrape.gekoproject.com/index.php/author-login"
        bot.Navigate(url)

        'Then look for two parameters useful to login
        token1 = bot.GetText_byXpath("//DIV[@class='login']//INPUT[@type='hidden'][1]", , "value")
        token2 = bot.GetText_byXpath("//DIV[@class='login']//INPUT[@type='hidden'][2]", , "name")

        'Now login with username e pssword
        url = "https://testscrape.gekoproject.com/index.php/author-login?task=user.login"
        data = "username=" & USER & "&password=" & PASS & "&return=" & token1 & "&" & token2 & "=1"
        bot.Post(url, data)
        mytext = bot.GetText_byXpath("//DT[contains(.,'Registered Date')]/following-sibling::DD[1]")
        Console.WriteLine("User {0}, Registered Date: {1}", USER, mytext.Trim)

        Return True

    End Function

    Private Function GetData() As Integer

        Dim url As String
        Dim name As String
        Dim desc As String
        Dim price_str As String
        Dim price As Double
        Dim img_path As String

        'Get products details
        url = "https://testscrape.gekoproject.com/index.php/front-end-store"
        bot.Navigate(url)

        Dim ns As XmlNodeList = bot.GetNodes_byXpath("//DIV[@class='row']//DIV[contains(@class, 'product ')]")
        If ns.Count > 0 Then

            Dim writer As XmlWriter = Nothing

            ' Create an XmlWriterSettings object with the correct options. 
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
                settings.Indent = True
                settings.IndentChars = (ControlChars.Tab)
                settings.OmitXmlDeclaration = True

            writer = XmlWriter.Create("data.xml", settings)
            writer.WriteStartElement("products")

            For Each n As XmlNode In ns
                name = bot.GetText_byXpath(".//DIV[@class='vm-product-descr-container-1']/H2", n)
                desc = bot.GetText_byXpath(".//DIV[@class='vm-product-descr-container-1']/P", n)
                desc = gkScrapeBot.FriendLeft(desc, 50)
                img_path = bot.GetText_byXpath(".//DIV[@class='vm-product-media-container']//IMG", n, "src")
                price_str = bot.GetText_byXpath(".//DIV[contains(@class,'PricesalesPrice')]", n)
                If price_str <> "" Then
                    price = gkScrapeBot.GetNumberPart(price_str, ",")
                End If

                writer.WriteStartElement("product")

                writer.WriteElementString("name", name)
                writer.WriteElementString("description", desc)
                writer.WriteElementString("price", price)
                writer.WriteElementString("image", img_path)

                writer.WriteEndElement()

                'Insert data into DB
                db.CommantType = DBCommandTypes.INSERT
                db.Table = "Articles"
                db.Fields("Name") = name
                db.Fields("Description") = desc
                db.Fields("Price") = price
                Dim ra As Integer = db.Execute()
                If ra = 1 Then
                    Console.WriteLine("Inserted new article: {0}", name)
                End If

            Next

            writer.WriteEndElement()
            writer.Flush()
            writer.Close()

        End If

    End Function


End Module
