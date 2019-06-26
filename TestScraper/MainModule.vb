'Imports mshtml
Imports System.Xml
Imports System.IO
Imports System.Data
Imports Gekoproject

Module MainModule

    Dim db As DBAccessHelper

    Dim bot As gkScrapeBot

    Sub Main()

        db = New DBAccessHelper
        db.ConnectDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=.\ScrapeDB.mdb;")

        If bot Is Nothing Then
            bot = New gkScrapeBot
        End If

        bot.Debug = True
        bot.Trace = True

        bot.GlobalWaitTime = 1000

        Try
            'TestPostJson()
            Scrape()

        Catch ex As Exception
            Console.WriteLine(ex.Message)

        End Try

        Console.ReadLine()
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
        url = "http://testscrape.gekoproject.com/index.php/author-login"
        bot.Navigate(url)

        'Then look for two parameters useful to login
        token1 = bot.GetText_byXpath("//DIV[@class='login']//INPUT[@type='hidden'][1]", , "value")
        token2 = bot.GetText_byXpath("//DIV[@class='login']//INPUT[@type='hidden'][2]", , "name")

        'Now login with username e pssword
        url = "http://testscrape.gekoproject.com/index.php/log-out?task=user.login/"
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

        'Get products details
        url = "http://testscrape.gekoproject.com/index.php/shop/categories"
        bot.Navigate(url)

        Dim ns As XmlNodeList = bot.GetNodes_byXpath("//DIV[@class='row']//DIV[contains(@class, 'product ')]")
        For Each n As XmlNode In ns
            name = bot.GetText_byXpath(".//DIV[@class='vm-product-descr-container-1']/H2", n)
            desc = bot.GetText_byXpath(".//DIV[@class='vm-product-descr-container-1']/P", n)
            desc = gkScrapeBot.FriendLeft(desc, 50)

            price_str = bot.GetText_byXpath(".//DIV[contains(@class,'PricesalesPrice')]", n)
            If price_str <> "" Then
                price = gkScrapeBot.GetNumberPart(price_str, ",")
            End If

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


    End Function

    Public Sub TestPostJson()

        Dim url As String
        Dim data As String
        Dim title As String
        Dim start As String

        url = "http://www.teatrodue.org/wp-admin/admin-ajax.php/"
        data = "action=get_events&readonly=true&categories=0&excluded=0&start=1424646000&end=1428271200"
        bot.Post(url, data)

        title = bot.GetValue_byXpath("//item[@type='object'][id='921']/title")
        start = bot.GetValue_byXpath("//item[@type='object'][id='921']/start")
        Console.WriteLine("Titolo {0}, Inizio: {1}", title, start)

    End Sub

End Module
