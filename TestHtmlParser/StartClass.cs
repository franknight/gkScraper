using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;

namespace gkHtmlParser
{
    class StartClass
    {
     
        static void Main(string[] args)
        {

            string url;
            string data;

            HtmlParser p;
            p = new HtmlParser();
            p.HtmlParsingDoneEvent += new HtmlParsingDoneEventHandler(p_HtmlParsingDoneEvent);

            //GET
            url = "http://www.gekoproject.com/";
            p.Navigate(url);

            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)p.Document;
            Console.WriteLine("Links:");
            foreach (mshtml.IHTMLElement e in doc.links)
            {
                mshtml.IHTMLAnchorElement a = (mshtml.IHTMLAnchorElement)e;
                if (a!=null)
                    Console.WriteLine(" " + a.href);
            }
            Console.WriteLine("Images:");
            foreach (mshtml.IHTMLElement e in doc.images)
            {
                mshtml.IHTMLImgElement i = (mshtml.IHTMLImgElement)e;
                if (i != null)
                    Console.WriteLine(" " + i.src);
            }

            //POST
            url = "http://testscrape.gekoproject.com/test_post.php";
            data = "name=francesco";
            p.Navigate(url, data);
            //doc = (mshtml.IHTMLDocument2)p.Document;
            Console.WriteLine(doc.body.innerHTML);

            Console.ReadLine();
        }

        static void p_HtmlParsingDoneEvent(object sender, HtmlParsingDoneEventArg e)
        {
            //This is a real asyncronous method, main process can terminate before this is completed.
            //Console.WriteLine(e.MSHTMDocument.documentElement.outerHTML);
        }



    }
}
