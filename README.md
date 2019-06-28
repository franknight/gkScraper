# gkScraper
gkScraper is a web scraping library for .NET Framework written in VB.NET.

It can navigate programmatically on web sites looking for links and information and collecting them.
Data are located using XPath syntax and extracted as node, HTML or text.

<b>Navigation features include:</b>
<ul>
<li>GET/POST requests subission, </li>
<li>http/https support, </li>
<li>cookies support, </li>
<li>redirection response, </li>
<li>custom headers attributes, </li>
<li>multipart form data submission, </li>
<li>ftp protocol helper methods </li>
</ul>

<b>Data scraping features:</b>
<ul>
<li>Absolute and relative XPath searches, </li>
<li>Find single node o nodes collection,</li>
<li>HTML o text extraction,</li>
<li>Parsing and cleaning functions,</li>
<li>include or exclude html tag and Attributes</li>
</ul>

<b>Debugging and Tracing</b>
To enable debug features there is a .Debug property in gkScrapeBot class.
Setting this property true, every navigation the parser write down (and overwrite) both the original html document received from host and the parsed XML document.
These files are located in the running process folder (<i>"debug.html"</i>, <i>"parser_orig.xml"</i>).

The libary implements also logging and tracing functionalities making use of NLOG library.
Setting the .Debug propery enable Debug logging level.
Setting the .Trace propery enable Trace logging level.

<h2>Test Scraping</h2>
This solution contains also a test project.
It scrapes products data from a demo e-commerce site and write results to a XML file and insert them into an MS Access Database.
Products are visible only to logged users.
The Test Project firstly securely log in and after collect data from the first page of displayed products.

<b>Note</b>: The library contains an helper class to insert data into a MS Access Database.
In order to use MS Access features you need to install the Microsoft.ACE.OLEDB.16.0 provider distributed with MS Access o Runtime library.
If you don't want make use of this features you can comment lines without affecting the scraping.








