# gkScraper
<p>gkScraper is a web scraping library for .NET Framework written in VB.NET.<br>

It can navigate programmatically on web sites looking for links and information and collecting them.<br>
Data are located using XPath syntax and extracted as node, HTML or text.</p>

<h3>Navigation features include:</h3>
<ul>
<li>GET/POST requests subission, </li>
<li>http/https support, </li>
<li>cookies support, </li>
<li>redirection response, </li>
<li>custom headers attributes, </li>
<li>multipart form data submission, </li>
<li>ftp protocol helper methods </li>
</ul>

<h3>Data scraping features:</h3>
<ul>
<li>Absolute and relative XPath searches, </li>
<li>Find single node o nodes collection,</li>
<li>HTML o text extraction,</li>
<li>Parsing and cleaning functions,</li>
<li>include or exclude html tag and Attributes</li>
</ul>

<h3>Debugging and Tracing</h3>
<p>To enable debug features there is a .Debug property in gkScrapeBot class.<br>
Setting this property true, every navigation the parser write down (and overwrite) both the original html document received from host and the parsed XML document.<br>
These files are located in the running process folder (<i>"debug.html"</i>, <i>"parser_orig.xml"</i>).<br>

The libary implements also logging and tracing functionalities making use of NLOG library.<br>
Setting the .Debug propery enable Debug logging level.<br>
Setting the .Trace propery enable Trace logging level.</p>

<h2>Test Scraping</h2>
<p>This solution contains also a test project.<br>
It scrapes products data from a demo e-commerce site and write results to a XML file and insert them into an MS Access Database.<br>
Products are visible only to logged users.<br>
The Test Project firstly securely log in and after collect data from the first page of displayed products.</p>

<h3>Note</h3>
<p>The library contains an helper class to insert data into a MS Access Database.<br>
In order to use MS Access features you need to install the Microsoft.ACE.OLEDB.16.0 provider distributed with MS Access o Runtime library. Be sure to target the right platform x86 or x64 according to the runtime installed.<br>
If you don't want make use of this features you can comment lines without affecting the scraping.<br>

</p>








