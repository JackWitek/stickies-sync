using System;

namespace X.HtmlToRtfConverter

{
	class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

			var html = @"
<p style='text-align: center; font-size: 12pt; font-family: Times New Roman'>center aligned</p>
<p style='text-align: left; font-size: 24pt; font-family: Arial'>left aligned</p>
<p style='text-align: right; font-size: 36pt; font-family: Helvetica'>right aligned</p>
<p style='text-align: justify; font-size: 48pt'>center aligned</p>

<div>
	<!--block-->
	Lorem ipsum dolor sit amet,
	<strong>consectetur</strong>
	adipiscing elit
	<del>Praesent lacus diam</del>
	, fermentum et venenatis quis, suscipit sed nisi. In pharetra sem eget orci posuere pretium.
	<em>Integer</em>
	non eros
	<strong>
		<em>scelerisque</em>
	</strong>
	, consequat lacus id, rutrum felis. Nulla elementum felis urna, at placerat arcu ultricies in.
</div>
<ul>
	<li>
		<!--block-->
		Proin elementum sollicitudin sodales.
	</li>
	<li>
		<!--block-->
		Nam id erat nec nibh dictum cursus.
	</li>
</ul>
<div>
	<!--block-->
	<br/>
</div>
<blockquote>
	<!--block-->
	<strong>asdasd</strong>
</blockquote>
<div>
	<!--block-->
	<br/>
</div>
<ol>
	<li>
		<!--block-->
		Proin elementum sollicitudin sodales.
	</li>
	<li>
		<!--block-->
		Nam id erat nec nibh dictum cursus.
	</li>
</ol>
<div>
	<!--block-->
	<br/>
</div>
<div>
	<!--block-->
	<br/>
</div>
<div>
	<!--block-->
	<br/>
</div>
<blockquote><!--block-->asd</blockquote>
<div>
	<!--block-->
	<br/>
</div>
<pre><!--block-->asdasd</pre>
<div style='background-color:green;'>
    Hello World
</div>
<blockquote>
	<!--block-->
	In et urna eros. Fusce molestie, orci vel laoreet tempus, sem justo blandit magna, at volutpat velit lacus id turpis.<br>Quisque malesuada sem at interdum congue. Aenean dapibus fermentum orci eu euismod.
</blockquote>
<div><!--block--><br></div>";



			var rtfConverter = new RtfConverter();
			var rtf = rtfConverter.Convert(html);

			Console.WriteLine(rtf);
		}
    }
}
