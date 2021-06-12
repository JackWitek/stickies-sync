using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using X.HtmlToRtfConverter.Html;
using X.HtmlToRtfConverter.Html.Dom;
using X.HtmlToRtfConverter.Rtf;

namespace X.HtmlToRtfConverter
{
    public class RtfConverter
    {
        private readonly RtfConverterSubject _subject;

        public RtfConverter()
            : this(new RtfConverterSubject()) { }
        internal RtfConverter(RtfConverterSubject subject)
            => _subject = subject;

        //-----Entry Point-----//
        public string Convert(string html)
        {
            //Get rid of newlines. They cause trouble
            html = Regex.Replace(html, @"\r\n?|\n", "");
            html = Regex.Replace(html, @">(\s*)<li", "><li");

            //Change <br> into <br />  (parsing <br> is wonky, not sure about <br></br>)
            //One day set the parser to handle just <br>
            html = html.Replace(@"<br>", @"<br />");

            Debug.WriteLine("Starting HTML to RTF conversion of:");
            Debug.WriteLine(html);

            //Full HTML structure/contents are defined right away
            //Make sure this is correct when debugging!
            IEnumerable<HtmlDomEntity> dom = HtmlParser.Parse(html);

            Debug.WriteLine("Your dom is: ");
            Debug.WriteLine(dom.GetHtml());
            Debug.WriteLine("Your dom is done");


            //printDom(dom);

            return Convert(dom);
        }

        //For printing out the dom (debugging only)
        public void printDom(IEnumerable<HtmlDomEntity> dom)
        {
            //IEnumerable doesn't have count (IEnumerable uses lazy evaluation and can be more efficient)
            //Lets get count manually when debugging
            int count = 0;
            foreach (var htmlDomEntity in dom)
            {
                count++;
            }
            Debug.WriteLine("Your dom has " + count + " root elements" );
            Debug.WriteLine("---------------------------------------------");


            //Going through list of <HtmlDomEntity> in dom
            count = 1;
            foreach (var htmlDomEntity in dom) { 

                if (htmlDomEntity is HtmlElement element)
                {
                    Debug.WriteLine("Root entity " + count + " is element: " + element.Name);
                    Debug.WriteLine("It has " + element.Children.Count + " children ");

                    Debug.WriteLine("Root entity " + count + " has tree:");
                    Debug.WriteLine(htmlDomEntity.GetHtml());

                }
                else if (htmlDomEntity is HtmlText text)
                {
                    Debug.WriteLine("Root entity " + count + " is text: " + text.Text);
                }
                else
                {
                    Debug.WriteLine("Root entity " + count + " is not an element or text");
                }

                Debug.WriteLine("---------------------------------------------");

            count++;
            }
        }

        public string Convert(HtmlDomEntity entity)
            => Convert(new[] {entity});

        //This converts each HTML "element"
        public string Convert(IEnumerable<HtmlDomEntity> dom)
            => new RtfDocumentBuilder(_subject)
                .Html(dom)
                .Build();
    }
}