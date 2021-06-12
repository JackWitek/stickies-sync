using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using X.HtmlToRtfConverter.Html.Dom;
using X.HtmlToRtfConverter.Rtf;

namespace X.HtmlToRtfConverter
{
    //TODO: ¨multilevel lists (ordered, unordered); background color; paragraphs
    internal static class RtfDocumentBuilderExtensions
    {
        private const string StyleAttributeName = "style";

        //Elements that cause new paragraph
        private static readonly HtmlElementType[] ClearingElements =
        {
            HtmlElementType.P,
            HtmlElementType.Li,
            //HtmlElementType.Div,
            HtmlElementType.Blockquote,
            HtmlElementType.Pre,
            //HtmlElementType.Ol,
            //HtmlElementType.Ul
        };

        public static RtfDocumentBuilder Html(this RtfDocumentBuilder builder, IEnumerable<HtmlDomEntity> dom)
            => builder.HtmlInner(dom, true);

        //Is this where the magic happens?
        private static RtfDocumentBuilder HtmlInner(this RtfDocumentBuilder builder, IEnumerable<HtmlDomEntity> dom, bool isCleared)
        {

            //Debug.WriteLine("First object in dom: "+ dom.ToList().First().GetHtml());

            //This is where the type of object is selected
            foreach (var obj in dom)
            {


                if (obj is HtmlElement element)
                {
                    Debug.WriteLine("");
                    Debug.WriteLine("Have element: " + element.Name);

                    var name = ToFirstLetterUpper(element.Name);

                    Debug.WriteLine(element.Name + " has children: " + element.Children.Count);


                    if (!Enum.TryParse(name, out HtmlElementType elementType))
                        continue;

                    //Hit opening tag
                    if (elementType == HtmlElementType.Ol || elementType == HtmlElementType.Ul)
                    {
                        builder.Context.ListLevel++;
                        builder.Context.ListItem = 0;

                    }

                    if (elementType == HtmlElementType.Li)
                    {
                        builder.Context.ListItem++;

                    }


                    if (elementType == HtmlElementType.P)
                    {
                        builder.Context.inP = true; 
                    }

                    if(builder.Context.inP && elementType == HtmlElementType.Br)
                    {
                        builder.Context.skipPar = true;

                    }
                    else
                    {
                        builder.Context.skipPar = false;

                    }



                    Debug.WriteLine(element.Name + " has children: " +element.Children.Count);




                    //Process the element, and then its children
                    builder
                        .OpenContext()
                        .ApplyElementOpeningModifiers(elementType, isCleared)
                        .ApplyAttributeModifiers(element)
                            .OpenContext()
                            .ApplyConfigurationModifiers(builder.Subject, elementType)
                            .HtmlInner(element.Children, false)
                            .CloseContext()
                        .ApplyElementClosingModifiers(elementType)
                        .CloseContext();

                    //Hit closing tag
                    if (elementType == HtmlElementType.Ol || elementType == HtmlElementType.Ul)
                        builder.Context.ListLevel--;

                    //Hit closing tag
                    if (elementType == HtmlElementType.P)
                    {
                        builder.Context.inP = false; 

                    }

                    isCleared = true;
                }
                else if (obj is HtmlText text)
                {
                    var textValue= text.Text; //= RemoveSpacing(text.Text); 
                    Debug.WriteLine("");
                    Debug.WriteLine("Have text: " + text.Text);


                    if (string.IsNullOrWhiteSpace(text.Text))
                    {

                        //Debug.WriteLine("text: '" + text.Text + "' Is all whitespace");

                        // Gets rid of random spaces between elements
                       // textValue = "";
                    }

                    textValue = FixSpecial(textValue);

                    builder.Text(textValue);
                }
            }
            return builder;
        }
        private static RtfDocumentBuilder ApplyConfigurationModifiers(this RtfDocumentBuilder builder, RtfConverterSubject subject, HtmlElementType elementType)
        {
            if (!subject.ElementSubjects.ContainsKey(elementType))
                return builder;

            var configuration = subject.ElementSubjects[elementType];

            if (!configuration.ForegroundColor.IsEmpty)
                builder.ForegroundColor(configuration.ForegroundColor);
            if (!configuration.BackgroundColor.IsEmpty)
                builder.BackgroundColor(configuration.BackgroundColor);
            if (Math.Abs(configuration.FontSize - default(float)) < 0)
                builder.FontSize(configuration.FontSize);
            if (!string.IsNullOrEmpty(configuration.FontStyle))
                builder.FontFamily(configuration.FontStyle);

            builder.HorizontalAlignment(configuration.HorizontalAlignment);

            return builder;
        }
        private static RtfDocumentBuilder ApplyElementClosingModifiers(this RtfDocumentBuilder builder, HtmlElementType elementType)
        {
            if (ClearingElements.Contains(elementType) && !builder.Context.skipPar)
            {
                builder.Paragraph(elementType);

            }
            return builder;
        }
        private static RtfDocumentBuilder ApplyElementOpeningModifiers(this RtfDocumentBuilder builder, HtmlElementType elementType, bool isCleared)
        {
            //if (!isCleared && ClearingElements.Contains(elementType))
                //builder.Paragraph();
            if (elementType == HtmlElementType.Strong)
                builder.Bold();
            if (elementType == HtmlElementType.Em)
                builder.Italic();
            if (elementType == HtmlElementType.U)
                builder.Underline();
            if (elementType == HtmlElementType.S)
                builder.Striketrough();
            if (elementType == HtmlElementType.Del)
                builder.Striketrough();
            if (elementType == HtmlElementType.Ol)
            {
                Debug.WriteLine("Adding ordered list");

                builder.Context.IsOrderedList = true;
                var indent = builder.Context.ListLevel * 300;

                builder.Rtf(@"\pard{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}\fi-" + indent + @"\li" + indent);

            }
            if (elementType == HtmlElementType.Ul)
            {
                Debug.WriteLine("Adding UNordered list");
                builder.Context.IsOrderedList = false;
                var indent = builder.Context.ListLevel * 300;

                builder.Rtf(@"\pard{\*\pn\pnlvlblt\pnf2\pnindent0{\pntxtb\'B7}}\fi-" + indent + @"\li" + indent );

            }
            if (elementType == HtmlElementType.Li)
            {

                if (builder.Context.IsOrderedList)
                {
                    Debug.WriteLine("Adding ordered list item: " + builder.Context.ListItem);

                }
                else
                {
                    Debug.WriteLine("Adding unordered list item" );

                }

                string ordered = @"{\pntext\f0 "+ (builder.Context.ListItem)+ @".\tab}";
                string unordered = @"{\pntext\f2\'B7\tab}" ;

                var listItem = builder.Context.IsOrderedList ? ordered  : unordered;
                builder.Rtf(listItem); //TODO: move to rtfDocumentBuilder
                


            }
            if (elementType == HtmlElementType.Blockquote)
                builder
                    .ForegroundColor(Color.LightSlateGray)
                    .HorizontalAlignment(HorizontalAlignment.Center);
            if (elementType == HtmlElementType.Pre)
                builder.Rtf(@"\brdrt\brdrs\brdrb\brdrs\brdrl\brdrs\brdrr\brdrs\brdrw10\brsp20\brdrcf0"); //TODO: move to rtfDocumentBuilder
            if (elementType == HtmlElementType.Br)
                builder.NewLine();

            return builder;
        }
        private static RtfDocumentBuilder ApplyAttributeModifiers(this RtfDocumentBuilder builder, HtmlElement element)
        {

            var style = element.Attributes.FirstOrDefault(x => x.Name.ToLower() == StyleAttributeName);

            Debug.WriteLine("Checking for style attributes in element: " + element.Name);

            if (style != null)
            {
                Debug.WriteLine("Element has style attributes: " + style.Value);

                style.Value = style.Value.Replace("&quot;", "\"");

                var styleParts = style.Value.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                foreach (var stylePart in styleParts)
                {
                    var keyValue = stylePart.Split(":".ToCharArray());
                    var key = keyValue[0].ToLower();
                    var value = keyValue[1];

                    Debug.WriteLine("Have style attribute: " + key.Trim() + " with value: " + value);

                    HtmlAttributeModifiers.ApplyAttribute(key.Trim(), value, builder);
                }
            }
            return builder;
        }

        private static string ToFirstLetterUpper(string value)
            => new string(value.Select((x, i) => i == 0 ? char.ToUpper(x) : char.ToLower(x)).ToArray());
        private static string RemoveSpacing(string text)
        {
            text = text
                .Replace("\r", "")
                .Replace("\n", "")
                .Replace("\t", "");

            while (text.Contains("  "))
                text = text.Replace("  ", " ");
            while (text.StartsWith(" "))
                text = text.Substring(1);
            while (text.EndsWith(" "))
                text = text.Substring(0, text.Length - 1);

            return text;
        }

        private static string FixSpecial(string text)
        {
            Debug.WriteLine("Replacing special chars");

            text = text
                .Replace("&nbsp;", " ")
                .Replace("&amp;", "&")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&#39;", "'")
                .Replace("&quot;", "\"")
                .Replace("&ndash;", "–")
                .Replace(@"\", @"\\")
                .Replace("{", @"\{")
                .Replace("}", @"\}");

            return text;
        }
    }
}