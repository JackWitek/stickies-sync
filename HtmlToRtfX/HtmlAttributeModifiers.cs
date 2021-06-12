using System;
using System.Collections.Generic;
using X.HtmlToRtfConverter.Css;
using X.HtmlToRtfConverter.Rtf;

namespace X.HtmlToRtfConverter
{
    internal static class HtmlAttributeModifiers
    {
        //These are confirmed working
        private const string Color = "color";
        private const string TextIndent = "text-indent"; //creates Tab
        private const string Display = "display"; //Creates tab if "inline block"

        //These are not confirmed
        private const string BackgroundColor = "background-color";
        private const string TextAlign = "text-align";
        private const string FontSize = "font-size";
        private const string FontFamily = "font-family";

        // All others are ignored (uncomment line throw new NotSupportedException() to get failure)



        private static readonly Dictionary<string, Action<RtfDocumentBuilder, string>> AttributeModifiers = new Dictionary<string, Action<RtfDocumentBuilder, string>>
        {
            { Color, (builder, value) => builder.ForegroundColor(value) },
            { BackgroundColor, (builder, value) => builder.BackgroundColor(value) },
            { TextAlign, (builder, value) =>
                {
                    if (Enum.TryParse(value.Trim().FirstCharToUpper(), out HorizontalAlignment horizontalAlignment))
                    {
                        builder.HorizontalAlignment(horizontalAlignment);
                    }
                }
            },
            { FontSize, (builder, value) => builder.FontSize(new CssSize(value).ToPoints()) },
            { FontFamily, (builder, value) => builder.FontFamily(value) },
            { TextIndent, (builder, value) => builder.TextIndent(new CssSize(value).ToPoints()) },
            { Display, (builder, value) => builder.Display(value) },
        };

        public static void ApplyAttribute(string name, string value, RtfDocumentBuilder builder)
        {
            //This checks for attributes (font-size)
            if (!AttributeModifiers.ContainsKey(name))
            {
                System.Diagnostics.Debug.WriteLine("Can NOT process attribute " + name);

                return;
                //throw new NotSupportedException();
            }
            System.Diagnostics.Debug.WriteLine("Can process " + name);

            AttributeModifiers[name].Invoke(builder, value);
        }
    }
}