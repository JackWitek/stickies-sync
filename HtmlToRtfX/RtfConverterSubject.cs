using System.Collections.Generic;
using X.HtmlToRtfConverter.Html.Dom;

namespace X.HtmlToRtfConverter
{
    public class RtfConverterSubject
    {
        public Dictionary<HtmlElementType, ElementSubject> ElementSubjects = new Dictionary<HtmlElementType, ElementSubject>();
        public string FontStyle;
    }
}