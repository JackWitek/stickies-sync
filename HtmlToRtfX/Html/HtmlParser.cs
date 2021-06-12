using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

using X.HtmlToRtfConverter.Html.Dom;

namespace X.HtmlToRtfConverter.Html
{
    public class HtmlParser
    {
        //------Entry Point For Parser -------//
        public static IEnumerable<HtmlDomEntity> Parse(string html)
        {
            //Get possible "tokens" to parse
            var tokenizer = new HtmlTokenizer();

            //Each char in html gets converted into token
            var tokens = tokenizer.Tokenize(html);

            //Chars already set at this point
            using (var enumerator = tokens.GetEnumerator())
            {
                while (enumerator.MoveNext() && enumerator.Current != null)
                {
                    var entity = Parse(enumerator);

                    Debug.WriteLine("Found Entity: " + entity.GetHtml());

                    if (entity != null)
                        yield return entity;
                }
            }
        }
        
        private static HtmlDomEntity Parse(IEnumerator<HtmlToken> enumerator)
        {
            if (enumerator.Current.TokenType == HtmlTokenType.CommentStart)
                return ParseComment(enumerator);
            if (enumerator.Current.TokenType == HtmlTokenType.ElementOpen)
                return ParseElement(enumerator);
            if (enumerator.Current.TokenType == HtmlTokenType.Text)
                return new HtmlText { Text = enumerator.Current.Value };
            return null;
        }

        //Figuring out what the element is
        private static HtmlDomEntity ParseElement(IEnumerator<HtmlToken> enumerator)
        {
            var isElementOpen = true;
            var element = new HtmlElement();
            var index = -1;
            while (enumerator.MoveNext() &&
                   enumerator.Current != null &&
                   enumerator.Current.TokenType != HtmlTokenType.ElementFinish &&
                   enumerator.Current.TokenType != HtmlTokenType.ElementInlineFinish)
            {
                if (string.IsNullOrEmpty(element.Name))
                    element.Name = ParseElementName(enumerator, out index);
                if (enumerator.Current.TokenType == HtmlTokenType.ElementClose)
                    isElementOpen = false;
                else if (isElementOpen)
                    element.Attributes.AddRange(ParseAttribute(enumerator, index));
                else
                {
                    var entity = Parse(enumerator);
                    if (entity != null)
                        element.Children.Add(entity);
                }

                index = 0;
            }

            if (enumerator.Current?.TokenType == HtmlTokenType.ElementFinish)
            {
                while (enumerator.Current.TokenType != HtmlTokenType.ElementClose)
                    enumerator.MoveNext();
            }

           // Debug.WriteLine("Have element group: " + element.GetHtml());
            

            return element;
        }
        private static HtmlComment ParseComment(IEnumerator<HtmlToken> enumerator)
        {
            var comment = new HtmlComment();
            while (enumerator.MoveNext() &&
                   enumerator.Current != null &&
                   enumerator.Current.TokenType != HtmlTokenType.CommentEnd)
            {
                if (enumerator.Current.TokenType == HtmlTokenType.Text)
                    comment.Value += enumerator.Current.Value;
            }
            return comment;
        }
        private static string ParseElementName(IEnumerator<HtmlToken> enumerator, out int index)
        {
            var elementRegex = new Regex("[a-zA-Z][a-zA-Z0-9]*");
            var matches = elementRegex.Matches(enumerator.Current?.Value);

            if (matches.Count > 0)
            {
                index = matches[0].Index + matches[0].Value.Length;
                return matches[0].Value;
            }
            index = -1;
            return null;
        }

        private static IEnumerable<HtmlAttribute> ParseAttribute(IEnumerator<HtmlToken> enumerator, int i)
        {
            var attributeStartRegex = new Regex("([a-zA-Z][a-zA-Z0-9]*=)(\"|')");
            //This will match style=' and style="

            var index = i;
            while (index < enumerator.Current.Value.Length)
            {
                var match = attributeStartRegex.Match(enumerator.Current.Value, index);
                if (match.Success)
                {
                    var attributeName = match.Value.Substring(0, match.Value.IndexOf("=", StringComparison.InvariantCulture));
                    var attributeStartOne = match.Value.IndexOf("\"", StringComparison.InvariantCulture);
                    var attributeStartTwo = match.Value.IndexOf("'", StringComparison.InvariantCulture);
                    var attributeStartType = attributeStartTwo > attributeStartOne ? '\'' : '"';
                    var attributeStart = Math.Max(attributeStartTwo, attributeStartOne);
                    var attributeValue = string.Empty;
                    var startIndex = match.Index + attributeStart + 1;
                    var isEscaped = false;
                    for (index = startIndex; index < enumerator.Current.Value.Length; index++)
                    {
                        if (!isEscaped && enumerator.Current.Value[index] == attributeStartType)
                            break;
                        if (!isEscaped && enumerator.Current.Value[index] == '\\')
                            isEscaped = true;
                        else
                        {
                            attributeValue += enumerator.Current.Value[index];
                            isEscaped = false;
                        }

                    }
                    yield return new HtmlAttribute
                    {
                        Name = attributeName,
                        Value = attributeValue
                    };
                }
                else break;
            }
        }
    }
}