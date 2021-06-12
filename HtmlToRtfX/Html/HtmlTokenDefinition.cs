using X.HtmlToRtfConverter.Tokenizer;

namespace X.HtmlToRtfConverter.Html
{
    public class HtmlTokenDefinition : TokenDefinition<HtmlTokenMatch>
    {
        public HtmlTokenType ReturnsToken { get; }

        public HtmlTokenDefinition(HtmlTokenType returnsToken, string regexPattern)
            : base(regexPattern)
        {
            ReturnsToken = returnsToken;
        }

        public override HtmlTokenMatch Match(string inputString)
        {
            var result = base.Match(inputString);
            if (result.IsMatch)
            {
                result.TokenType = ReturnsToken;
            }
            return result;
        }
    }
}
