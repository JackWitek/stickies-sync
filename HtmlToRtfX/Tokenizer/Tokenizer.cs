using System.Collections.Generic;

namespace X.HtmlToRtfConverter.Tokenizer
{
    public abstract class Tokenizer<TToken, TTokenMatch, TTokenDefinition>
        where TTokenMatch : TokenMatch, new()
        where TTokenDefinition : TokenDefinition<TTokenMatch>
    {
        protected readonly List<TTokenDefinition> TokenDefinitions = new List<TTokenDefinition>();

        public IEnumerable<TToken> Tokenize(string text)
        {
            var remainingText = text;

            while (!string.IsNullOrWhiteSpace(remainingText))
            {
                //Matches start of "text" to a type of token (Ex. Element Open, )
                var match = FindMatch(remainingText);


                if (match.IsMatch)
                {
                    yield return Construct(match);
                    remainingText = match.RemainingText;
                }
                else
                {
                    remainingText = remainingText.Substring(1);
                }
            }
        }

        protected abstract TToken Construct(TTokenMatch match);

        //Matches start of "text" to a type of token (Ex. Element Open, )
        private TTokenMatch FindMatch(string text)
        {
            foreach (var tokenDefinition in TokenDefinitions)
            {
                var match = tokenDefinition.Match(text);
                if (match.IsMatch)
                {
                    //System.Diagnostics.Debug.WriteLine("Matched " + text);
                    //Diagnostics.Debug.WriteLine(match.Value + " is a " + tokenDefinition.Regex);

                    return match;

                }
            }

            return new TTokenMatch { IsMatch = false };
        }
    }
}