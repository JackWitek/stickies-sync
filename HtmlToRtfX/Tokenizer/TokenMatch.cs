namespace X.HtmlToRtfConverter.Tokenizer
{
    public abstract class TokenMatch
    {
        public bool IsMatch { get; set; }
        public string Value { get; set; }
        public string RemainingText { get; set; }
    }
}
