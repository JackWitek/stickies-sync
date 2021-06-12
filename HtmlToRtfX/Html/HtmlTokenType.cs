﻿namespace X.HtmlToRtfConverter.Html
{
    public enum HtmlTokenType
    {
        ElementOpen,
        ElementClose,
        ElementInlineFinish,
        ElementFinish,
        
        Text,
        NewLine,

        CommentStart,
        CommentEnd
    }
}
