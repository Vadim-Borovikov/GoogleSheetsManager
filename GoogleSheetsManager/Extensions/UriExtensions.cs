using System;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Extensions;

[PublicAPI]
public static class UriExtensions
{
    public static string ToHyperlink(this Uri uri, string? caption = null)
    {
        if (string.IsNullOrWhiteSpace(caption))
        {
            caption = uri.AbsoluteUri;
        }
        return string.Format(HyperlinkFormat, uri.AbsoluteUri, caption);
    }

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}