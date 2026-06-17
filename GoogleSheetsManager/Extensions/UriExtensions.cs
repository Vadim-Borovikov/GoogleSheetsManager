using GryphonUtilities.Extensions;
using JetBrains.Annotations;
using System;

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
        return HyperlinkFormat.Format(uri.AbsoluteUri, caption);
    }

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}