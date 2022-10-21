using System;

namespace GoogleSheetsManager;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class SheetFieldAttribute : Attribute
{
    internal readonly string Title;
    internal readonly string? Format;

    public SheetFieldAttribute(string title, string? format = null)
    {
        Title = title;
        Format = format;
    }
}