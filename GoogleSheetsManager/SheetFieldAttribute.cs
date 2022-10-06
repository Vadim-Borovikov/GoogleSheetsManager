using System;

namespace GoogleSheetsManager;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class SheetFieldAttribute : Attribute
{
    internal readonly string Title;

    public SheetFieldAttribute(string title) => Title = title;
}