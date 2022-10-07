using System.ComponentModel.DataAnnotations;
// ReSharper disable NullableWarningSuppressionIsUsed

namespace GoogleSheetsManager.Tests;

internal sealed class TestInstance
{
    [Required]
    [SheetField("RequiredBoolTitle")]
    public bool Bool;

    [SheetField("NullableIntTitle")]
    public int? Int;

    [Required]
    [SheetField("RequiredStringTitle")]
    public string String1 = null!;

    [SheetField("NullableStringTitle")]
    public string? String2;
}