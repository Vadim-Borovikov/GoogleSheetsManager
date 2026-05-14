using GryphonUtilities.Time;
using System.ComponentModel.DataAnnotations;
// ReSharper disable NullableWarningSuppressionIsUsed

namespace GoogleSheetsManager.Tests;

internal sealed class TestInstance
{
    [SheetField("NullableIntTitle")]
    public int? Int;

    [Required]
    [SheetField("RequiredBoolTitle")]
    public bool Bool;

    [SheetField("NullableStringTitle")]
    public string? String1;

    [Required]
    [SheetField("RequiredStringTitle")]
    public string String2 = null!;

    [SheetField("NullableDateTimeFullTitle", "{0:dd.MM.yyyy HH:mm:ss}")]
    public DateTimeFull? DateTimeFull;
}