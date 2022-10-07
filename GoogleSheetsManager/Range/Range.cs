namespace GoogleSheetsManager.Range;

internal sealed class Range
{
    public readonly Bound IntervalStart;
    public readonly Bound? IntervalEnd;

    private Range(Bound intervalStart, Bound? intervalEnd = null, string? sheet = null)
    {
        IntervalStart = intervalStart;
        IntervalEnd = intervalEnd;
        _sheet = sheet;
    }

    public override string ToString()
    {
        string? intervalEnd = IntervalEnd is null ? null : $":{IntervalEnd}";
        string? sheet = _sheet is null ? null : $"{_sheet}!";
        return $"{sheet}{IntervalStart}{intervalEnd}";
    }

    public static Range? Parse(string s)
    {
        string? sheet = null;
        int sheetSeparatorIndex = s.LastIndexOf('!');
        if (sheetSeparatorIndex > -1)
        {
            sheet = s[..sheetSeparatorIndex];
            s = s[(sheetSeparatorIndex + 1)..];
        }

        Bound? intervalStart;
        Bound? intervalEnd = null;
        int intervalSeparatorIndex = s.LastIndexOf(':');
        if (intervalSeparatorIndex > -1)
        {
            intervalStart = Bound.Parse(s[..intervalSeparatorIndex]);
            intervalEnd =  Bound.Parse(s[(intervalSeparatorIndex + 1)..]);
        }
        else
        {
            intervalStart = Bound.Parse(s);
        }

        return intervalStart is null ? null : new Range(intervalStart, intervalEnd, sheet);
    }

    private readonly string? _sheet;
}
