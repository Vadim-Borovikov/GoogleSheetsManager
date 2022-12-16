using System;

namespace GoogleSheetsManager.Range;

internal readonly struct Range
{
    public readonly Bound IntervalStart;
    public readonly Bound? IntervalEnd;

    public Range(Bound intervalStart, Bound? intervalEnd = null, string? sheet = null)
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

    public Range GetFirstRow()
    {
        if (IntervalEnd is null)
        {
            return this;
        }

        Bound intervalEnd = new(IntervalEnd.Value.Column, IntervalStart.Row);
        return new Range(IntervalStart, intervalEnd, _sheet);
    }

    public static Range Parse(string s)
    {
        Range? parsed = TryParse(s);
        if (parsed is null)
        {
            throw new InvalidOperationException($"Can't parse range {s}.");
        }
        return parsed.Value;
    }

    private static Range? TryParse(string s)
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

        return intervalStart is null ? null : new Range(intervalStart.Value, intervalEnd, sheet);
    }

    private readonly string? _sheet;
}
