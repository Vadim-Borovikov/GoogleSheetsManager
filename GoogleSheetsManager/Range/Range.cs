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

    public static Range Parse(string value)
    {
        if (!TryParse(value, out Range result))
        {
            throw new FormatException($"String \"{value}\" was not recognized as a valid {nameof(Range)}.");
        }
        return result;
    }

    private static bool TryParse(string value, out Range result)
    {
        result = default;

        string? sheet = null;
        int sheetSeparatorIndex = value.LastIndexOf('!');
        if (sheetSeparatorIndex > -1)
        {
            sheet = value[..sheetSeparatorIndex];
            value = value[(sheetSeparatorIndex + 1)..];
        }

        Bound intervalStart;
        Bound? intervalEnd = null;
        int intervalSeparatorIndex = value.LastIndexOf(':');
        if (intervalSeparatorIndex > -1)
        {
            if (!Bound.TryParse(value[..intervalSeparatorIndex], out intervalStart))
            {
                return false;
            }
            if (Bound.TryParse(value[(intervalSeparatorIndex + 1)..], out Bound intervalEndParsed))
            {
                intervalEnd = intervalEndParsed;
            }
        }
        else
        {
            if (!Bound.TryParse(value, out intervalStart))
            {
                return false;
            }
        }

        result = new Range(intervalStart, intervalEnd, sheet);
        return true;
    }

    private readonly string? _sheet;
}
