using System;

namespace GoogleSheetsManager.Range;

internal readonly struct Range
{
    private Range(Bound? intervalStart, Bound? intervalEnd, string? sheet)
    {
        _intervalStart = intervalStart;
        _intervalEnd = intervalEnd;
        _sheet = sheet;
    }

    public override string ToString()
    {
        if (_intervalStart is null)
        {
            return _sheet ?? string.Empty;
        }

        string? intervalEnd = _intervalEnd is null ? null : $":{_intervalEnd}";
        string? sheet = _sheet is null ? null : $"{_sheet}!";
        return $"{sheet}{_intervalStart}{intervalEnd}";
    }

    public Range GetFirstRow() => new(_intervalStart ?? One, _intervalEnd ?? One, _sheet);

    public static Range ParseAndAddName(string name, string? value = null)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new Range(null, null, name);
        }

        if (!TryParse(value, out Range result))
        {
            throw new FormatException($"String \"{value}\" was not recognized as a valid {nameof(Range)}.");
        }
        return string.IsNullOrEmpty(name) ? result : new Range(result._intervalStart, result._intervalEnd, name);
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
    private readonly Bound? _intervalStart;
    private readonly Bound? _intervalEnd;

    private static readonly Bound One = new(null, 1);
}
