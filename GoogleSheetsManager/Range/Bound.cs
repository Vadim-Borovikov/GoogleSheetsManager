using System;
using System.Linq;

namespace GoogleSheetsManager.Range;

internal readonly struct Bound
{
    internal readonly string? Column;
    internal readonly ushort? Row;

    internal Bound(string? column = null, ushort? row = null)
    {
        Column = column;
        Row = row;
    }

    public override string ToString() => $"{Column}{Row}";

    // ReSharper disable once UnusedMember.Global
    public static Bound Parse(string value)
    {
        if (!TryParse(value, out Bound result))
        {
            throw new FormatException($"String \"{value}\" was not recognized as a valid {nameof(Bound)}.");
        }
        return result;
    }

    public static bool TryParse(string value, out Bound result)
    {
        result = default;

        int? rowIndex = null;
        for (int i = 0; i < value.Length; ++i)
        {
            if (char.IsDigit(value[i]))
            {
                rowIndex = i;
                break;
            }
        }

        string? column;
        string? rowString = null;

        if (rowIndex.HasValue)
        {
            column = value[..rowIndex.Value];
            rowString = value[rowIndex.Value..];
        }
        else
        {
            column = value;
        }

        ushort? row = null;
        if (rowString is not null)
        {
            bool success = ushort.TryParse(rowString, out ushort u);
            if (!success)
            {
                return false;
            }

            row = u;
        }

        if (column.Any(c => !char.IsLetter(c)))
        {
            return false;
        }

        result = new Bound(column, row);
        return true;
    }
}