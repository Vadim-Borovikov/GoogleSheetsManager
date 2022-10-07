using System.Linq;

namespace GoogleSheetsManager.Range;

internal sealed class Bound
{
    public ushort? Row;

    private Bound(string? column = null, ushort? row = null)
    {
        _column = column;
        Row = row;
    }

    public override string ToString() => $"{_column}{Row}";

    public static Bound? Parse(string s)
    {
        int? rowIndex = null;
        for (int i = 0; i < s.Length; ++i)
        {
            if (char.IsDigit(s[i]))
            {
                rowIndex = i;
                break;
            }
        }

        string? column;
        string? rowString = null;

        if (rowIndex.HasValue)
        {
            column = s[..rowIndex.Value];
            rowString = s[rowIndex.Value..];
        }
        else
        {
            column = s;
        }

        ushort? row = null;
        if (rowString is not null)
        {
            bool success = ushort.TryParse(rowString, out ushort result);
            if (!success)
            {
                return null;
            }

            row = result;
        }

        if (column.Any(c => !char.IsLetter(c)))
        {
            return null;
        }

        return new Bound(column, row);
    }

    private readonly string? _column;
}