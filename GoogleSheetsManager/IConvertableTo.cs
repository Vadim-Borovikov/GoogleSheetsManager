namespace GoogleSheetsManager;

public interface IConvertibleTo<out T>
{
    T? Convert();
}
