namespace GoogleSheetsManager;

public interface IConvertableTo<out T>
{
    T? Convert();
}
