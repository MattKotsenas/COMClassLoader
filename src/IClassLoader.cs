namespace COMClassLoader
{
    public interface IClassLoader
    {
        TResult Load<T, TResult>() where T : class where TResult : class;
    }
}