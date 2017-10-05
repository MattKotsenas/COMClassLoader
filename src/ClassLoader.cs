namespace COMClassLoader
{
    public class ClassLoader
    {
        public static TResult Load<T, TResult>() where T : class, new() where TResult : class
        {
            return new T() as TResult;
        }
    }
}
