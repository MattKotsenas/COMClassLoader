using System;

namespace COMClassLoader
{
    public class RegisteredClassLoader : IClassLoader
    {
        public TResult Load<T, TResult>() where T : class where TResult : class
        {
            if (!Utilities.IsComImport(typeof(T)))
            {
                throw new ArgumentException($"Type {typeof(T)} is not a COM object");
            }

            var guid = typeof(T).GUID;
            var type = Type.GetTypeFromCLSID(guid);

            return Activator.CreateInstance(type) as TResult;
        }
    }
}
