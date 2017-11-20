using System;
using System.Runtime.InteropServices;
using PInvoke;

namespace COMClassLoader
{
    public class UnregisteredClassLoader : IClassLoader
    {
        private delegate int DllGetClassObject(ref Guid classId, ref Guid interfaceId, [Out, MarshalAs(UnmanagedType.Interface)] out object unknownObject);
        private readonly string _path;

        public UnregisteredClassLoader(string path)
        {
            _path = path;
        }

        public TResult Load<T, TResult>() where T : class where TResult : class
        {
            if (!Utilities.IsComImport(typeof(T)))
            {
                throw new ArgumentException($"Type {typeof(T)} is not a COM object");
            }

            var classFactoryId = typeof(IClassFactory).GUID;
            var classId = typeof(T).GUID;
            var interfaceId = typeof(TResult).GUID;

            var handle = Kernel32.LoadLibrary(_path);
            var dllGetClassObjectPtr = Kernel32.GetProcAddress(handle, "DllGetClassObject");
            var dllGetClassObjectDelegate = Marshal.GetDelegateForFunctionPointer<DllGetClassObject>(dllGetClassObjectPtr);

            dllGetClassObjectDelegate(ref classId, ref classFactoryId, out var classFactoryAsUnknown);
            var classFactory = classFactoryAsUnknown as IClassFactory;

            classFactory.CreateInstance(null, ref interfaceId, out var resultAsUnknown);
            return resultAsUnknown as TResult;
        }
    }
}