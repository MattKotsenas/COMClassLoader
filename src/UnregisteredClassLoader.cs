using System;
using System.Runtime.InteropServices;
using PInvoke;

namespace COMClassLoader
{
    // TODO: Since this loader implements IDispoable, it should be added to the IClassLoader interface instead
    public class UnregisteredClassLoader : IClassLoader, IDisposable
    {
        private delegate int DllGetClassObject(ref Guid classId, ref Guid interfaceId, [Out, MarshalAs(UnmanagedType.Interface)] out object unknownObject);
        private readonly string _path;
        private Kernel32.SafeLibraryHandle _library;

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_library != null && !_library.IsClosed)
                {
                    _library.Dispose();
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

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

            var path = EnsureLibraryExists(_path);

            var classFactoryId = typeof(IClassFactory).GUID;
            var classId = typeof(T).GUID;
            var interfaceId = typeof(TResult).GUID;

            _library = EnsureLibraryLoaded(path);

            var dllGetClassObjectPtr = Kernel32.GetProcAddress(_library, "DllGetClassObject");
            var dllGetClassObjectDelegate = Marshal.GetDelegateForFunctionPointer<DllGetClassObject>(dllGetClassObjectPtr);

            dllGetClassObjectDelegate(ref classId, ref classFactoryId, out var classFactoryAsUnknown);
            var classFactory = classFactoryAsUnknown as IClassFactory;

            classFactory.CreateInstance(null, ref interfaceId, out var resultAsUnknown);
            return resultAsUnknown as TResult;
        }

        private static string EnsureLibraryExists(string path)
        {
            var file = new System.IO.FileInfo(path);
            if (!file.Exists)
            {
                throw new ArgumentException($"The path {path} does not exist");
            }
            return file.FullName;
        }

        private static Kernel32.SafeLibraryHandle EnsureLibraryLoaded(string path)
        {
            var library = Kernel32.LoadLibrary(path);
            if (library.IsInvalid)
            {
                throw new TypeLoadException($"LoadLibrary of library {path} failed");
            }
            return library;
        }
    }
}