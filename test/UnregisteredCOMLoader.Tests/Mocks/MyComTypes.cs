using System.Runtime.InteropServices;

namespace UnregisteredCOMLoader.Tests.Mocks
{
    [ComImport]
    [Guid("E6756135-1E65-4D17-8576-610761398C3C")]
    public class MyComEntryClass
    {
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("79F1BB5F-B66E-48E5-B6A9-1545C323CA3D")]
    public interface IMyComInterface
    {
    }
}