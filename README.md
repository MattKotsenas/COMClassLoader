# COMClassLoader
An extensible library for resolving and loading COM classes in .NET.

## Introduction

Using COM objects from .NET can be a pain. Like with the Global Assembly Cache (GAC), .NET prefers to interact with registered COM objects, which
complicates deployments, and is a big cause of "works on my machine" problems.

Making matters worse, .NET has accumulated multiple ways working with COM objects over the years, including Primary Interop Assembilies (PIAs), registration
free COM manifests, `tlbimp`, and more.

`COMClassLoader` provides a unified interface, extensibilitiy, and configuration-as-code for dealing with COM types.

## Examples

### Creating registered COM objects

Here's an example of creating a late-bound instance of a COM object, `IShellLink`, for a type that is known to be registered. All you have to do is declare your
managed types (or get `tlbimp` or Visual Studio to generate them for you), and call `RegisteredClassLoader.Load`.

```csharp
// Definitions of COM types (this may be done automatically by adding a COM library as a project reference)
[ComImport]
[Guid("0AFACED1-E828-11D1-9187-B532F1E9575D")]
internal class ShellLinkFolder
{
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("000214F9-0000-0000-C000-000000000046")]
internal interface IShellLink
{
    void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
    void GetIDList(out IntPtr ppidl);
    void SetIDList(IntPtr pidl);
    void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
    void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
    void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
    void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
    void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
    void GetHotkey(out short pwHotkey);
    void SetHotkey(short wHotkey);
    void GetShowCmd(out int piShowCmd);
    void SetShowCmd(int iShowCmd);
    void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
    void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
    void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
    void Resolve(IntPtr hwnd, int fFlags);
    void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
}

// Create a new instance of IShellLink
var link = new RegisteredClassLoader().Load<ShellLinkFolder, IShellLink>();

// Use methods as usual
var sb = new StringBuilder();
link.SetDescription("Hello, World!");
link.GetDescription(sb, sb.MaxCapacity);
Console.WriteLine(sb.ToString()); // Hello, World!
```

### Creating COM objects without registering

Using COM objects that haven't been registered is just as easy. The only difference is that you need to provide the path to the COM assembly like this

```csharp
var myComObject = new UnregisteredClassLoader(@"path\to\MyComClass.dll").Load<MyComEntryClass, IMyComInterface>();
```

and now you _must_ provide your own `ComImport` definitions.

## // TODO
* [ ] Support safe unloading of unmanaged assemblies (currently the COM server lives for the life of the host process)