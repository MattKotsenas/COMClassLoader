using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace COMClassLoader
{
    public class Utilities
    {
        internal static bool IsComImport(Type type)
        {
            var attribute = type.GetCustomAttribute(typeof(GuidAttribute));
            return attribute != null;
        }
    }
}