using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using COMClassLoader;
using UnregisteredCOMLoader.Tests.Mocks;

namespace UnregisteredCOMLoader.Tests
{
    [TestClass]
    public class given_an_unregistered_COM_object_with_an_absolute_path
    {
        [TestMethod]
        [DeploymentItem(@"Mocks\MyComClass.dll")]
        [DeploymentItem("Microsoft.VisualStudio.TestPlatform.TestFramework.Extensions.dll")] // https://github.com/Microsoft/testfx/issues/91
        public void it_can_be_instantiated_through_the_dllclassobject()
        {
            var path = Path.Combine(System.Environment.CurrentDirectory, "MyComClass.dll");

            using (var loader = new UnregisteredClassLoader(path))
            {
                var myComObject = loader.Load<MyComEntryClass, IMyComInterface>();
                Assert.IsNotNull(myComObject);
            }
        }
    }

    [TestClass]
    public class given_an_unregistered_COM_object_with_a_relative_path
    {
        [TestMethod]
        [DeploymentItem(@"Mocks\MyComClass.dll")]
        [DeploymentItem("Microsoft.VisualStudio.TestPlatform.TestFramework.Extensions.dll")] // https://github.com/Microsoft/testfx/issues/91
        public void it_can_be_instantiated_through_the_dllclassobject()
        {
            var path = "MyComClass.dll";

            using (var loader = new UnregisteredClassLoader(path))
            {
                var myComObject = loader.Load<MyComEntryClass, IMyComInterface>();
                Assert.IsNotNull(myComObject);
            }
        }
    }
}
