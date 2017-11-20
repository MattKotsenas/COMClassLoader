using COMClassLoader;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UnregisteredCOMLoader.Tests.Mocks;

namespace UnregisteredCOMLoader.Tests
{
    [TestClass]
    public class given_an_unregistered_COM_object
    {
        [TestMethod]
        [DeploymentItem(@"Mocks\MyComClass.dll")]
        [DeploymentItem("Microsoft.VisualStudio.TestPlatform.TestFramework.Extensions.dll")] // https://github.com/Microsoft/testfx/issues/91
        public void it_can_be_instantiated_through_the_dllclassobject()
        {
            var myComObject = new UnregisteredClassLoader("MyComClass.dll").Load<MyComEntryClass, IMyComInterface>();
            Assert.IsNotNull(myComObject);
        }
    }
}
