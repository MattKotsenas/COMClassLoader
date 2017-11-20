using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RegisteredCOMLoader.Tests.Mocks;
using COMClassLoader;

namespace RegisteredCOMLoader.Tests
{
    [TestClass]
    public class given_a_registered_COM_object
    {
        private const string Expected = "Hello World";

        [TestMethod]
        public void it_can_be_instantiated_directly()
        {
            var link = new RegisteredClassLoader().Load<ShellLinkFolder, IShellLink>();

            var sb = new StringBuilder();
            link.SetDescription(Expected);
            link.GetDescription(sb, sb.MaxCapacity);

            Assert.AreEqual(Expected, sb.ToString());
        }
    }
}
