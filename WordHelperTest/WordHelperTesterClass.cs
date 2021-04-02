using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word_Helper;
namespace WordHelperTest
{
    [TestClass]
    public class WordHelperTesterClass
    {
        [TestMethod]
        public void insertImageTest1()
        {

            WordHelperClass wordHelperObj = new WordHelperClass();

            string status = wordHelperObj.insertPicture(@"E:\Study Backups\Word Helper\Test\Test.docx", @"E:\Study Backups\Word Helper\Test\Devneet.jpg");

            Assert.IsTrue(status.Contains("SUCCESS"));

        }
    }
}
