using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock;
namespace MarvellousWorks.PracticalPattern.StagePractice.Tests
{
    [TestClass]
    public class AuthenticatorFixture
    {
        /// <summary>
        /// 验证USB Key类凭证
        /// </summary>
        [TestMethod]
        public void TestUsbKeyAuthentication()
        {
            Trace.WriteLine("TestUsbKeyAuthentication");
            var credential = new UsbKeyCredential();
            new AuthenticatorFactory().Create(credential).Authenticate(credential);
        }

        /// <summary>
        /// 验证用户名/口令类凭证
        /// </summary>
        [TestMethod]
        public void TestUserNameAuthentication()
        {
            Trace.WriteLine("TestUserNameAuthentication");
            var credential = new UserNameCredential();
            new AuthenticatorFactory().Create(credential).Authenticate(credential);
        }

        /// <summary>
        /// 新增的自定义凭证类型
        /// </summary>
        /// <exception cref="ArgumentException">因为没有定义相应的认证服务Provider，因此会抛出异常</exception>
        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void TestCustomsAuthentication()
        {
            Trace.WriteLine("TestCustomsAuthentication");
            var credential = new CustomsCredential();
            new AuthenticatorFactory().Create(credential).Authenticate(credential);
        }
    }
}
