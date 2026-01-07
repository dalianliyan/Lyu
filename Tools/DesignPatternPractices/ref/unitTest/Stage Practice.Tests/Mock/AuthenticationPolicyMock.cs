using System;
using System.Collections.Generic;
using System.Diagnostics;
using MarvellousWorks.PracticalPattern.StagePractice.Tests.Properties;
namespace MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock
{
    class AuthenticationPolicyMockBase : IAuthenticationPolicy
    {
        protected IList<Type> ApplyToCredentialTypes { get; set; }

        public bool IsMatch(CredentialBase credential)
        {
            if (credential == null) throw new ArgumentNullException("credential");
            bool result;
            if (ApplyToCredentialTypes.Count == 0)
                result = false;
            else
                result = ApplyToCredentialTypes.Contains(credential.GetType());
            if(result)
                Trace.WriteLine(string.Format(Resources.OutputFormat, "Policy", GetType().Name, "match"));
            return result;
        }
    }

    /// <summary>
    /// 仅适于Usbkey类型凭证
    /// </summary>
    class UsbKeyPolicy : AuthenticationPolicyMockBase
    {
        public UsbKeyPolicy()
        {
            ApplyToCredentialTypes = new List<Type>(){typeof(UsbKeyCredential)};
        }
    }

    /// <summary>
    /// 仅适于用户名口令类型凭证
    /// </summary>
    class UserNamePolicy : AuthenticationPolicyMockBase
    {
        public UserNamePolicy()
        {
            ApplyToCredentialTypes = new List<Type>() { typeof(UserNameCredential) };
        }
    }

    /// <summary>
    /// 适于用户名口令和UsbKey两种凭证类型
    /// </summary>
    class UserNameAndUsbKeyPolicy : AuthenticationPolicyMockBase
    {
        public UserNameAndUsbKeyPolicy()
        {
            ApplyToCredentialTypes = new List<Type>() { typeof(UserNameCredential), typeof(UsbKeyCredential) };
        }
    }

    /// <summary>
    /// 适用于所有凭证类型
    /// </summary>
    class AllPolicy : IAuthenticationPolicy
    {
        public bool IsMatch(CredentialBase credential)
        {
            Trace.WriteLine(string.Format(Resources.OutputFormat, "Policy", GetType().Name, "match"));
            return true;
        }
    }

}
