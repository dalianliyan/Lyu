namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 凭据抽象类型
    /// </summary>
    public abstract class CredentialBase{}

    public partial class UserNameCredential : CredentialBase { }
    public partial class WindowsCredential : CredentialBase { }
    public partial class UsbKeyCredential : CredentialBase { }
}
