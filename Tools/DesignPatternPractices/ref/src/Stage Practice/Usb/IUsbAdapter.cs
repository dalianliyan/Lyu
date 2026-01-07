namespace MarvellousWorks.PracticalPattern.StagePractice.Usb
{
    /// <summary>
    /// USB key 驱动适配接口
    /// </summary>
    /// <remarks>
    ///     为了隔绝第三方厂商的不确定性，认证服务先手定义的访问接口
    /// </remarks>
    public interface IUsbKeyAdapter
    {
        /// <summary>
        /// 设备是否打开
        /// </summary>
        bool IsOpen { get; set; }

        /// <summary>
        /// 打开端口
        /// </summary>
        void Open();

        /// <summary>
        /// 获取Usb设备内凭证
        /// </summary>
        /// <param name="pin">与驱动通信的pin</param>
        /// <returns></returns>
        CredentialBase GetCredential(string pin);

        /// <summary>
        /// 关闭端口
        /// </summary>
        void Close();
    }
}
