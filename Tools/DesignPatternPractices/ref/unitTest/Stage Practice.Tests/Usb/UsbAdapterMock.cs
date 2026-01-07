using System.Diagnostics;

namespace External
{
    /// <summary>
    /// 模拟第三方提供的接口不兼容的驱动程序
    /// </summary>
    class UsbDriver
    {
        bool avaliable;
        public bool Avaliable { 
            get
            {
                Trace.WriteLine("External  UsbDriver.Avaliable.get()");
                return avaliable;
            }
            set 
            { 
                Trace.WriteLine("External  UsbDriver.Avaliable.set()");
                avaliable = value;
            }
        }
    }
}

namespace MarvellousWorks.PracticalPattern.StagePractice.Tests.Usb
{
    using System.ServiceModel.Security;
    using MarvellousWorks.PracticalPattern.StagePractice.Usb;
    using External;

    public class UsbAdapter : IUsbKeyAdapter
    {
        public const string Pin = "hello";
        public const string Name = "external";

        UsbDriver driver = new UsbDriver();

        public void Close()
        {
            Trace.WriteLine("UsbAdapter.Close()");
            driver.Avaliable = false;
        }

        public CredentialBase GetCredential(string pin)
        {
            Trace.WriteLine("UsbAdapter.GetCredential()");
            if(!string.Equals(pin, Pin)) throw new SecurityAccessDeniedException();
            return new UsbKeyCredential();
        }

        public bool IsOpen
        {
            get { return driver.Avaliable; } 
            set{ driver.Avaliable = value;}
        }

        public void Open()
        {
            Trace.WriteLine("UsbAdapter.Open()");
            if(!IsOpen)
                driver.Avaliable = true;
        }
    }
}
