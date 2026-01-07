using System.ServiceModel;

namespace MarvellousWorks.PracticalPattern.Proxy.Classic
{
/// <summary>
/// 定义客户程序需要的抽象类型
/// </summary>
[ServiceContract]
public interface ISubject
{
    [OperationContract]
    string Request();
}

/// <summary>
/// 具体实现客户程序需要的类型
/// </summary>
public class RealSubject : ISubject
{
    RealSubject(){}
    public static readonly RealSubject Instance = new RealSubject();
    public string Request() { return "from real subject"; }
}

/// <summary>
/// 代理类型，他知道如何满足客户程序的要求，同时知道具体类型如何访问
/// </summary>
public class Proxy : ISubject
{
    public string Request() { return RealSubject.Instance.Request(); }
}
}
