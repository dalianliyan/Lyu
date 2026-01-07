using System.Configuration;
namespace MarvellousWorks.PracticalPattern.StagePractice.Configuration
{
    /// <summary>
    /// 具有类型信息定义的配置元素
    /// </summary>
    public class WithKeyAndTypeNameConfigurationElement : ConfigurationElement
    {
        [ConfigurationProperty(Constant.KeyItem, IsRequired = true)]
        public string Key { get { return base[Constant.KeyItem] as string; } }

        [ConfigurationProperty(Constant.TypeNameItem, IsRequired = true)]
        public string TypeName { get { return base[Constant.TypeNameItem] as string; } }
    }

}
