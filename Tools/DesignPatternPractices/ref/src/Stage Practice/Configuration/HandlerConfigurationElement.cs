using System.Configuration;
namespace MarvellousWorks.PracticalPattern.StagePractice.Configuration
{
    public interface IHandlerConfigurationEntry
    {
        string Key { get; }
        string Sequence { get; }
        string Policy { get; }
        string TypeName { get; }
    }

    /// <summary>
    /// 认证过程非功能性控制对象的配置元素
    /// </summary>
    public class HandlerConfigurationElement : WithKeyAndTypeNameConfigurationElement, IHandlerConfigurationEntry
    {
        [ConfigurationProperty(Constant.SequenceItem, IsRequired = true)]
        public string Sequence { get { return base[Constant.SequenceItem] as string; } }

        [ConfigurationProperty(Constant.PolicyItem, IsRequired = true)]
        public string Policy { get { return base[Constant.PolicyItem] as string; } }
    }

    /// <summary>
    /// 认证过程非功能性控制对象的配置元素集合
    /// </summary>
    [ConfigurationCollection(typeof(HandlerConfigurationElement),
        CollectionType = ConfigurationElementCollectionType.AddRemoveClearMap)]
    public class HandlerConfigurationElementCollection : ConfigurationElementCollection
    {
        public HandlerConfigurationElement this[int index] { get { return (HandlerConfigurationElement)base.BaseGet(index); } }
        public new HandlerConfigurationElement this[string name] { get { return (HandlerConfigurationElement)base.BaseGet(name); } }
        protected override ConfigurationElement CreateNewElement(){return new HandlerConfigurationElement();}
        protected override object GetElementKey(ConfigurationElement element) {return (element as HandlerConfigurationElement).Key; }
    }
}
