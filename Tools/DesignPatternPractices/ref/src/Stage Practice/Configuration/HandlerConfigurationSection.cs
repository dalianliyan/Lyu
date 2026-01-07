using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace MarvellousWorks.PracticalPattern.StagePractice.Configuration
{
    /// <summary>
    /// 认证过程非功能性控制对象的配置节
    /// </summary>
    public class HandlerConfigurationSection : ConfigurationSection
    {
        public IEnumerable<IHandlerConfigurationEntry> GetHandlers()
        {
            if((Handlers == null) || (Handlers.Count == 0)) return null;
            var result = new List<IHandlerConfigurationEntry>();
            for(var i=0; i<Handlers.Count; i++)
                result.Add(Handlers[i]);
            return result;
        }

        [ConfigurationProperty(Constant.HandlerCollectionName, IsRequired = false)]
        public HandlerConfigurationElementCollection Handlers
        {
            get
            {
                return base[Constant.HandlerCollectionName] as HandlerConfigurationElementCollection;
            }
        }


        /// <summary>
        /// IAuthenticationHandler CoR的构造对象
        /// </summary>
        /// <remarks>默认情况下为内置的AuthenticationHandlerCoRBuilder</remarks>
        [ConfigurationProperty(Constant.BuilderItem, IsRequired = false)]
        public string BuilderTypeName
        {
            get
            {
                var typeName = base[Constant.BuilderItem] as string;
                if (string.IsNullOrEmpty(typeName))
                    typeName = typeof (AuthenticationHandlerCoRBuilder).AssemblyQualifiedName;
                return typeName;
            }
        }
    }
}
