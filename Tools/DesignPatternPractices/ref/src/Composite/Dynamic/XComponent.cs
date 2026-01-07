using System;
using System.Dynamic;
using System.Xml.Linq;
namespace MarvellousWorks.PracticalPattern.Composite.Dynamic
{
    public class XComponent : DynamicObject
    {
        public XComponent(XElement node)
        {
            this.Element = node;
        }

        public XComponent()
        {
        }

        public XComponent(String name)
        {
            Element = new XElement(name);
        }

        public XElement Element { get; private set; }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            var temp = Element.Element(binder.Name);
            if(temp == null)
            {
                if (value.GetType() == typeof(XComponent))
                    Element.Add(new XElement(binder.Name));
                else
                    Element.Add(new XElement(binder.Name, value));
            }
            else
                temp.SetValue(value);
            return true;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            var temp = Element.Element(binder.Name);
            if (temp == null)
            {
                result = null;
                return false;
            }
            else
            {
                result = new XComponent(temp);
                return true;
            }
        }
    }
}

