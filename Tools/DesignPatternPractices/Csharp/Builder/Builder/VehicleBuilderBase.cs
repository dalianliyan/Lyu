using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    public abstract class VehicleBuilderBase
    {
        public Vehicle Vehicle { get; protected set; }

        public virtual void Create()
        {
            Vehicle = new Vehicle();
        }

        public abstract void AddWheels();
        public abstract void AddLights();

    }
}
