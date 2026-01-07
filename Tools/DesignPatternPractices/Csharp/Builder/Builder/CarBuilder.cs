using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    public class CarBuilder : VehicleBuilderBase
    {
        #region VehicleBuilderBase Members
        public override void AddWheels()
        {
            Vehicle.Wheels = new string[] { "front", "front", "rear", "rear" };
        }

        public override void AddLights()
        {
            Vehicle.Lights = new string[] { "front", "front", "rear", "rear" };
        }
        #endregion
    }
}
