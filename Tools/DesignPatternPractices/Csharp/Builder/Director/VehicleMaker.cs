using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    public class VehicleMaker
    {
        public VehicleBuilderBase Builder { get; set; }
        public Vehicle Vehicle { get { return Builder.Vehicle; } }

        public void Construct()
        {
            Builder.Create();
            Builder.AddWheels();
            Builder.AddLights();
        }
    }
}
