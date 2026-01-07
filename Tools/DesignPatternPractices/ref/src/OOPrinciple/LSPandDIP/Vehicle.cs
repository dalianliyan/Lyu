namespace MarvellousWorks.PracticalPattern.OO.LSPandDIP
{
    interface IVehicle
    {
        void Run();
    }

    class Bicycle : IVehicle
    {
        public void Run()
        {
            // run like a bicycle
        }
    }

    class Train : IVehicle
    {
        public void Run()
        {
            // run like a train
        }
    }

    class Client
    {
        public void ShowAVehicleRun(IVehicle vehicle)
        {
            vehicle.Run();
        }
    }
}
