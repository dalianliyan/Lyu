using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Command
{
    public class PrintCommand : ICommand
    {
        private Receiver mReceiver;

        public PrintCommand(Receiver receiver)
        {
            this.mReceiver = receiver;
        }

        public void Execute()
        {
            mReceiver.Action();
        }
    }
}
