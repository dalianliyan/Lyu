using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Command
{
    class Client
    {
        static void Main(string[] args)
        {
            Receiver receiver = new Receiver();
            ICommand cmd = new PrintCommand(receiver);
            Invoker invoker = new Invoker();
            invoker.StoreCommand(cmd);
            invoker.Invoke();
        }    
    }


    public class Receiver
    {
        public void Action()
        {

        }
    }

    public class Invoker
    {
        private ICommand _cmd;
        public void StoreCommand(ICommand cmd)
        {
            _cmd = cmd;
        }

        public void Invoke()
        {
            if (_cmd != null)
            {
                _cmd.Execute();
            }
        }
    }
}
