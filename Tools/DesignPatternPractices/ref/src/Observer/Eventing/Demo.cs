using System;
namespace MarvellousWorks.PracticalPattern.Observer.Eventing
{
    /// <summary>
    /// ¾ßÌåµÄSubject
    /// </summary>
    public class UserEventArgs : EventArgs
    {
        public string Name { get; set; }
    }

    public class User
    {
        public event EventHandler<UserEventArgs> NameChanged;

        private string name;
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                // notify
                NameChanged(this, new UserEventArgs() {Name = value});
            }
        }
    }
}
