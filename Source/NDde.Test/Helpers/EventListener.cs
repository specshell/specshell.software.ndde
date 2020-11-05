using System.Collections.Generic;
using System.Threading;

namespace NDde.Test.Helpers
{
    internal sealed class EventListener
    {
        private ManualResetEvent _Received = new ManualResetEvent(false);
        private List<DdeEventArgs> _Events = new List<DdeEventArgs>();

        public List<DdeEventArgs> Events => _Events;

        public WaitHandle Received => _Received;

        public void OnEvent(object sender, DdeEventArgs args)
        {
            _Events.Add(args);
            _Received.Set();
        }
    } // class
} // namespace