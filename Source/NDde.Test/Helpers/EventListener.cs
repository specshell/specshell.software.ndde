using System.Collections.Generic;
using System.Threading;

namespace NDde.Test.Helpers;

internal sealed class EventListener
{
    private readonly ManualResetEvent _Received = new(false);

    public List<DdeEventArgs> Events { get; } = new();

    public WaitHandle Received => _Received;

    public void OnEvent(object sender, DdeEventArgs args)
    {
        Events.Add(args);
        _Received.Set();
    }
} // class
// namespace