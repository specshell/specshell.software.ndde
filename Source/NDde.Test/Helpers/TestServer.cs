using System.Collections;
using System.Linq;
using System.Timers;
using NDde.Advanced;
using NDde.Server;

namespace NDde.Test.Helpers;

internal class TestServer : TracingServer
{
    private readonly IDictionary _Conversation = new Hashtable();
    private readonly IDictionary _Data = new Hashtable();
    private readonly Timer _Timer = new();

    public TestServer(string service)
        : base(service)
    {
        _Timer.Elapsed += OnTimerElapsed;
        _Timer.Interval = 1000;
        _Timer.SynchronizingObject = Context;
    }

    public TestServer(string service, DdeContext context)
        : base(service, context)
    {
        _Timer.Elapsed += OnTimerElapsed;
        _Timer.Interval = 1000;
        _Timer.SynchronizingObject = Context;
    }

    public double Interval => _Timer.Interval;

    public string Command { get; private set; } = "";

    public byte[] GetData(string topic, string item, int format)
    {
        var key = topic + ":" + item + ":" + format;
        return (byte[]) _Data[key];
    }

    public void SetData(string topic, string item, int format, byte[] data)
    {
        var key = topic + ":" + item + ":" + format;
        _Data[key] = data;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _Timer.Dispose();
        base.Dispose(true);
    }

    private void OnTimerElapsed(object sender, ElapsedEventArgs args)
    {
        foreach (DdeConversation c in _Conversation.Values)
            if (c.IsPaused)
                Resume(c);

        if (_Conversation.Values.Cast<DdeConversation>().Any(c => c.IsPaused))
        {
            return;
        }

        _Timer.Stop();
    }

    protected override bool OnBeforeConnect(string topic)
    {
        base.OnBeforeConnect(topic);
        return true;
    }

    protected override void OnAfterConnect(DdeConversation conversation)
    {
        base.OnAfterConnect(conversation);
        _Conversation.Add(conversation.Handle, conversation);
    }

    protected override void OnDisconnect(DdeConversation conversation)
    {
        base.OnDisconnect(conversation);
        _Conversation.Remove(conversation.Handle);
    }

    protected override bool OnStartAdvise(DdeConversation conversation, string item, int format)
    {
        base.OnStartAdvise(conversation, item, format);
        return true;
    }

    protected override ExecuteResult OnExecute(DdeConversation conversation, string command)
    {
        base.OnExecute(conversation, command);
        Command = command;
        switch (command)
        {
            case "#NotProcessed":
            {
                return ExecuteResult.NotProcessed;
            }
            case "#PauseConversation":
            {
                if ((string) conversation.Tag == command)
                {
                    conversation.Tag = null;
                    return ExecuteResult.Processed;
                }

                conversation.Tag = command;
                if (!_Timer.Enabled) _Timer.Start();
                return ExecuteResult.PauseConversation;
            }
            case "#Processed":
            {
                return ExecuteResult.Processed;
            }
            case "#TooBusy":
            {
                return ExecuteResult.TooBusy;
            }
        }

        return ExecuteResult.Processed;
    }

    protected override PokeResult OnPoke(DdeConversation conversation, string item, byte[] data, int format)
    {
        base.OnPoke(conversation, item, data, format);
        var key = conversation.Topic + ":" + item + ":" + format;
        _Data[key] = data;
        switch (item)
        {
            case "#NotProcessed":
            {
                return PokeResult.NotProcessed;
            }
            case "#PauseConversation":
            {
                if ((string) conversation.Tag == item)
                {
                    conversation.Tag = null;
                    return PokeResult.Processed;
                }

                conversation.Tag = item;
                if (!_Timer.Enabled) _Timer.Start();
                return PokeResult.PauseConversation;
            }
            case "#Processed":
            {
                return PokeResult.Processed;
            }
            case "#TooBusy":
            {
                return PokeResult.TooBusy;
            }
        }

        return PokeResult.Processed;
    }

    protected override RequestResult OnRequest(DdeConversation conversation, string item, int format)
    {
        base.OnRequest(conversation, item, format);
        var key = conversation.Topic + ":" + item + ":" + format;
        return _Data.Contains(key) ? new RequestResult((byte[]) _Data[key]) : RequestResult.NotProcessed;
    }

    protected override byte[] OnAdvise(string topic, string item, int format)
    {
        base.OnAdvise(topic, item, format);
        var key = topic + ":" + item + ":" + format;
        if (_Data.Contains(key)) return (byte[]) _Data[key];
        return null;
    }
} // class
// namespace