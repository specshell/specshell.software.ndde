using System;
using System.Collections;
using System.Threading;
using NDde.Advanced;
using NDde.Server;
using NDde.Test.Helpers;
using NUnit.Framework;

namespace NDde.Test;

[TestFixture]
public sealed class Test_DdeContext
{
    private const string ServiceName = "test";
    private const int Timeout = 1000;

    [Test]
    public void Test_Ctor_Overload_1()
    {
        var context = new DdeContext();
    }

    [Test]
    public void Test_Ctor_Overload_2()
    {
        var context = new DdeContext(new DdeContext());
    }

    [Test]
    public void Test_Dispose()
    {
        using (var context = new DdeContext())
        {
        }
    }

    [Test]
    public void Test_Initialize()
    {
        using var context = new DdeContext();
        context.Initialize();
    }

    [Test]
    public void Test_Initialize_After_Dispose()
    {
        using var context = new DdeContext();
        context.Dispose();
        Assert.Throws<ObjectDisposedException>(() => context.Initialize());
    }

    [Test]
    public void Test_Initialize_After_Initialize()
    {
        using var context = new DdeContext();
        context.Initialize();
        Assert.Throws<InvalidOperationException>(() => context.Initialize());
    }

    [Test]
    public void Test_IsInitialized_Variation_1()
    {
        using var context = new DdeContext();
        Assert.IsFalse(context.IsInitialized);
    }

    [Test]
    public void Test_IsInitialized_Variation_2()
    {
        using var context = new DdeContext();
        context.Initialize();
        Assert.IsTrue(context.IsInitialized);
    }

    [Test]
    public void Test_AddTransactionFilter()
    {
        using var context = new DdeContext();
        IDdeTransactionFilter filter = new TransactionFilter();
        context.AddTransactionFilter(filter);
    }

    [Test]
    public void Test_AddTransactionFilter_After_Dispose()
    {
        using var context = new DdeContext();
        IDdeTransactionFilter filter = new TransactionFilter();
        context.Dispose();
        Assert.Throws<ObjectDisposedException>(() => context.AddTransactionFilter(filter));
    }

    [Test]
    public void Test_RemoveTransactionFilter()
    {
        using var context = new DdeContext();
        var filter = new TransactionFilter();
        context.AddTransactionFilter(filter);
        context.RemoveTransactionFilter(filter);
    }

    [Test]
    public void Test_RemoveTransactionFilter_After_Dispose()
    {
        using var context = new DdeContext();
        var filter = new TransactionFilter();
        context.AddTransactionFilter(filter);
        context.Dispose();
        Assert.Throws<ObjectDisposedException>(() => context.RemoveTransactionFilter(filter));
    }

    [Test]
    public void Test_TransactionFilter()
    {
        using var context = new DdeContext();
        var filter = new TransactionFilter();
        context.AddTransactionFilter(filter);
        context.Initialize();
        using (DdeServer server = new TestServer(ServiceName))
        {
            server.Register();
        }

        Assert.IsTrue(filter.Received.WaitOne(Timeout, false));
    }

    [Test]
    public void Test_Register()
    {
        using var context = new DdeContext();
        var listener = new EventListener();
        context.Register += listener.OnEvent;
        context.Initialize();
        using (DdeServer server = new TestServer(ServiceName))
        {
            server.Register();
        }

        Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
    }

    [Test]
    public void Test_Unregister()
    {
        using var context = new DdeContext();
        var listener = new EventListener();
        context.Unregister += listener.OnEvent;
        context.Initialize();
        using (DdeServer server = new TestServer(ServiceName))
        {
            server.Register();
            server.Unregister();
        }

        Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
    }

    private sealed class TransactionFilter : IDdeTransactionFilter
    {
        private readonly ManualResetEvent _Received = new(false);
        private readonly ArrayList _Transactions = new();

        public IList Transactions => ArrayList.ReadOnly(_Transactions);

        public WaitHandle Received => _Received;

        public bool PreFilterTransaction(DdeTransaction t)
        {
            _Transactions.Add(t);
            _Received.Set();
            return false;
        }
    }
} // class
// namespace