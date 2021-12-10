using System;
using System.Text;
using NDde.Advanced;
using NDde.Client;
using NDde.Server;
using NDde.Test.Helpers;
using NUnit.Framework;

namespace NDde.Test;

[TestFixture]
public sealed class Test_DdeServer
{
    private const string ServiceName = "myservice";
    private const string TopicName = "mytopic";
    private const string ItemName = "myitem";
    private const string CommandText = "mycommand";
    private const string TestData = "Hello World";
    private const int Timeout = 1000;

    [Test]
    public void Test_Ctor_Overload_1()
    {
        DdeServer server = new TestServer(ServiceName);
    }

    [Test]
    public void Test_Ctor_Overload_2()
    {
        using var context = new DdeContext();
        DdeServer server = new TestServer(ServiceName);
    }

    [Test]
    public void Test_Register()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
    }

    [Test]
    public void Test_Register_After_Dispose()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Dispose();
        Assert.Throws<ObjectDisposedException>(() => server.Register());
    }

    [Test]
    public void Test_Register_After_Register()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Register();
        Assert.Throws<InvalidOperationException>(() => server.Register());
    }

    [Test]
    public void Test_Unregister()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Register();
        server.Unregister();
        using var client = new DdeClient(ServiceName, TopicName);
        try
        {
            client.Connect();
            Assert.Fail();
        }
        catch (DdeException e)
        {
            Assert.AreEqual(0x400A, e.Code);
        }
    }

    [Test]
    public void Test_Unregister_After_Dispose()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Register();
        server.Dispose();
        Assert.Throws<ObjectDisposedException>(() => server.Unregister());
    }

    [Test]
    public void Test_Unregister_Before_Register()
    {
        using DdeServer server = new TestServer(ServiceName);
        Assert.Throws<InvalidOperationException>(() => server.Unregister());
    }

    [Test]
    public void Test_Execute()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        client.Execute(CommandText, Timeout);
        Assert.AreEqual(CommandText, server.Command);
    }

    [Test]
    public void Test_Execute_NotProcessed()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        try
        {
            client.Execute("#NotProcessed", Timeout);
        }
        catch (DdeException e)
        {
            Assert.AreEqual(0x4009, e.Code);
        }
    }

    [Test]
    public void Test_Execute_PauseConversation()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        client.Execute("#PauseConversation", (int) server.Interval * 2);
    }

    [Test]
    public void Test_Execute_TooBusy()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        try
        {
            client.Execute("#TooBusy", Timeout);
        }
        catch (DdeException e)
        {
            Assert.AreEqual(0x4001, e.Code);
        }
    }

    [Test]
    public void Test_Poke()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        client.Poke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout);
        Assert.AreEqual(TestData, Encoding.ASCII.GetString(server.GetData(TopicName, ItemName, 1)));
    }

    [Test]
    public void Test_Request()
    {
        using var server = new TestServer(ServiceName);
        server.Register();
        server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
        using var client = new DdeClient(ServiceName, TopicName);
        client.Connect();
        var data = client.Request(ItemName, 1, Timeout);
        Assert.AreEqual(TestData, Encoding.ASCII.GetString(data));
    }

    [Test]
    public void Test_IsRegistered_Variation_1()
    {
        using DdeServer server = new TestServer(ServiceName);
        Assert.IsFalse(server.IsRegistered);
    }

    [Test]
    public void Test_IsRegistered_Variation_2()
    {
        using DdeServer server = new TestServer(ServiceName);
        server.Register();
        Assert.IsTrue(server.IsRegistered);
    }

    [Test]
    public void Test_Service()
    {
        using DdeServer server = new TestServer(ServiceName);
        Assert.AreEqual(ServiceName, server.Service);
    }
} // class
// namespace