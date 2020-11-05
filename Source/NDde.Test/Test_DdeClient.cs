using NDde.Test.Helpers;

namespace NDde.Test
{
    using System;
    using System.Collections;
    using System.Text;
    using NDde;
    using Advanced;
    using Client;
    using Server;
    using NUnit.Framework;

    [TestFixture]
    public sealed class Test_DdeClient
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
            var client = new DdeClient(ServiceName, TopicName);
        }

        [Test]
        public void Test_Ctor_Overload_2()
        {
            using (var context = new DdeContext())
            {
                var client = new DdeClient(ServiceName, TopicName, context);
            }
        }

        [Test]
        public void Test_Dispose()
        {
            using (var client = new DdeClient(ServiceName, TopicName))
            {
            }
        }

        [Test]
        public void Test_Service()
        {
            using (var client = new DdeClient(ServiceName, TopicName))
            {
                Assert.AreEqual(ServiceName, client.Service);
            }
        }

        [Test]
        public void Test_Topic()
        {
            using (var client = new DdeClient(ServiceName, TopicName))
            {
                Assert.AreEqual(TopicName, client.Topic);
            }
        }

        [Test]
        public void Test_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                }
            }
        }

        [Test]
        public void Test_Connect_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Connect());
                }
            }
        }

        [Test]
        public void Test_Connect_After_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.Throws<InvalidOperationException>(() => client.Connect());
                }
            }
        }

        [Test]
        public void Test_Disconnect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Disconnect();
                }
            }
        }

        [Test]
        public void Test_Disconnect_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Disconnect());
                }
            }
        }

        [Test]
        public void Test_Disconnect_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.Disconnect());
                }
            }
        }


        [Test]
        public void Test_Handle_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.AreEqual(IntPtr.Zero, client.Handle);
                }
            }
        }

        [Test]
        public void Test_Handle_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.AreNotEqual(IntPtr.Zero, client.Handle);
                }
            }
        }

        [Test]
        public void Test_Handle_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Disconnect();
                    Assert.AreEqual(IntPtr.Zero, client.Handle);
                }
            }
        }

        [Test]
        public void Test_IsConnected_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.IsFalse(client.IsConnected);
                }
            }
        }

        [Test]
        public void Test_IsConnected_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.IsTrue(client.IsConnected);
                }
            }
        }

        [Test]
        public void Test_IsConnected_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Disconnect();
                    Assert.IsFalse(client.IsConnected);
                }
            }
        }

        [Test]
        public void Test_IsConnected_Variation_4()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Disconnected += listener.OnEvent;
                    client.Connect();
                    server.Disconnect();
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    Assert.IsFalse(client.IsConnected);
                }
            }
        }

        [Test]
        public void Test_Pause()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    var ar = client.BeginExecute(CommandText, null, null);
                    Assert.IsFalse(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_Pause_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Pause());
                }
            }
        }

        [Test]
        public void Test_Pause_After_Pause()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    Assert.Throws<InvalidOperationException>(() => client.Pause());
                }
            }
        }

        [Test]
        public void Test_Resume()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    var ar = client.BeginExecute(CommandText, null, null);
                    Assert.IsFalse(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Resume();
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_Resume_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Resume());
                }
            }
        }

        [Test]
        public void Test_Resume_Before_Pause()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.Throws<InvalidOperationException>(() => client.Resume());
                }
            }
        }

        [Test]
        public void Test_Abandon()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    var ar = client.BeginExecute(CommandText, null, null);
                    Assert.IsFalse(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Abandon(ar);
                    client.Resume();
                    Assert.IsFalse(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_Abandon_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    var ar = client.BeginExecute(CommandText, null, null);
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Abandon(ar));
                }
            }
        }

        [Test]
        public void Test_IsPaused_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.IsFalse(client.IsPaused);
                }
            }
        }

        [Test]
        public void Test_IsPaused_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    Assert.IsTrue(client.IsPaused);
                }
            }
        }

        [Test]
        public void Test_IsPaused_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Pause();
                    client.Resume();
                    Assert.IsFalse(client.IsPaused);
                }
            }
        }

        [Test]
        public void Test_Poke()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Poke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(server.GetData(TopicName, ItemName, 1)));
                }
            }
        }

        [Test]
        public void Test_TryPoke_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var result = client.TryPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout);
                    Assert.AreNotEqual(0, result);
                }
            }
        }

        [Test]
        public void Test_TryPoke_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var result = client.TryPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout);
                    Assert.AreEqual(0, result);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(server.GetData(TopicName, ItemName, 1)));
                }
            }
        }

        [Test]
        public void Test_Poke_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Poke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout));
                }
            }
        }

        [Test]
        public void Test_Poke_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.Poke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, Timeout));
                }
            }
        }

        [Test]
        public void Test_BeginPoke()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_BeginPoke_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() =>
                        client.BeginPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, null, null));
                }
            }
        }

        [Test]
        public void Test_BeginPoke_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() =>
                        client.BeginPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, null, null));
                }
            }
        }

        [Test]
        public void Test_EndPoke()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.EndPoke(ar);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(server.GetData(TopicName, ItemName, 1)));
                }
            }
        }

        [Test]
        public void Test_EndPoke_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginPoke(ItemName, Encoding.ASCII.GetBytes(TestData), 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.EndPoke(ar));
                }
            }
        }

        [Test]
        public void Test_Request_Overload_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var data = client.Request(ItemName, 1, Timeout);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(data));
                }
            }
        }

        [Test]
        public void Test_Request_Overload_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var data = client.Request(ItemName, Timeout);
                    Assert.AreEqual(TestData, data);
                }
            }
        }

        [Test]
        public void Test_TryRequest_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    byte[] data;
                    var result = client.TryRequest(ItemName, 1, Timeout, out data);
                    Assert.AreNotEqual(0, result);
                }
            }
        }

        [Test]
        public void Test_TryRequest_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    byte[] data;
                    var result = client.TryRequest(ItemName, 1, Timeout, out data);
                    Assert.AreEqual(0, result);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(data));
                }
            }
        }

        [Test]
        public void Test_Request_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Request(ItemName, 1, Timeout));
                }
            }
        }

        [Test]
        public void Test_Request_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.Request(ItemName, 1, Timeout));
                }
            }
        }

        [Test]
        public void Test_BeginRequest()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginRequest(ItemName, 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_BeginRequest_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.BeginRequest(ItemName, 1, null, null));
                }
            }
        }

        [Test]
        public void Test_BeginRequest_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.BeginRequest(ItemName, 1, null, null));
                }
            }
        }

        [Test]
        public void Test_EndRequest()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginRequest(ItemName, 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    var data = client.EndRequest(ar);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(data));
                }
            }
        }

        [Test]
        public void Test_EndRequest_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginRequest(ItemName, 1, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.EndRequest(ar));
                }
            }
        }

        [Test]
        public void Test_Execute()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Execute(TestData, Timeout);
                    Assert.AreEqual(TestData, server.Command);
                }
            }
        }

        [Test]
        public void Test_TryExecute_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var result = client.TryExecute(TestData, Timeout);
                    Assert.AreNotEqual(0, result);
                }
            }
        }

        [Test]
        public void Test_TryExecute_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var result = client.TryExecute(TestData, Timeout);
                    Assert.AreEqual(0, result);
                    Assert.AreEqual(TestData, server.Command);
                }
            }
        }

        [Test]
        public void Test_Execute_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.Execute(TestData, Timeout));
                }
            }
        }

        [Test]
        public void Test_Execute_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.Execute(TestData, Timeout));
                }
            }
        }

        [Test]
        public void Test_BeginExecute()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginExecute(TestData, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                }
            }
        }

        [Test]
        public void Test_BeginExecute_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.BeginExecute(TestData, null, null));
                }
            }
        }

        [Test]
        public void Test_BeginExecute_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.BeginExecute(TestData, null, null));
                }
            }
        }

        [Test]
        public void Test_EndExecute()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginExecute(TestData, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.EndExecute(ar);
                    Assert.AreEqual(TestData, server.Command);
                }
            }
        }

        [Test]
        public void Test_EndExecute_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginExecute(TestData, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.EndExecute(ar));
                }
            }
        }

        [Test]
        public void Test_Disconnected_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Disconnected += listener.OnEvent;
                    client.Connect();
                    client.Disconnect();
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeDisconnectedEventArgs) listener.Events[0];
                    Assert.IsFalse(args.IsServerInitiated);
                    Assert.IsFalse(args.IsDisposed);
                }
            }
        }

        [Test]
        public void Test_Disconnected_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Disconnected += listener.OnEvent;
                    client.Connect();
                    server.Disconnect();
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeDisconnectedEventArgs) listener.Events[0];
                    Assert.IsTrue(args.IsServerInitiated);
                    Assert.IsFalse(args.IsDisposed);
                }
            }
        }

        [Test]
        public void Test_Disconnected_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Disconnected += listener.OnEvent;
                    client.Connect();
                    client.Dispose();
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeDisconnectedEventArgs) listener.Events[0];
                    Assert.IsFalse(args.IsServerInitiated);
                    Assert.IsTrue(args.IsDisposed);
                }
            }
        }

        [Test]
        public void Test_StartAdvise_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    client.StartAdvise(ItemName, 1, true, Timeout);
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(args.Data));
                    Assert.AreEqual(TestData, args.Text);
                }
            }
        }

        [Test]
        public void Test_StartAdvise_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    client.StartAdvise(ItemName, 1, false, Timeout);
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.IsNull(args.Data);
                    Assert.IsNull(args.Text);
                }
            }
        }

        [Test]
        public void Test_StartAdvise_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    client.StartAdvise(ItemName, 1, true, true, Timeout, "MyStateObject");
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.AreEqual("MyStateObject", args.State);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(args.Data));
                    Assert.AreEqual(TestData, args.Text);
                }
            }
        }

        [Test]
        public void Test_StartAdvise_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.StartAdvise(ItemName, 1, false, Timeout));
                }
            }
        }

        [Test]
        public void Test_StartAdvise_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.StartAdvise(ItemName, 1, false, Timeout));
                }
            }
        }

        [Test]
        public void Test_StartAdvise_After_StartAdvise()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.StartAdvise(ItemName, 1, false, Timeout);
                    Assert.Throws<InvalidOperationException>(() => client.StartAdvise(ItemName, 1, false, Timeout));
                }
            }
        }

        [Test]
        public void Test_BeginStartAdvise_Variation_1()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    var ar = client.BeginStartAdvise(ItemName, 1, true, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(args.Data));
                }
            }
        }

        [Test]
        public void Test_BeginStartAdvise_Variation_2()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    var ar = client.BeginStartAdvise(ItemName, 1, false, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.IsNull(args.Data);
                }
            }
        }

        [Test]
        public void Test_BeginStartAdvise_Variation_3()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    var listener = new EventListener();
                    client.Advise += listener.OnEvent;
                    client.Connect();
                    var ar = client.BeginStartAdvise(ItemName, 1, true, true, null, null, "MyStateObject");
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    server.Advise(TopicName, ItemName);
                    Assert.IsTrue(listener.Received.WaitOne(Timeout, false));
                    var args = (DdeAdviseEventArgs) listener.Events[0];
                    Assert.AreEqual(ItemName, args.Item);
                    Assert.AreEqual(1, args.Format);
                    Assert.AreEqual("MyStateObject", args.State);
                    Assert.AreEqual(TestData, Encoding.ASCII.GetString(args.Data));
                }
            }
        }

        [Test]
        public void Test_BeginStartAdvise_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.BeginStartAdvise(ItemName, 1, false, null, null));
                }
            }
        }

        [Test]
        public void Test_BeginStartAdvise_Before_Connect()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    Assert.Throws<InvalidOperationException>(() => client.BeginStartAdvise(ItemName, 1, false, null, null));
                }
            }
        }

        [Test]
        public void Test_EndStartAdvise()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginStartAdvise(ItemName, 1, true, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.EndStartAdvise(ar);
                }
            }
        }

        [Test]
        public void Test_EndStartAdvise_After_Dispose()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    var ar = client.BeginStartAdvise(ItemName, 1, true, null, null);
                    Assert.IsTrue(ar.AsyncWaitHandle.WaitOne(Timeout, false));
                    client.Dispose();
                    Assert.Throws<ObjectDisposedException>(() => client.EndStartAdvise(ar));
                }
            }
        }

        [Test]
        public void Test_StopAdvise()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    client.StartAdvise(ItemName, 1, true, Timeout);
                    client.StopAdvise(ItemName, Timeout);
                }
            }
        }

        [Test]
        public void Test_StopAdvise_Before_StartAdvise()
        {
            using (var server = new TestServer(ServiceName))
            {
                server.Register();
                server.SetData(TopicName, ItemName, 1, Encoding.ASCII.GetBytes(TestData));
                using (var client = new DdeClient(ServiceName, TopicName))
                {
                    client.Connect();
                    Assert.Throws<InvalidOperationException>(() => client.StopAdvise(ItemName, Timeout));
                }
            }
        }
    } // class
} // namespace