#region Copyright (c) 2005 by Brian Gideon (briangideon@yahoo.com)

/* Shared Source License for NDde
 *
 * This license governs use of the accompanying software ('Software'), and your use of the Software constitutes acceptance of this license.
 *
 * You may use the Software for any commercial or noncommercial purpose, including distributing derivative works.
 *
 * In return, we simply require that you agree:
 *  1. Not to remove any copyright or other notices from the Software.
 *  2. That if you distribute the Software in source code form you do so only under this license (i.e. you must include a complete copy of this
 *     license with your distribution), and if you distribute the Software solely in object form you only do so under a license that complies with
 *     this license.
 *  3. That the Software comes "as is", with no warranties.  None whatsoever.  This means no express, implied or statutory warranty, including
 *     without limitation, warranties of merchantability or fitness for a particular purpose or any warranty of title or non-infringement.  Also,
 *     you must pass this disclaimer on whenever you distribute the Software or derivative works.
 *  4. That no contributor to the Software will be liable for any of those types of damages known as indirect, special, consequential, or incidental
 *     related to the Software or this license, to the maximum extent the law permits, no matter what legal theory it's based on.  Also, you must
 *     pass this limitation of liability on whenever you distribute the Software or derivative works.
 *  5. That if you sue anyone over patents that you think may apply to the Software for a person's use of the Software, your license to the Software
 *     ends automatically.
 *  6. That the patent rights, if any, granted in this license only apply to the Software, not to any derivative works you make.
 *  7. That the Software is subject to U.S. export jurisdiction at the time it is licensed to you, and it may be subject to additional export or
 *     import laws in other places.  You agree to comply with all such laws and regulations that may apply to the Software after delivery of the
 *     software to you.
 *  8. That if you are an agency of the U.S. Government, (i) Software provided pursuant to a solicitation issued on or after December 1, 1995, is
 *     provided with the commercial license rights set forth in this license, and (ii) Software provided pursuant to a solicitation issued prior to
 *     December 1, 1995, is provided with "Restricted Rights" as set forth in FAR, 48 C.F.R. 52.227-14 (June 1987) or DFAR, 48 C.F.R. 252.227-7013
 *     (Oct 1988), as applicable.
 *  9. That your rights under this License end automatically if you breach it in any way.
 * 10. That all rights not expressly granted to you in this license are reserved.
 */

#endregion

using System.ComponentModel;
using NDde.Advanced;
using NDde.Internal;
using NDde.Internal.Client;

namespace NDde.Client;

/// <summary>
///     This represents the client side of a DDE conversation.
/// </summary>
/// <threadsafety static="true" instance="true" />
/// <remarks>
///     <para>
///         DDE conversations are established by specifying a service name and topic name pair.  The service name is
///         usually the name of the application
///         acting as a DDE server.  A DDE server can respond to multiple service names, but most servers usually only
///         respond to one.  The topic name
///         is a logical context for data and is defined by the server application.  A server can and usually does support
///         many topic names.
///     </para>
///     <para>
///         After a conversation has been established by calling <c>Connect</c> an application can read and write data
///         using the <c>Request</c> and
///         <c>Poke</c> methods respectively by specifying an item name supported by the active conversation.  An item name
///         identifies a unit of data.
///         An application can also be notified of changes by initiating an advise loop on an item name using the
///         <c>StartAdvise</c> method.  Advise
///         loops can either be warm or hot.  A hot advise loop returns the data associated with an item name when it
///         changes whereas a warm advise loop
///         only notifies the application without sending any data.  Commands can be sent to the server using the
///         <c>Execute</c> method.
///     </para>
///     <para>
///         Callbacks and events are invoked on the thread hosting the <c>DdeContext</c>.  All operations must be marshaled
///         onto the thread hosting the
///         <c>DdeContext</c> associated with this object.  Method calls will block until that thread becomes available.
///         An exception will be generated
///         if the thread does not become available in a timely manner.
///     </para>
/// </remarks>
public class DdeClient : IDisposable, IDdeClient
{
    private readonly object _LockObject = new();
    //internal static EventLog EventLogWriter =
    //CreateEventsLogger.CreaterEventLogger("NDDE Events", "NdDeEventsLog");

    private EventHandler<DdeAdviseEventArgs> _AdviseEvent;
    private DdeContext _Context;

    private DdemlClient _DdemlObject; // This has lazy initialization through a property.
    private EventHandler<DdeDisconnectedEventArgs> _DisconnectedEvent;

    private IntPtr _Handle = IntPtr.Zero; // This is a cached DdemlClient property.
    private bool _IsConnected; // This is a cached DdemlClient property.
    private bool _IsPaused; // This is a cached DdemlClient property.
    private string _Service = ""; // This is a cached DdemlClient property.
    private string _Topic = ""; // This is a cached DdemlClient property.

    /// <overloads>
    ///     <summary>
    ///     </summary>
    /// </overloads>
    /// <summary>
    ///     This initializes a new instance of the <c>DdeClient</c> class that can connect to a server that supports the
    ///     specified service name and
    ///     topic name pair.
    /// </summary>
    /// <param name="service">
    ///     A service name supported by a server application.
    /// </param>
    /// <param name="topic">
    ///     A topic name support by a server application.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when servic or topic exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when service or topic is a null reference.
    /// </exception>
    public DdeClient(string service, string topic)
        : this(service, topic, DdeContext.GetDefault())
    {
    }

    /// <summary>
    ///     This initializes a new instance of the <c>DdeClient</c> class that can connect to a server that supports the
    ///     specified service name and
    ///     topic name pair using the specified synchronizing object.
    /// </summary>
    /// <param name="service">
    ///     A service name supported by a server application.
    /// </param>
    /// <param name="topic">
    ///     A topic name support by a server application.
    /// </param>
    /// <param name="synchronizingObject">
    ///     The synchronizing object to use for this instance.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when service or topic exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when service or topic is a null reference.
    /// </exception>
    public DdeClient(string service, string topic, ISynchronizeInvoke synchronizingObject)
        : this(service, topic, DdeContext.GetDefault(synchronizingObject))
    {
    }

    /// <summary>
    ///     This initializes a new instance of the <c>DdeClient</c> class that can connect to a server that supports the
    ///     specified service name and
    ///     topic name pair and uses the specified context.
    /// </summary>
    /// <param name="service">
    ///     A service name supported by a server application.
    /// </param>
    /// <param name="topic">
    ///     A topic name support by a server application.
    /// </param>
    /// <param name="context">
    ///     The context to use for execution.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when servic or topic exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when service or topic is a null reference.
    /// </exception>
    public DdeClient(string service, string topic, DdeContext context)
    {
        Service = service;
        Topic = topic;
        Context = context;
    }

    /// <summary>
    /// </summary>
    internal DdemlClient DdemlObject
    {
        get
        {
            lock (_LockObject)
            {
                if (_DdemlObject != null) return _DdemlObject;
                _DdemlObject = new DdemlClient(Service, Topic, Context.DdemlObject);
                _DdemlObject.Advise += OnAdviseReceived;
                _DdemlObject.Disconnected += OnDisconnected;
                _DdemlObject.StateChange += OnStateChange;

                return _DdemlObject;
            }
        }
    }

    /// <summary>
    ///     This gets the context associated with this instance.
    /// </summary>
    public virtual DdeContext Context
    {
        get
        {
            lock (_LockObject)
            {
                return _Context;
            }
        }
        private set
        {
            lock (_LockObject)
            {
                _Context = value;
            }
        }
    }

    /// <summary>
    ///     This gets the service name associated with this conversation.
    /// </summary>
    public virtual string Service
    {
        get
        {
            lock (_LockObject)
            {
                return _Service;
            }
        }
        private set
        {
            lock (_LockObject)
            {
                _Service = value;
            }
        }
    }

    /// <summary>
    ///     This gets the topic name associated with this conversation.
    /// </summary>
    public virtual string Topic
    {
        get
        {
            lock (_LockObject)
            {
                return _Topic;
            }
        }
        private set
        {
            lock (_LockObject)
            {
                _Topic = value;
            }
        }
    }

    /// <summary>
    ///     This gets the DDEML handle associated with this conversation.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         This can be used in any DDEML function requiring a conversation handle.
    ///     </para>
    ///     <para>
    ///         <note type="caution">
    ///             Incorrect usage of the DDEML can cause this object to function incorrectly and can lead to resource leaks.
    ///         </note>
    ///     </para>
    /// </remarks>
    public virtual IntPtr Handle
    {
        get
        {
            lock (_LockObject)
            {
                return _Handle;
            }
        }
    }

    /// <summary>
    ///     This gets a bool indicating whether this conversation is paused.
    /// </summary>
    public virtual bool IsPaused
    {
        get
        {
            lock (_LockObject)
            {
                return _IsPaused;
            }
        }
    }

    /// <summary>
    ///     This gets a bool indicating whether the conversation is established.
    /// </summary>
    /// <remarks>
    ///     <note type="caution">
    ///         Do not assume that the conversation is still established after checking this property.  The conversation can
    ///         terminate at any time.
    ///     </note>
    /// </remarks>
    public virtual bool IsConnected
    {
        get
        {
            lock (_LockObject)
            {
                return _IsConnected;
            }
        }
    }

    /// <summary>
    ///     This terminates the current conversation and releases all resources held by this instance.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
    }

    /// <summary>
    ///     This is raised when the data has changed for an item name that has an advise loop.
    /// </summary>
    public event EventHandler<DdeAdviseEventArgs> Advise
    {
        add
        {
            lock (_LockObject)
            {
                _AdviseEvent += value;
            }
        }
        remove
        {
            lock (_LockObject)
            {
                _AdviseEvent -= value;
            }
        }
    }

    /// <summary>
    ///     This is raised when the client has been disconnected.
    /// </summary>
    public event EventHandler<DdeDisconnectedEventArgs> Disconnected
    {
        add
        {
            lock (_LockObject)
            {
                _DisconnectedEvent += value;
            }
        }
        remove
        {
            lock (_LockObject)
            {
                _DisconnectedEvent -= value;
            }
        }
    }

    /// <summary>
    ///     This contains the implementation to release all resources held by this instance.
    /// </summary>
    /// <param name="disposing">
    ///     True if called by Dispose, false otherwise.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (!disposing) return;

        void InnerDispose()
        {
            DdemlObject.Dispose();
        }

        try
        {
            Context.Invoke(InnerDispose);
        }
        catch
        {
            // Swallow any exception that occurs.
        }
    }

    /// <summary>
    ///     This establishes a conversation with a server that supports the specified service name and topic name pair.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is already connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the client could not connect to the server.
    /// </exception>
    public virtual void Connect()
    {
        void InnerConnect()
        {
            DdemlObject.Connect();
        }

        try
        {
            //Thread.Sleep(1000);
            Context.Invoke(InnerConnect);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task ConnectAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.StartNew(Connect, cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This establishes a conversation with a server that supports the specified service name and topic name pair.
    /// </summary>
    /// <returns>
    ///     Zero if the operation succeed or non-zero if the operation failed.
    /// </returns>
    public virtual int TryConnect()
    {
        var result = 0;

        void InnerTryConnect()
        {
            result = DdemlObject.TryConnect();
        }

        try
        {
            Context.Invoke(InnerTryConnect);
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This terminates the current conversation.
    /// </summary>
    /// <event cref="Disconnected" />
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client was not previously connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thown when the client could not disconnect from the server.
    /// </exception>
    public virtual void Disconnect()
    {
        void InnerDisconnect()
        {
            DdemlObject.Disconnect();
        }

        try
        {
            Context.Invoke(InnerDisconnect);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task DisconnectAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.StartNew(Disconnect, cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This pauses the current conversation.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the conversation is already paused.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the conversation could not be paused or when the client is not connected.
    /// </exception>
    /// <remarks>
    ///     Synchronous operations will timeout if the conversation is paused.  Asynchronous operations can begin, but will not
    ///     complete until the
    ///     conversation has resumed.
    /// </remarks>
    public virtual void Pause()
    {
        void InnerPause()
        {
            DdemlObject.Pause();
        }

        try
        {
            Context.Invoke(InnerPause);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task PauseAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.StartNew(Pause, cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This resumes the current conversation.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the conversation was not previously paused or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the conversation could not be resumed.
    /// </exception>
    public virtual void Resume()
    {
        void InnerResume()
        {
            DdemlObject.Resume();
        }

        try
        {
            Context.Invoke(InnerResume);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task ResumeAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.StartNew(Resume, cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This terminates an asychronous operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call that begins an asynchronous operation.
    /// </param>
    /// <remarks>
    ///     This method does nothing if the asynchronous operation has already completed.
    /// </remarks>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not be abandoned.
    /// </exception>
    public virtual void Abandon(IAsyncResult asyncResult)
    {
        void InnerAbandon()
        {
            if (asyncResult is AsyncResult result)
                DdemlObject.Abandon(result.DdemlAsyncResult);
            else
                DdemlObject.Abandon(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerAbandon);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This sends a command to the server application.
    /// </summary>
    /// <param name="command">
    ///     The command to be sent to the server application.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when command exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when command is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the command.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual void Execute(string command, int timeout = 500)
    {
        void InnerExecute()
        {
            DdemlObject.Execute(command, timeout);
        }

        try
        {
            Context.Invoke(InnerExecute);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This sends a command to the server application.
    /// </summary>
    /// <param name="command">
    ///     The command to be sent to the server application.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <returns>
    ///     Zero if the operation succeed or non-zero if the operation failed.
    /// </returns>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual int TryExecute(string command, int timeout)
    {
        var result = 0;

        void InnerTryExecute()
        {
            result = DdemlObject.TryExecute(command, timeout);
        }

        try
        {
            Context.Invoke(InnerTryExecute);
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task ExecuteAsync(string command, CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.FromAsync(BeginExecute(command, null, null), EndExecute).HandleCancellation(cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This begins an asynchronous operation to send a command to the server application.
    /// </summary>
    /// <param name="command">
    ///     The command to be sent to the server application.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="state">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when command exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when command is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginExecute(string command, AsyncCallback callback, object state)
    {
        var ar = new AsyncResult(Context) {Callback = callback, State = state};

        void InnerBeginExecute()
        {
            ar.DdemlAsyncResult = DdemlObject.BeginExecute(command, OnExecuteComplete, ar);
        }

        try
        {
            Context.Invoke(InnerBeginExecute);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }

        return ar;
    }

    /// <summary>
    ///     This throws any exception that occurred during the asynchronous operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginExecute</c>.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the command.
    /// </exception>
    public virtual void EndExecute(IAsyncResult asyncResult)
    {
        void InnerEndExecute()
        {
            if (asyncResult is AsyncResult result)
                DdemlObject.EndExecute(result.DdemlAsyncResult);
            else
                DdemlObject.EndExecute(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerEndExecute);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <overloads>
    ///     <summary>
    ///     </summary>
    /// </overloads>
    /// <summary>
    ///     This sends data to the server application.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="data">
    ///     The data to send.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item or data is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the data.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual void Poke(string item, string data, int timeout)
    {
        Poke(item, Context.Encoding.GetBytes(data + "\0"), 1, timeout);
    }

    /// <summary>
    ///     This sends data to the server application.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="data">
    ///     The data to send.
    /// </param>
    /// <param name="format">
    ///     The format of the data.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item or data is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the data.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual void Poke(string item, byte[] data, int format, int timeout)
    {
        void InnerPoke()
        {
            DdemlObject.Poke(item, data, format, timeout);
        }

        try
        {
            Context.Invoke(InnerPoke);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This sends data to the server application.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="data">
    ///     The data to send.
    /// </param>
    /// <param name="format">
    ///     The format of the data.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <returns>
    ///     Zero if the operation succeed or non-zero if the operation failed.
    /// </returns>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual int TryPoke(string item, byte[] data, int format, int timeout)
    {
        var result = 0;

        void InnerTryPoke()
        {
            result = DdemlObject.TryPoke(item, data, format, timeout);
        }

        try
        {
            Context.Invoke(InnerTryPoke);
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual Task PokeAsync(string item, string data, CancellationToken cancellationToken = default)
    {
        return PokeAsync(item, Context.Encoding.GetBytes(data + "\0"), 1, cancellationToken);
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task PokeAsync(string item, byte[] data, int format, CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.FromAsync(BeginPoke(item, data, format, null, null), EndPoke).HandleCancellation(cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This begins an asynchronous operation to send data to the server application.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="data">
    ///     The data to send.
    /// </param>
    /// <param name="format">
    ///     The format of the data.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="state">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item or data is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginPoke(string item, byte[] data, int format, AsyncCallback callback,
        object state)
    {
        var ar = new AsyncResult(Context) {Callback = callback, State = state};

        void InnerBeginPoke()
        {
            ar.DdemlAsyncResult = DdemlObject.BeginPoke(item, data, format, OnPokeComplete, ar);
        }

        try
        {
            Context.Invoke(InnerBeginPoke);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }

        return ar;
    }

    /// <summary>
    ///     This throws any exception that occurred during the asynchronous operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginPoke</c>.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the data.
    /// </exception>
    public virtual void EndPoke(IAsyncResult asyncResult)
    {
        void InnerEndPoke()
        {
            if (asyncResult is AsyncResult result)
                DdemlObject.EndPoke(result.DdemlAsyncResult);
            else
                DdemlObject.EndPoke(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerEndPoke);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <overloads>
    ///     <summary>
    ///     </summary>
    /// </overloads>
    /// <summary>
    ///     This requests data using the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <returns>
    ///     The data returned by the server application in CF_TEXT format.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the request.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual string Request(string item, int timeout)
    {
        return Context.Encoding.GetString(Request(item, 1, timeout));
    }

    /// <summary>
    ///     This requests data using the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to return.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <returns>
    ///     The data returned by the server application.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the request.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual byte[] Request(string item, int format, int timeout)
    {
        byte[] result = null;

        void InnerRequest()
        {
            result = DdemlObject.Request(item, format, timeout);
        }

        try
        {
            Context.Invoke(InnerRequest);
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual char[] RequestChars(string item, int timeout)
    {
        return Context.Encoding.GetChars(Request(item, 1, timeout));
    }

    /// <summary>
    ///     This requests data using the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to return.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <param name="data">
    ///     The data returned by the server application.
    /// </param>
    /// <returns>
    ///     Zero if the operation succeeded or non-zero if the operation failed.
    /// </returns>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual int TryRequest(string item, int format, int timeout, out byte[] data)
    {
        byte[] data2 = null;
        var result = 0;

        void InnerTryRequest()
        {
            result = DdemlObject.TryRequest(item, format, timeout, out data2);
        }

        try
        {
            Context.Invoke(InnerTryRequest);
            data = data2;
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task<string> RequestAsync(string item, CancellationToken cancellationToken = default)
    {
        return Context.Encoding.GetString(await RequestAsync(item, 1, cancellationToken));
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task<byte[]> RequestAsync(string item, int format, CancellationToken cancellationToken = default)
    {
        try
        {
            return await Task.Factory.FromAsync(BeginRequest(item, format, null, null), EndRequest).HandleCancellation(cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task<char[]> RequestCharsAsync(string item, CancellationToken cancellationToken = default)
    {
        return Context.Encoding.GetChars(await RequestAsync(item, 1, cancellationToken));
    }

    /// <summary>
    ///     This begins an asynchronous operation to request data using the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to return.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="state">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginRequest(string item, int format, AsyncCallback callback, object state)
    {
        var ar = new AsyncResult(Context) {Callback = callback, State = state};

        void InnerBeginRequest()
        {
            ar.DdemlAsyncResult = DdemlObject.BeginRequest(item, format, OnRequestComplete, ar);
        }

        try
        {
            Context.Invoke(InnerBeginRequest);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }

        return ar;
    }

    /// <summary>
    ///     This gets the data returned by the server application for the operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginRequest</c>.
    /// </param>
    /// <returns>
    ///     The data returned by the server application.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not process the request.
    /// </exception>
    public virtual byte[] EndRequest(IAsyncResult asyncResult)
    {
        byte[] result = null;

        void InnerEndRequest()
        {
            if (asyncResult is AsyncResult resultAsync)
                result = DdemlObject.EndRequest(resultAsync.DdemlAsyncResult);
            else
                result = DdemlObject.EndRequest(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerEndRequest);
            return result;
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <overloads>
    ///     <summary>
    ///     </summary>
    /// </overloads>
    /// <summary>
    ///     This initiates an advise loop on the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to return.
    /// </param>
    /// <param name="hot">
    ///     A bool indicating whether data should be included with the notification.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <event cref="Advise" />
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is already being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not initiate the advise loop.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual void StartAdvise(string item, int format, bool hot, int timeout)
    {
        StartAdvise(item, format, hot, true, timeout, null);
    }

    /// <summary>
    ///     This initiates an advise loop on the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to return.
    /// </param>
    /// <param name="hot">
    ///     A bool indicating whether data should be included with the notification.
    /// </param>
    /// <param name="acknowledge">
    ///     A bool indicating whether the client should acknowledge each advisory before the server will send send another.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <param name="adviseState">
    ///     An application defined data object to associate with this advise loop.
    /// </param>
    /// <event cref="Advise" />
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is already being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not initiate the advise loop.
    /// </exception>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    public virtual void StartAdvise(string item, int format, bool hot, bool acknowledge, int timeout,
        object adviseState)
    {
        void InnerStartAdvice()
        {
            DdemlObject.StartAdvise(item, format, hot, acknowledge, timeout, adviseState);
        }

        try
        {
            Context.Invoke(InnerStartAdvice);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual Task StartAdviseAsync(string item, int format, bool hot, CancellationToken cancellationToken = default)
    {
        return StartAdviseAsync(item, format, hot, true, null, cancellationToken);
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task StartAdviseAsync(string item, int format, bool hot, bool acknowledge, object adviseState, CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.FromAsync(BeginStartAdvise(item, format, hot, acknowledge, null, null, adviseState), EndStartAdvise).HandleCancellation(cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This begins an asynchronous operation to initiate an advise loop on the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to be returned.
    /// </param>
    /// <param name="hot">
    ///     A bool indicating whether data should be included with the notification.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="asyncState">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <event cref="Advise" />
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is already being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginStartAdvise(string item, int format, bool hot, AsyncCallback callback,
        object asyncState)
    {
        return BeginStartAdvise(item, format, hot, true, callback, asyncState, null);
    }

    /// <summary>
    ///     This begins an asynchronous operation to initiate an advise loop on the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name supported by the current conversation.
    /// </param>
    /// <param name="format">
    ///     The format of the data to be returned.
    /// </param>
    /// <param name="hot">
    ///     A bool indicating whether data should be included with the notification.
    /// </param>
    /// <param name="acknowledge">
    ///     A bool indicating whether the client should acknowledge each advisory before the server will send send another.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="asyncState">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <param name="adviseState">
    ///     An application defined data object to associate with this advise loop.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <event cref="Advise" />
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is already being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginStartAdvise(string item, int format, bool hot, bool acknowledge,
        AsyncCallback callback, object asyncState, object adviseState)
    {
        var ar = new AsyncResult(Context) {Callback = callback, State = asyncState};

        void InnerBeginStartAdvise()
        {
            ar.DdemlAsyncResult = DdemlObject.BeginStartAdvise(item, format, hot, acknowledge, OnStartAdviseComplete, ar, adviseState);
        }

        try
        {
            Context.Invoke(InnerBeginStartAdvise);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }

        return ar;
    }

    /// <summary>
    ///     This throws any exception that occurred during the operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginPoke</c>.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not initiate the advise loop.
    /// </exception>
    public virtual void EndStartAdvise(IAsyncResult asyncResult)
    {
        void InnerEndStartAdvise()
        {
            if (asyncResult is AsyncResult result)
                DdemlObject.EndStartAdvise(result.DdemlAsyncResult);
            else
                DdemlObject.EndStartAdvise(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerEndStartAdvise);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This terminates the advise loop for the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name that has an active advise loop.
    /// </param>
    /// <param name="timeout">
    ///     The amount of time in milliseconds to wait for a response.
    /// </param>
    /// <remarks>
    ///     This operation will timeout if the conversation is paused.
    /// </remarks>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters or timeout is negative.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is not being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not terminate the advise loop.
    /// </exception>
    public virtual void StopAdvise(string item, int timeout)
    {
        void InnerStopAdvise()
        {
            DdemlObject.StopAdvise(item, timeout);
        }

        try
        {
            Context.Invoke(InnerStopAdvise);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    ///<inheritdoc cref="IDdeClient"/>
    public virtual async Task StopAdviseAsync(string item, CancellationToken cancellationToken = default)
    {
        try
        {
            await Task.Factory.FromAsync(BeginStopAdvise(item, null, null), EndStopAdvise).HandleCancellation(cancellationToken);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This begins an asynchronous operation to terminate the advise loop for the specified item name.
    /// </summary>
    /// <param name="item">
    ///     An item name that has an active advise loop.
    /// </param>
    /// <param name="callback">
    ///     The delegate to invoke when this operation completes.
    /// </param>
    /// <param name="state">
    ///     An application defined data object to associate with this operation.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     This is thown when item exceeds 255 characters.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when item is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the item is not being advised or when the client is not connected.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the asynchronous operation could not begin.
    /// </exception>
    public virtual IAsyncResult BeginStopAdvise(string item, AsyncCallback callback, object state)
    {
        var ar = new AsyncResult(Context) {Callback = callback, State = state};

        void InnerBeginStopAdvise()
        {
            ar.DdemlAsyncResult = DdemlObject.BeginStopAdvise(item, OnStopAdviseComplete, ar);
        }

        try
        {
            Context.Invoke(InnerBeginStopAdvise);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }

        return ar;
    }

    /// <summary>
    ///     This throws any exception that occurred during the operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginPoke</c>.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     This is thown when asyncResult is an invalid IAsyncResult.
    /// </exception>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when asyncResult is a null reference.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the server does not terminate the advise loop.
    /// </exception>
    public virtual void EndStopAdvise(IAsyncResult asyncResult)
    {
        void InnerEndStopAdvise()
        {
            if (asyncResult is AsyncResult result)
                DdemlObject.EndStopAdvise(result.DdemlAsyncResult);
            else
                DdemlObject.EndStopAdvise(InvalidAsyncResult.Instance);
        }

        try
        {
            Context.Invoke(InnerEndStopAdvise);
        }
        catch (DdemlException e)
        {
            throw new DdeException(e);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    private void OnExecuteComplete(IAsyncResult asyncResult)
    {
        var ar = (AsyncResult) asyncResult.AsyncState;
        ar?.Callback?.Invoke(ar);
    }

    private void OnPokeComplete(IAsyncResult asyncResult)
    {
        var ar = (AsyncResult) asyncResult.AsyncState;
        ar?.Callback?.Invoke(ar);
    }

    private void OnRequestComplete(IAsyncResult asyncResult)
    {
        var ar = (AsyncResult) asyncResult.AsyncState;
        ar?.Callback?.Invoke(ar);
    }

    private void OnStartAdviseComplete(IAsyncResult asyncResult)
    {
        var ar = (AsyncResult) asyncResult.AsyncState;
        ar?.Callback?.Invoke(ar);
    }

    private void OnStopAdviseComplete(IAsyncResult asyncResult)
    {
        var ar = (AsyncResult) asyncResult.AsyncState;
        ar?.Callback?.Invoke(ar);
    }

    private void OnAdviseReceived(object sender, DdemlAdviseEventArgs internalArgs)
    {
        EventHandler<DdeAdviseEventArgs> copy;

        // To make this thread-safe we need to hold a local copy of the reference to the invocation list.  This works because delegates are
        //immutable.
        lock (_LockObject)
        {
            copy = _AdviseEvent;
        }

        copy?.Invoke(this, new DdeAdviseEventArgs(internalArgs, Context.Encoding));
    }

    private void OnDisconnected(object sender, DdemlDisconnectedEventArgs internalArgs)
    {
        EventHandler<DdeDisconnectedEventArgs> copy;

        // To make this thread-safe we need to hold a local copy of the reference to the invocation list.  This works because delegates are
        //immutable.
        lock (_LockObject)
        {
            copy = _DisconnectedEvent;
        }

        copy?.Invoke(this, new DdeDisconnectedEventArgs(internalArgs));
    }

    private void OnStateChange(object sender, EventArgs args)
    {
        lock (_LockObject)
        {
            _Handle = _DdemlObject.Handle;
            _IsConnected = _DdemlObject.IsConnected;
            _IsPaused = _DdemlObject.IsPaused;
            _Service = _DdemlObject.Service;
            _Topic = _DdemlObject.Topic;
        }
    }

    /// <threadsafety static="true" instance="false" />
    private sealed class AsyncResult : IAsyncResult
    {
        private DdeContext _Context;

        public AsyncResult(DdeContext context)
        {
            _Context = context;
        }

        public AsyncCallback Callback { get; set; }

        public object State { get; set; }

        public IAsyncResult DdemlAsyncResult { get; set; }

        public object AsyncState => State;

        public WaitHandle AsyncWaitHandle => DdemlAsyncResult.AsyncWaitHandle;

        public bool CompletedSynchronously => DdemlAsyncResult.CompletedSynchronously;

        public bool IsCompleted => DdemlAsyncResult.IsCompleted;
    } // class

    /// <threadsafety static="true" instance="false" />
    private sealed class InvalidAsyncResult : IAsyncResult
    {
        private InvalidAsyncResult()
        {
        }

        public static InvalidAsyncResult Instance { get; } = new();

        public object AsyncState => null;

        public bool CompletedSynchronously => false;

        public WaitHandle AsyncWaitHandle => null!;

        public bool IsCompleted => false;
    } // class
} // class
// namespace
