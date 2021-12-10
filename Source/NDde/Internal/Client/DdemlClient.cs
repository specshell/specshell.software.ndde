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

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using NDde.Internal.Advanced;
using NDde.Properties;

namespace NDde.Internal.Client;

internal class DdemlClient : IDisposable
{
    //internal static EventLog EventLogWriter =
    //CreateEventsLogger.CreaterEventLogger("NDDE Events", "NdDeEventsLog");
    private readonly IDictionary<string, AdviseLoop>
        _AdviseLoopTable = new Dictionary<string, AdviseLoop>(); // Active DDEML advise loops

    private readonly IDictionary<int, AsyncResultBase> _AsynchronousTransactionTable =
        new Dictionary<int, AsyncResultBase>(); // Active DDEML transactions

    private readonly DdemlContext _Context; // DDEML instance manager
    private readonly string _Service = ""; // DDEML service name
    private readonly string _Topic = ""; // DDEML topic name
    private IntPtr _ConversationHandle = IntPtr.Zero; // DDEML conversation handle
    private int _InstanceId; // DDEML instance identifier
    private bool _Paused; // DDEML callback enabled?

    public DdemlClient(string service, string topic)
        : this(service, topic, DdemlContext.GetDefault())
    {
    }

    public DdemlClient(string service, string topic, DdemlContext context)
    {
        if (service == null)
            throw new ArgumentNullException(nameof(service));
        if (service.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(service));
        if (topic == null)
            throw new ArgumentNullException(nameof(topic));
        if (topic.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(topic));

        _Service = service;
        _Topic = topic;
        _Context = context ?? throw new ArgumentNullException(nameof(context));
    }

    public virtual string Service => _Service;

    public virtual string Topic => _Topic;

    public virtual IntPtr Handle => _ConversationHandle;

    public virtual bool IsPaused => _Paused;

    public virtual bool IsConnected => _ConversationHandle != IntPtr.Zero;

    internal bool IsDisposed { get; private set; }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    public event EventHandler<DdemlAdviseEventArgs> Advise;

    public event EventHandler<DdemlDisconnectedEventArgs> Disconnected;

    internal event EventHandler StateChange;

    ~DdemlClient()
    {
        Dispose(false);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (IsDisposed) return;
        IsDisposed = true;
        if (disposing)
        {
            if (!IsConnected) return;
            // Terminate the conversation.
            ConversationManager.Disconnect(_ConversationHandle);

            // Assign each active asynchronous transaction an exception so that the EndXXX methods do not deadlock.
            foreach (var arb in _AsynchronousTransactionTable.Values)
                arb.Process(new DdemlException(Resources.NotConnectedMessage));

            // Make sure the asynchronous transaction and advise loop tables are empty.
            _AsynchronousTransactionTable.Clear();
            _AdviseLoopTable.Clear();

            // Unregister this client from the context so that it will not receive DDEML callbacks.
            _Context.UnregisterClient(this);

            // Indicate that this object is no longer connected or paused.
            _Paused = false;
            _ConversationHandle = IntPtr.Zero;
            _InstanceId = 0;

            // Raise the StateChange event.
            if (StateChange != null)
                foreach (var handler in StateChange.GetInvocationList())
                {
                    if (handler is not EventHandler eventHandler) continue;
                    try
                    {
                        eventHandler(this, EventArgs.Empty);
                    }
                    catch
                    {
                        // Swallow any exception that occurs.
                    }
                }

            // Raise the Disconnected event.
            if (Disconnected == null) return;
            foreach (var handler in
                     Disconnected.GetInvocationList())
            {
                if (handler is not EventHandler<DdemlDisconnectedEventArgs> eventHandler) continue;
                try
                {
                    eventHandler(this,
                        new DdemlDisconnectedEventArgs(false, true));
                }
                catch
                {
                    // Swallow any exception that occurs.
                }
            }
        }
        else
        {
            if (IsConnected)
                ConversationManager.Disconnect(_ConversationHandle);
        }
    }

    public virtual void Connect()
    {
        var error = TryConnect();

        switch (error)
        {
            case -1:
                throw new ObjectDisposedException(GetType().ToString());
            case -2:
                throw new InvalidOperationException(Resources.AlreadyConnectedMessage);
            case > Ddeml.DMLERR_NO_ERROR:
            {
                var message = Resources.ConnectFailedMessage;
                message = message.Replace("${service}", _Service);
                message = message.Replace("${topic}", _Topic);

                throw new DdemlException(message, error);
            }
        }
    }

    public virtual int TryConnect()
    {
        if (IsDisposed)
            return -1;
        if (IsConnected)
            return -2;

        // Make sure the context is initialized.
        if (!_Context.IsInitialized)
            _Context.Initialize();

        // Get a local copy of the DDEML instance identifier so that it can be used in the finalizer.
        _InstanceId = _Context.InstanceId;

        // Establish a conversation with a server that supports the service name and topic name pair.
        _ConversationHandle = ConversationManager.Connect(_InstanceId, _Service, _Topic);

        // If the conversation handle is null then the conversation could not be established.
        if (_ConversationHandle == IntPtr.Zero)
            return Ddeml.DdeGetLastError(_InstanceId);

        // Register this client with the context so that it can receive DDEML callbacks.
        _Context.RegisterClient(this);

        // Raise the StateChange event.
        StateChange?.Invoke(this, EventArgs.Empty);

        return Ddeml.DMLERR_NO_ERROR;
    }

    public virtual void Disconnect()
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);

        // Terminate the conversation.
        ConversationManager.Disconnect(_ConversationHandle);

        // Assign each active asynchronous transaction an exception so that the EndXXX methods do not deadlock.
        foreach (var arb in _AsynchronousTransactionTable.Values)
            arb.Process(new DdemlException(Resources.NotConnectedMessage));

        // Make sure the asynchronous transaction and advise loop tables are empty.
        _AsynchronousTransactionTable.Clear();
        _AdviseLoopTable.Clear();

        // Unregister this client from the context so that it will not receive DDEML callbacks.
        _Context.UnregisterClient(this);

        // Indicate that this object is no longer connected or paused.
        _Paused = false;
        _ConversationHandle = IntPtr.Zero;
        _InstanceId = 0;

        // Raise the StateChange event.
        StateChange?.Invoke(this, EventArgs.Empty);

        // Raise the Disconnected event.
        Disconnected?.Invoke(this, new DdemlDisconnectedEventArgs(false, false));
    }

    public virtual void Pause()
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (IsPaused)
            throw new InvalidOperationException(Resources.AlreadyPausedMessage);

        // Disable the DDEML callback.
        var result = Ddeml.DdeEnableCallback(_InstanceId, _ConversationHandle, Ddeml.EC_DISABLE);

        // Check to see if the DDEML callback was disabled.
        if (!result)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            throw new DdemlException(Resources.ClientPauseFailedMessage, error);
        }

        // The DDEML callback was disabled successfully.
        _Paused = true;

        // Raise the StateChange event.
        StateChange?.Invoke(this, EventArgs.Empty);
    }

    public virtual void Resume()
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (!IsPaused)
            throw new InvalidOperationException(Resources.NotPausedMessage);

        // Enable the DDEML callback.
        var result = Ddeml.DdeEnableCallback(_InstanceId, _ConversationHandle, Ddeml.EC_ENABLEALL);

        // Check to see if the DDEML callback was enabled.
        if (!result)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            throw new DdemlException(Resources.ClientResumeFailedMessage, error);
        }

        // The DDEML callback was enabled successfully.
        _Paused = false;

        // Raise the StateChange event.
        StateChange?.Invoke(this, EventArgs.Empty);
    }

    public virtual void Abandon(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (!(asyncResult is AsyncResultBase))
            throw new ArgumentException(Resources.AsyncResultParameterInvalidMessage, nameof(asyncResult));

        var arb = (AsyncResultBase) asyncResult;
        if (arb.IsCompleted) return;
        // Abandon the asynchronous transaction.
        var result =
            Ddeml.DdeAbandonTransaction(_InstanceId, _ConversationHandle, arb.TransactionId);

        // Remove the IAsyncResult from the transaction table.
        if (_AsynchronousTransactionTable.ContainsKey(arb.TransactionId))
            _AsynchronousTransactionTable.Remove(arb.TransactionId);
    }

    public virtual void Execute(string command, int timeout)
    {
        var error = TryExecute(command, timeout);

        switch (error)
        {
            case -1:
                throw new ObjectDisposedException(GetType().ToString());
            case -2:
                throw new InvalidOperationException(Resources.NotConnectedMessage);
            case -3 when command == null:
                throw new ArgumentNullException(nameof(command));
            case -3 when command.Length > Ddeml.MAX_STRING_SIZE:
                throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(command));
            case -3 when timeout <= 0:
                throw new ArgumentException(Resources.TimeoutParameterInvalidMessage, nameof(timeout));
            case > Ddeml.DMLERR_NO_ERROR:
            {
                var message = Resources.ExecuteFailedMessage;
                message = message.Replace("${command}", command);
                throw new DdemlException(message, error);
            }
        }
    }

    public virtual int TryExecute(string command, int timeout)
    {
        if (IsDisposed)
            return -1;
        if (!IsConnected)
            return -2;
        if (command == null)
            return -3;
        if (command.Length > Ddeml.MAX_STRING_SIZE)
            return -3;
        if (timeout <= 0)
            return -3;

        // Convert the command to a byte array with a null terminating character.
        var data = _Context.Encoding.GetBytes(command + "\0");

        // Send the command to the server.
        var returnFlags = 0;
        var result = Ddeml.DdeClientTransaction(
            data,
            data.Length,
            _ConversationHandle,
            IntPtr.Zero,
            Ddeml.CF_TEXT,
            Ddeml.XTYP_EXECUTE,
            timeout,
            ref returnFlags);

        // If the result is null then the server did not process the command.
        return result == IntPtr.Zero ? Ddeml.DdeGetLastError(_InstanceId) : Ddeml.DMLERR_NO_ERROR;
    }

    public virtual IAsyncResult BeginExecute(string command, AsyncCallback callback, object state)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (command == null)
            throw new ArgumentNullException(nameof(command));
        if (command.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(command));

        // Convert the command to a byte array with a null terminating character.
        var data = _Context.Encoding.GetBytes(command + "\0");

        // Send the command to the server.
        var transactionId = 0;
        var result = Ddeml.DdeClientTransaction(
            data,
            data.Length,
            _ConversationHandle,
            IntPtr.Zero,
            Ddeml.CF_TEXT,
            Ddeml.XTYP_EXECUTE,
            Ddeml.TIMEOUT_ASYNC,
            ref transactionId);

        // If the result is null then the asynchronous operation could not begin.
        if (result == IntPtr.Zero)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.ExecuteFailedMessage;
            message = message.Replace("${command}", command);
            throw new DdemlException(message, error);
        }

        // Create an IAsyncResult for this asynchronous operation and add it to the asynchronous transaction table.
        var ar = new ExecuteAsyncResult(this)
        {
            Command = command, Callback = callback, AsyncState = state, TransactionId = transactionId,
        };
        _AsynchronousTransactionTable.Add(transactionId, ar);

        return ar;
    }

    public virtual void EndExecute(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!(asyncResult is ExecuteAsyncResult))
        {
            var message = Resources.AsyncResultParameterInvalidMessage;
            message = message.Replace("${method}", MethodBase.GetCurrentMethod()?.Name);
            throw new ArgumentException(message, nameof(asyncResult));
        }

        var ar = (ExecuteAsyncResult) asyncResult;
        if (!ar.IsCompleted)
            ar.AsyncWaitHandle.WaitOne();
        if (ar.ExceptionObject != null)
            throw ar.ExceptionObject;
    }

    public virtual void Poke(string item, byte[] data, int format, int timeout)
    {
        var error = TryPoke(item, data, format, timeout);

        switch (error)
        {
            case -1:
                throw new ObjectDisposedException(GetType().ToString());
            case -2:
                throw new InvalidOperationException(Resources.NotConnectedMessage);
            case -3 when data == null:
                throw new ArgumentNullException(nameof(data));
            case -3 when item == null:
                throw new ArgumentNullException(nameof(item));
            case -3 when item.Length > Ddeml.MAX_STRING_SIZE:
                throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
            case -3 when timeout <= 0:
                throw new ArgumentException(Resources.TimeoutParameterInvalidMessage, nameof(timeout));
            case > Ddeml.DMLERR_NO_ERROR:
            {
                var message = Resources.PokeFailedMessage;
                message = message.Replace("${service}", _Service);
                message = message.Replace("${topic}", _Topic);
                message = message.Replace("${item}", item);
                throw new DdemlException(message, error);
            }
        }
    }

    public virtual int TryPoke(string item, byte[] data, int format, int timeout)
    {
        if (IsDisposed)
            return -1;
        if (!IsConnected)
            return -2;
        if (data == null)
            return -3;
        if (item == null)
            return -3;
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            return -3;
        if (timeout <= 0)
            return -3;

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        try
        {
            // Create a data handle for the data being poked.
            var dataHandle =
                Ddeml.DdeCreateDataHandle(_InstanceId, data, data.Length, 0, itemHandle, format, 0);

            // If the data handle is null then it could not be created.
            if (dataHandle == IntPtr.Zero)
                return Ddeml.DdeGetLastError(_InstanceId);

            // Send the data to the server.
            var returnFlags = 0;
            var result = Ddeml.DdeClientTransaction(
                dataHandle,
                -1,
                _ConversationHandle,
                itemHandle,
                format,
                Ddeml.XTYP_POKE,
                timeout,
                ref returnFlags);

            // If the result is null then the server did not process the poke.
            if (result == IntPtr.Zero)
                return Ddeml.DdeGetLastError(_InstanceId);
        }
        finally
        {
            // Free the string handle created earlier.
            Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);
        }

        return Ddeml.DMLERR_NO_ERROR;
    }

    public virtual IAsyncResult BeginPoke(string item, byte[] data, int format, AsyncCallback callback,
        object state)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (data == null)
            throw new ArgumentNullException(nameof(data));
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        try
        {
            // Create a data handle for the data being poked.
            var dataHandle =
                Ddeml.DdeCreateDataHandle(_InstanceId, data, data.Length, 0, itemHandle, format, 0);

            // If the data handle is null then it could not be created.
            if (dataHandle == IntPtr.Zero)
            {
                var error = Ddeml.DdeGetLastError(_InstanceId);
                var message = Resources.PokeFailedMessage;
                message = message.Replace("${service}", _Service);
                message = message.Replace("${topic}", _Topic);
                message = message.Replace("${item}", item);
                throw new DdemlException(message, error);
            }

            // Send the data to the server.
            var transactionId = 0;
            var result = Ddeml.DdeClientTransaction(
                dataHandle,
                -1,
                _ConversationHandle,
                itemHandle,
                format,
                Ddeml.XTYP_POKE,
                Ddeml.TIMEOUT_ASYNC,
                ref transactionId);

            // If the result is null then the asynchronous operation could not begin.
            if (result == IntPtr.Zero)
            {
                var error = Ddeml.DdeGetLastError(_InstanceId);
                var message = Resources.PokeFailedMessage;
                message = message.Replace("${service}", _Service);
                message = message.Replace("${topic}", _Topic);
                message = message.Replace("${item}", item);
                throw new DdemlException(message, error);
            }

            // Create an IAsyncResult for the asynchronous operation and add it to the asynchronous transaction table.
            var ar = new PokeAsyncResult(this)
            {
                Item = item,
                Format = format,
                Callback = callback,
                AsyncState = state,
                TransactionId = transactionId,
            };
            _AsynchronousTransactionTable.Add(transactionId, ar);

            return ar;
        }
        finally
        {
            // Free the string handle created earlier.
            Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);
        }
    }

    public virtual void EndPoke(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!(asyncResult is PokeAsyncResult))
        {
            var message = Resources.AsyncResultParameterInvalidMessage;
            message = message.Replace("${method}", MethodBase.GetCurrentMethod().Name);
            throw new ArgumentException(message, nameof(asyncResult));
        }

        var ar = (PokeAsyncResult) asyncResult;
        if (!ar.IsCompleted)
            ar.AsyncWaitHandle.WaitOne();
        if (ar.ExceptionObject != null)
            throw ar.ExceptionObject;
    }

    public virtual byte[] Request(string item, int format, int timeout)
    {
        var error = TryRequest(item, format, timeout, out var data);

        switch (error)
        {
            case -1:
                throw new ObjectDisposedException(GetType().ToString());
            case -2:
                throw new InvalidOperationException(Resources.NotConnectedMessage);
            case -3 when item == null:
                throw new ArgumentNullException(nameof(item));
            case -3 when item.Length > Ddeml.MAX_STRING_SIZE:
                throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
            case -3 when timeout <= 0:
                throw new ArgumentException(Resources.TimeoutParameterInvalidMessage, nameof(timeout));
            case > Ddeml.DMLERR_NO_ERROR:
            {
                var message = Resources.RequestFailedMessage;
                message = message.Replace("${service}", _Service);
                message = message.Replace("${topic}", _Topic);
                message = message.Replace("${item}", item);
                throw new DdemlException(message, error);
            }
            default:
                return data;
        }
    }

    public virtual int TryRequest(string item, int format, int timeout, out byte[] data)
    {
        data = null;

        if (IsDisposed)
            return -1;
        if (!IsConnected)
            return -2;
        if (item == null)
            return -3;
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            return -3;
        if (timeout <= 0)
            return -3;

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // Request the data from the server.
        var returnFlags = 0;
        var dataHandle = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            format,
            Ddeml.XTYP_REQUEST,
            timeout,
            ref returnFlags);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the data handle is null then the server did not process the request.
        if (dataHandle == IntPtr.Zero)
            return Ddeml.DdeGetLastError(_InstanceId);

        // Get the data from the data handle.
        var length = Ddeml.DdeGetData(dataHandle, null, 0, 0);
        data = new byte[length];
        length = Ddeml.DdeGetData(dataHandle, data, data.Length, 0);

        // Free the data handle created by the server.
        Ddeml.DdeFreeDataHandle(dataHandle);

        return Ddeml.DMLERR_NO_ERROR;
    }

    public virtual IAsyncResult BeginRequest(string item, int format, AsyncCallback callback, object state)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // TODO: It might be possible that the request completed synchronously.
        // Request the data from the server.
        var transactionId = 0;
        var result = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            format,
            Ddeml.XTYP_REQUEST,
            Ddeml.TIMEOUT_ASYNC,
            ref transactionId);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the result is null then the asynchronous operation could not begin.
        if (result == IntPtr.Zero)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.RequestFailedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new DdemlException(message, error);
        }

        // Create an IAsyncResult for the asynchronous operation and add it to the asynchronous transaction table.
        var ar = new RequestAsyncResult(this)
        {
            Item = item,
            Format = format,
            Callback = callback,
            AsyncState = state,
            TransactionId = transactionId,
        };
        _AsynchronousTransactionTable.Add(transactionId, ar);

        return ar;
    }

    public virtual byte[] EndRequest(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!(asyncResult is RequestAsyncResult))
        {
            var message = Resources.AsyncResultParameterInvalidMessage;
            message = message.Replace("${method}", MethodBase.GetCurrentMethod()?.Name);
            throw new ArgumentException(message, nameof(asyncResult));
        }

        var ar = (RequestAsyncResult) asyncResult;
        if (!ar.IsCompleted)
            ar.AsyncWaitHandle.WaitOne();
        if (ar.ExceptionObject != null)
            throw ar.ExceptionObject;

        return ar.Data;
    }

    public virtual void StartAdvise(string item, int format, bool hot, bool acknowledge, int timeout,
        object adviseState)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
        if (timeout <= 0)
            throw new ArgumentException(Resources.TimeoutParameterInvalidMessage, nameof(timeout));
        if (_AdviseLoopTable.ContainsKey(item))
        {
            var message = Resources.AlreadyBeingAdvisedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new InvalidOperationException(message);
        }

        // Create a AdviseLoop object to associate with this advise loop and add it to the advise loop table.
        // The object is added to the advise loop table first because an advisory could come in synchronously during the call
        // DdeClientTransaction.  The assumption is that the advise loop will be initiated successfully.  If it is not then the object must
        // be removed from the advise loop table prior to leaving this method.
        var adviseLoop = new AdviseLoop(this) {Item = item, Format = format, State = adviseState};
        _AdviseLoopTable.Add(item, adviseLoop);

        // Determine whether the client should acknowledge an advisory before the server posts another.
        var ack = acknowledge;

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // Initiate an advise loop.
        var type = Ddeml.XTYP_ADVSTART;
        type = !hot ? type | Ddeml.XTYPF_NODATA : type;
        type = ack ? type | Ddeml.XTYPF_ACKREQ : type;
        var returnFlags = 0;
        var result = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            format,
            type,
            timeout,
            ref returnFlags);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the result is null then the server did not initate the advise loop.
        if (result != IntPtr.Zero) return;
        {
            // Remove the AdviseLoop object created earlier from the advise loop table.  It is no longer valid.
            _AdviseLoopTable.Remove(item);

            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.StartAdviseFailedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new DdemlException(message, error);
        }
    }

    public virtual IAsyncResult BeginStartAdvise(string item, int format, bool hot, bool acknowledge,
        AsyncCallback callback, object asyncState, object adviseState)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
        if (_AdviseLoopTable.ContainsKey(item))
        {
            var message = Resources.AlreadyBeingAdvisedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new InvalidOperationException(message);
        }

        // Determine whether the client should acknowledge an advisory before the server posts another.
        var ack = acknowledge;

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // Initiate an advise loop.
        var type = Ddeml.XTYP_ADVSTART;
        type = !hot ? type | Ddeml.XTYPF_NODATA : type;
        type = ack ? type | Ddeml.XTYPF_ACKREQ : type;
        var transactionId = 0;
        var result = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            format,
            type,
            Ddeml.TIMEOUT_ASYNC,
            ref transactionId);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the result is null then the asynchronous operation could not begin.
        if (result == IntPtr.Zero)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.StartAdviseFailedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new DdemlException(message, error);
        }

        // Create an IAsyncResult for the asynchronous operation and add it to the asynchronous transaction table.
        var ar = new StartAdviseAsyncResult(this)
        {
            Item = item,
            Format = format,
            State = adviseState,
            Callback = callback,
            AsyncState = asyncState,
            TransactionId = transactionId,
        };
        _AsynchronousTransactionTable.Add(transactionId, ar);

        return ar;
    }

    public virtual void EndStartAdvise(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!(asyncResult is StartAdviseAsyncResult))
        {
            var message = Resources.AsyncResultParameterInvalidMessage;
            message = message.Replace("${method}", MethodBase.GetCurrentMethod()?.Name);
            throw new ArgumentException(message, nameof(asyncResult));
        }

        var ar = (StartAdviseAsyncResult) asyncResult;
        if (!ar.IsCompleted)
            ar.AsyncWaitHandle.WaitOne();
        if (ar.ExceptionObject != null)
            throw ar.ExceptionObject;
    }

    public virtual void StopAdvise(string item, int timeout)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
        if (timeout <= 0)
            throw new ArgumentException(Resources.TimeoutParameterInvalidMessage, nameof(timeout));
        if (!_AdviseLoopTable.ContainsKey(item))
        {
            var message = Resources.NotBeingAdvisedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new InvalidOperationException(message);
        }

        // Get the advise loop object from the advise loop table.
        var adviseLoop = _AdviseLoopTable[item];

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // Terminate the advise loop.
        var returnFlags = 0;
        var result = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            adviseLoop.Format,
            Ddeml.XTYP_ADVSTOP,
            timeout,
            ref returnFlags);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the result is null then the server could not terminate the advise loop.
        if (result == IntPtr.Zero)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.StopAdviseFailedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new DdemlException(message, error);
        }

        // Remove the advise loop object from the advise loop table.
        _AdviseLoopTable.Remove(item);
    }

    public virtual IAsyncResult BeginStopAdvise(string item, AsyncCallback callback, object state)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!IsConnected)
            throw new InvalidOperationException(Resources.NotConnectedMessage);
        if (item == null)
            throw new ArgumentNullException(nameof(item));
        if (item.Length > Ddeml.MAX_STRING_SIZE)
            throw new ArgumentException(Resources.StringParameterInvalidMessage, nameof(item));
        if (!_AdviseLoopTable.ContainsKey(item))
        {
            var message = Resources.NotBeingAdvisedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new InvalidOperationException(message);
        }

        // Get the advise object from the advise loop table.
        var adviseLoop = _AdviseLoopTable[item];

        // Create a string handle for the item name.
        var itemHandle = Ddeml.DdeCreateStringHandle(_InstanceId, item, Ddeml.CP_WINANSI);

        // Terminate the advise loop.
        var transactionId = 0;
        var result = Ddeml.DdeClientTransaction(
            IntPtr.Zero,
            0,
            _ConversationHandle,
            itemHandle,
            adviseLoop.Format,
            Ddeml.XTYP_ADVSTOP,
            Ddeml.TIMEOUT_ASYNC,
            ref transactionId);

        // Free the string handle created earlier.
        Ddeml.DdeFreeStringHandle(_InstanceId, itemHandle);

        // If the result is null then the asynchronous operation could not begin.
        if (result == IntPtr.Zero)
        {
            var error = Ddeml.DdeGetLastError(_InstanceId);
            var message = Resources.StopAdviseFailedMessage;
            message = message.Replace("${service}", _Service);
            message = message.Replace("${topic}", _Topic);
            message = message.Replace("${item}", item);
            throw new DdemlException(message, error);
        }

        // Create an IAsyncResult for the asyncronous operation and add it to the asynchronous transaction table.
        var ar = new StopAdviseAsyncResult(this)
        {
            Item = item,
            Format = adviseLoop.Format,
            Callback = callback,
            AsyncState = state,
            TransactionId = transactionId,
        };
        _AsynchronousTransactionTable.Add(transactionId, ar);

        return ar;
    }

    public virtual void EndStopAdvise(IAsyncResult asyncResult)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (!(asyncResult is StopAdviseAsyncResult))
        {
            var message = Resources.AsyncResultParameterInvalidMessage;
            message = message.Replace("${method}", MethodBase.GetCurrentMethod().Name);
            throw new ArgumentException(message, nameof(asyncResult));
        }

        var ar = (StopAdviseAsyncResult) asyncResult;
        if (!ar.IsCompleted)
            ar.AsyncWaitHandle.WaitOne();
        if (ar.ExceptionObject != null)
            throw ar.ExceptionObject;
    }

    internal bool ProcessCallback(DdemlTransaction transaction)
    {
        // This is here to alias the transaction object with a shorter variable name.
        var t = transaction;

        switch (t.uType)
        {
            case Ddeml.XTYP_ADVDATA:
            {
                // Get the item name from the hsz2 string handle.
                var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
                var length = Ddeml.DdeQueryString(_InstanceId, t.hsz2, psz, psz.Capacity,
                    Ddeml.CP_WINANSI);
                var item = psz.ToString();

                // Delegate processing to the advise loop object.
                if (!_AdviseLoopTable.ContainsKey(item)) return false;
                t.dwRet = _AdviseLoopTable[item]
                    .Process(t.uType, t.uFmt, t.hConv, t.hsz1, t.hsz2, t.hData,
                        t.dwData1, t.dwData2);
                return true;

                // This transaction could not be processed here.
            }
            case Ddeml.XTYP_XACT_COMPLETE:
            {
                // Get the transaction identifier from dwData1.
                var transactionId = t.dwData1.ToInt32();

                // Get the IAsyncResult from the asynchronous transaction table and delegate processing to it.
                if (!_AsynchronousTransactionTable.ContainsKey(transactionId)) return false;
                var arb = _AsynchronousTransactionTable[transactionId];

                // Remove the IAsyncResult from the asynchronous transaction table.
                _AsynchronousTransactionTable.Remove(arb.TransactionId);

                t.dwRet = arb.Process(t.uType, t.uFmt, t.hConv, t.hsz1, t.hsz2, t.hData,
                    t.dwData1, t.dwData2);
                return true;

                // This transaction could not be processed here.
            }
            case Ddeml.XTYP_DISCONNECT:
            {
                // Assign each active asynchronous transaction an exception so that the EndXXX methods do not deadlock.
                foreach (var arb in _AsynchronousTransactionTable.Values)
                    arb.Process(new DdemlException(Resources.NotConnectedMessage));

                // Make sure the asynchronous transaction and advise loop tables are empty.
                _AsynchronousTransactionTable.Clear();
                _AdviseLoopTable.Clear();

                // Unregister this client from the context so that it will not receive DDEML callbacks.
                _Context.UnregisterClient(this);

                // Indicate that this object is no longer connected or paused.
                _Paused = false;
                _ConversationHandle = IntPtr.Zero;
                _InstanceId = 0;

                // Raise the StateChange event.
                StateChange?.Invoke(this, EventArgs.Empty);

                // Raise the Disconnected event.
                Disconnected?.Invoke(this, new DdemlDisconnectedEventArgs(true, false));

                // Return zero to indicate that there are no problems.
                t.dwRet = IntPtr.Zero;
                return true;
            }
        }

        // This transaction could not be processed here.
        return false;
    }

    /// <summary>
    ///     This class is needed to dispose of DDEML resources correctly since the DDEML is thread specific.
    /// </summary>
    private sealed class ConversationManager : IMessageFilter
    {
        private const int WM_APP = 0x8000;

        private static readonly string DataSlot = typeof(ConversationManager).FullName;

        private static readonly IDictionary<IntPtr, int> _Table = new Dictionary<IntPtr, int>();

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_APP + 2)
                Ddeml.DdeDisconnect(m.WParam);
            return false;
        }

        [DllImport("user32.dll")]
        private static extern void PostThreadMessage(int idThread, int Msg, IntPtr wParam,
            IntPtr lParam);

        public static IntPtr Connect(int instanceId, string service, string topic)
        {
            lock (_Table)
            {
                // Create string handles for the service name and topic name.
                var serviceHandle =
                    Ddeml.DdeCreateStringHandle(instanceId, service, Ddeml.CP_WINANSI);
                var topicHandle =
                    Ddeml.DdeCreateStringHandle(instanceId, topic, Ddeml.CP_WINANSI);

                // Establish a conversation with a server that suppoerts the service name and topic name pair.
                var handle =
                    Ddeml.DdeConnect(instanceId, serviceHandle, topicHandle, IntPtr.Zero);

                // Free the string handles that were created earlier.
                Ddeml.DdeFreeStringHandle(instanceId, topicHandle);
                Ddeml.DdeFreeStringHandle(instanceId, serviceHandle);

                if (handle == IntPtr.Zero) return handle;
                // Make sure this thread has an IMessageFilter on it.
                var slot = Thread.GetNamedDataSlot(DataSlot);
                if (Thread.GetData(slot) == null)
                {
                    var filter = new ConversationManager();
                    Application.AddMessageFilter(filter);
                    Thread.SetData(slot, filter);
                }

                // Add an entry to the table that maps the conversation handle to the current thread.
                _Table.Add(handle, Ddeml.GetCurrentThreadId());

                return handle;
            }
        }

        public static void Disconnect(IntPtr conversationHandle)
        {
            // This method could be called by the GC finalizer thread.  If it is then a direct call to the DDEML will fail since the DDEML is
            // thread specific.  A message will be posted to the DDEML thread instead.
            lock (_Table)
            {
                if (!_Table.ContainsKey(conversationHandle)) return;
                // Determine if the current thread matches what is in the table.
                var threadId = _Table[conversationHandle];
                if (threadId == Ddeml.GetCurrentThreadId())
                    Ddeml.DdeDisconnect(conversationHandle);
                else
                    PostThreadMessage(threadId, WM_APP + 2, conversationHandle,
                        IntPtr.Zero);

                // Remove the conversation handle from the table because it is no longer in use.
                _Table.Remove(conversationHandle);
            }
        }
    } // class

    private sealed class AdviseLoop
    {
        private readonly DdemlClient _Client;

        public AdviseLoop(DdemlClient client)
        {
            _Client = client;
        }

        public string Item { get; set; } = "";

        public int Format { get; set; }

        public object State { get; set; }

        public IntPtr Process(int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData,
            IntPtr dwData1, IntPtr dwData2)
        {
            if (_Client.Advise == null) return new IntPtr(Ddeml.DDE_FACK);
            // Assume this is a warm advise (XTYPF_NODATA).
            byte[] data = null;

            // If the data handle is not null then it is a hot advise.
            if (hData != IntPtr.Zero)
            {
                // Get the data from the data handle.
                var length = Ddeml.DdeGetData(hData, null, 0, 0);
                data = new byte[length];
                length = Ddeml.DdeGetData(hData, data, data.Length, 0);
            }

            // Raise the Advise event.
            _Client.Advise(_Client, new DdemlAdviseEventArgs(Item, Format, State, data));

            // Return DDE_FACK to indicate that are no problems.
            return new IntPtr(Ddeml.DDE_FACK);
        }
    } // class

    private abstract class AsyncResultBase : IAsyncResult
    {
        private readonly ManualResetEvent _CompletionEvent = new(false);

        public AsyncResultBase(DdemlClient client)
        {
            Client = client;
        }

        public AsyncCallback Callback { get; set; }

        public DdemlClient Client { get; }

        public int TransactionId { get; set; }

        public Exception ExceptionObject { get; set; }

        public object AsyncState { get; set; }

        public WaitHandle AsyncWaitHandle => _CompletionEvent;

        public bool CompletedSynchronously => false;

        public bool IsCompleted { get; private set; }

        public void Process(Exception exception)
        {
            ExceptionObject = exception;

            // Mark this IAsyncResult as complete and invoke the callback.
            IsCompleted = true;
            _CompletionEvent.Set();
            Callback?.Invoke(this);
        }

        public IntPtr Process(int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData,
            IntPtr dwData1, IntPtr dwData2)
        {
            // Delegate processing to the concrete class.
            var returnValue =
                ProcessCallback(uType, uFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2);

            // Mark this IAsyncResult as complete and invoke the callback.
            IsCompleted = true;
            _CompletionEvent.Set();
            Callback?.Invoke(this);

            // The return value is sent to the DDEML.
            return returnValue;
        }

        protected virtual IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // The default implementation will return zero to the DDEML.
            return IntPtr.Zero;
        }
    } // class

    private sealed class ExecuteAsyncResult : AsyncResultBase
    {
        public ExecuteAsyncResult(DdemlClient client) : base(client)
        {
        }

        public string Command { get; set; } = "";

        protected override IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // If the data handle is null then the server did not process the command.
            if (hData != IntPtr.Zero) return IntPtr.Zero;
            var message = Resources.ExecuteFailedMessage;
            message = message.Replace("${command}", Command);
            ExceptionObject = new DdemlException(message);

            // Return zero to indicate that there are no problems.
            return IntPtr.Zero;
        }
    } // class

    private sealed class PokeAsyncResult : AsyncResultBase
    {
        public PokeAsyncResult(DdemlClient client) : base(client)
        {
        }

        public string Item { get; set; } = "";

        public int Format { get; set; }

        protected override IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // If the data handle is null then the server did not process the poke.
            if (hData != IntPtr.Zero) return IntPtr.Zero;
            var message = Resources.PokeFailedMessage;
            message = message.Replace("${service}", Client._Service);
            message = message.Replace("${topic}", Client._Topic);
            message = message.Replace("${item}", Item);
            ExceptionObject = new DdemlException(message);

            // Return zero to indicate that there are no problems.
            return IntPtr.Zero;
        }
    } // class

    private sealed class RequestAsyncResult : AsyncResultBase
    {
        public RequestAsyncResult(DdemlClient client) : base(client)
        {
        }

        public byte[] Data { get; set; }

        public string Item { get; set; } = "";

        public int Format { get; set; }

        protected override IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // If the data handle is null then the server did not process the request.
            // TODO: Some servers may process the request, but return null anyway?
            if (hData == IntPtr.Zero)
            {
                var message = Resources.RequestFailedMessage;
                message = message.Replace("${service}", Client._Service);
                message = message.Replace("${topic}", Client._Topic);
                message = message.Replace("${item}", Item);
                ExceptionObject = new DdemlException(message);
            }
            else
            {
                // Get the data from the data handle.
                var length = Ddeml.DdeGetData(hData, null, 0, 0);
                Data = new byte[length];
                length = Ddeml.DdeGetData(hData, Data, Data.Length, 0);
            }

            // Return zero to indicate that there are no problems.
            return IntPtr.Zero;
        }
    } // class

    private sealed class StartAdviseAsyncResult : AsyncResultBase
    {
        public StartAdviseAsyncResult(DdemlClient client) : base(client)
        {
        }

        public string Item { get; set; } = "";

        public int Format { get; set; }

        public object State { get; set; }

        protected override IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // If the data handle is null then the server did not initiate the advise loop.
            if (hData == IntPtr.Zero)
            {
                var message = Resources.StartAdviseFailedMessage;
                message = message.Replace("${service}", Client._Service);
                message = message.Replace("${topic}", Client._Topic);
                message = message.Replace("${item}", Item);
                ExceptionObject = new DdemlException(message);
            }
            else
            {
                // Create a AdviseLoop object to associate with this advise loop and add it to the owner's advise loop table.
                var adviseLoop = new AdviseLoop(Client) {Item = Item, Format = Format, State = State};
                Client._AdviseLoopTable.Add(Item, adviseLoop);
            }

            // Return zero to indicate that there are no problems.
            return IntPtr.Zero;
        }
    } // class

    private sealed class StopAdviseAsyncResult : AsyncResultBase
    {
        public StopAdviseAsyncResult(DdemlClient client) : base(client)
        {
        }

        public string Item { get; set; } = "";

        public int Format { get; set; }

        protected override IntPtr ProcessCallback(
            int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, IntPtr dwData1,
            IntPtr dwData2)
        {
            // If the data handle is null then the server could not terminate the advise loop.
            if (hData == IntPtr.Zero)
            {
                var message = Resources.StopAdviseFailedMessage;
                message = message.Replace("${service}", Client._Service);
                message = message.Replace("${topic}", Client._Topic);
                message = message.Replace("${item}", Item);
                ExceptionObject = new DdemlException(message);
            }
            else
            {
                // Remove the advise object from the owner's advise loop table.
                Client._AdviseLoopTable.Remove(Item);
            }

            // Return zero to indicate that there are no problems.
            return IntPtr.Zero;
        }
    } // class
} // class
// namespace
