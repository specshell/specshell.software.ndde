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
using System.Runtime.InteropServices;
using System.Text;
using NDde.Internal;
using NDde.Internal.Advanced;
using NDde.Internal.Utility;
using Specshell.WinForm.HiddenForm;

namespace NDde.Advanced;

/// <summary>
///     This provides an execution context for <c>DdeClient</c> and <c>DdeServer</c>.
/// </summary>
/// <threadsafety static="true" instance="true" />
/// <remarks>
///     <para>
///         This class provides a context for DDE activity.  All <c>DdeClient</c> and <c>DdeServer</c> objects must be
///         associated with an instance of
///         this class.  If one is not specified in their constructors then a default instance of this class is used.  This
///         class must be initialized
///         before it can begin sending and receiving DDE messages.  This happens automatically upon its first use by a
///         <c>DdeClient</c> or
///         <c>DdeServer</c>.  An application can call <c>Initialize</c> to make the initialization process occur
///         immediately.  This is useful when a
///         calling application expects this class to raise the <c>Register</c> and <c>Unregister</c> events or invoke the
///         <c>ITransactionFilter.PreFilterTransaction</c> method before being used by a <c>DdeClient</c> or
///         <c>DdeServer</c>.
///     </para>
///     <para>
///         Since forms and controls implement <c>ISynchronizeInvoke</c> they can be used as the synchronizing object for
///         this class.  When an instance
///         of this class is created to use a form or control as the synchronizing object it will use the UI thread for
///         execution.  This is the
///         preferred way of creating an instance of this class when used in a windows application since it avoids
///         multithreaded synchronization issues
///         and cross thread marshaling.  When an instance of this class is created without specifying a synchronizing
///         object it will create and manage
///         its own thread for execution.  This is convenient if you wish to use this library in a console or service
///         application, but with the added
///         cost of cross thread marshaling and the potential for deadlocking application threads.
///     </para>
///     <para>
///         Events are invoked on the thread hosting the <c>DdeContext</c>.  All operations must be marshaled onto the
///         thread hosting the
///         <c>DdeContext</c>.  Method calls will block until that thread becomes available.  An exception will be
///         generated if the thread does not
///         become available in a timely manner.
///     </para>
/// </remarks>
/// <include file='Documentation/Examples.xml' path='Comment/Member[@name="DdeContext"]/*' />
public sealed class DdeContext : IDisposable, ISynchronizeInvoke
{
    //internal static EventLog EventLogWriter = CreateEventsLogger.CreaterEventLogger("NDDE Events", "NdDeEventsLog");
    private static DdeContext _Instance;
    private static readonly object _InstanceLock = new();

    private static readonly WeakReferenceDictionary<ISynchronizeInvoke, DdeContext> _Instances =
        new();

    private readonly object _LockObject = new();

    private DdemlContext _DdemlObject; // This has lazy initialization through a property.
    private Encoding _Encoding; // This is a cached DdemlContext property.

    private int _InstanceId; // This is a cached DdemlContext property.
    private bool _IsInitialized; // This is a cached DdemlContext property.

    private EventHandler<DdeRegistrationEventArgs> _RegisterEvent;
    private ISynchronizeInvoke _Synchronizer;
    private EventHandler<DdeRegistrationEventArgs> _UnregisterEvent;

    /// <overloads>
    ///     <summary>
    ///     </summary>
    /// </overloads>
    /// <summary>
    ///     This initializes a new instance of the <c>DdeContext</c> class that uses a dedicated thread for execution.
    /// </summary>
    /// <remarks>
    ///     This constructor is used when you want the context to create and manage its own thread for DDE message pumping.
    /// </remarks>
    public DdeContext()
    {
    }

    /// <summary>
    ///     This initializes a new instance of the <c>DdeContext</c> class that uses the specified synchronizing object for
    ///     execution.
    /// </summary>
    /// <param name="synchronizingObject">
    ///     The synchronizing object to use for execution.
    /// </param>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when synchronizer is a null reference.
    /// </exception>
    /// <remarks>
    ///     This constructor is used when you want the context to use the specified synchronizing object for DDE message
    ///     pumping.  Since forms and
    ///     controls implement <c>ISynchronizeInvoke</c> they can be used as the synchronizing object.  In that case the
    ///     windows application UI
    ///     thread that is hosting the form or control is used.
    /// </remarks>
    public DdeContext(ISynchronizeInvoke synchronizingObject)
    {
        Synchronizer = synchronizingObject;
    }

    private ISynchronizeInvoke Synchronizer
    {
        get
        {
            lock (_LockObject)
            {
                return _Synchronizer ??= new DdeThread();
            }
        }
        set
        {
            lock (_LockObject)
            {
                _Synchronizer = value;
            }
        }
    }

    /// <summary>
    /// </summary>
    internal DdemlContext DdemlObject
    {
        get
        {
            lock (_LockObject)
            {
                if (_DdemlObject != null) return _DdemlObject;
                _DdemlObject = new DdemlContext();
                _DdemlObject.Register += OnRegister;
                _DdemlObject.Unregister += OnUnregister;
                _DdemlObject.StateChange += OnStateChange;

                return _DdemlObject;
            }
        }
    }

    /// <summary>
    ///     This gets the DDEML instance identifier.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         This can be used in any DDEML function requiring an instance identifier.
    ///     </para>
    ///     <para>
    ///         <note type="caution">
    ///             Incorrect usage of the DDEML can cause this library to function incorrectly and can lead to resource leaks.
    ///         </note>
    ///     </para>
    /// </remarks>
    public int InstanceId
    {
        get
        {
            lock (_LockObject)
            {
                return _InstanceId;
            }
        }
    }

    /// <summary>
    ///     This gets a bool indicating whether the context is initialized.
    /// </summary>
    public bool IsInitialized
    {
        get
        {
            lock (_LockObject)
            {
                return _IsInitialized;
            }
        }
    }

    /// <summary>
    ///     This gets or sets the default encoding that is used.
    /// </summary>
    public Encoding Encoding
    {
        get
        {
            lock (_LockObject)
            {
                if (_Encoding != null) return _Encoding;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _Encoding = Encoding.GetEncoding(1252);

                return _Encoding;
            }
        }
        set
        {
            void InnerEncoding()
            {
                DdemlObject.Encoding = value;
                _Encoding = value;
            }

            Invoke(InnerEncoding);
        }
    }

    /// <summary>
    ///     This releases all resources held by this instance.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
    }

    /// <summary>
    ///     This gets a bool indicating whether the caller must use Invoke.
    /// </summary>
    public bool InvokeRequired => Synchronizer.InvokeRequired;

    /// <summary>
    ///     This executes a delegate on the thread hosting this object.
    /// </summary>
    /// <param name="method">
    ///     The delegate to execute.
    /// </param>
    /// <param name="args">
    ///     The arguments to pass to the delegate.
    /// </param>
    /// <returns>
    ///     The object returned by the delegate.
    /// </returns>
    public object Invoke(Delegate method, object[] args)
    {
        return Synchronizer.Invoke(method, args);
    }

    /// <summary>
    ///     This begins an asynchronous operation to execute a delegate on the thread hosting this object.
    /// </summary>
    /// <param name="method">
    ///     The delegate to execute.
    /// </param>
    /// <param name="args">
    ///     The arguments to pass to the delegate.
    /// </param>
    /// <returns>
    ///     An <c>IAsyncResult</c> object for this operation.
    /// </returns>
    public IAsyncResult BeginInvoke(Delegate method, object[] args)
    {
        return Synchronizer.BeginInvoke(method, args);
    }

    /// <summary>
    ///     This returns the object that the delegate returned in the operation.
    /// </summary>
    /// <param name="asyncResult">
    ///     The <c>IAsyncResult</c> object returned by a call to <c>BeginInvoke</c>.
    /// </param>
    /// <returns>
    ///     The object returned by the delegate.
    /// </returns>
    public object EndInvoke(IAsyncResult asyncResult)
    {
        return Synchronizer.EndInvoke(asyncResult);
    }

    internal static DdeContext GetDefault()
    {
        lock (_InstanceLock)
        {
            return _Instance ??= new DdeContext();
        }
    }

    internal static DdeContext GetDefault(ISynchronizeInvoke synchronizingObject)
    {
        if (synchronizingObject == null) return GetDefault();
        lock (_Instances)
        {
            var context = _Instances[synchronizingObject];
            if (context != null) return context;
            if (synchronizingObject is DdeContext ddeContext)
                context = ddeContext;
            else
                context = new DdeContext(synchronizingObject);
            _Instances.Add(synchronizingObject, context);

            return context;
        }
    }

    /// <summary>
    ///     This is raised when a service name has been registered by a server using the DDEML.
    /// </summary>
    /// <remarks>
    ///     This event will not be raised by servers that do not use the DDEML.
    /// </remarks>
    public event EventHandler<DdeRegistrationEventArgs> Register
    {
        add
        {
            lock (_LockObject)
            {
                _RegisterEvent += value;
            }
        }
        remove
        {
            lock (_LockObject)
            {
                _RegisterEvent -= value;
            }
        }
    }

    /// <summary>
    ///     This is raised when a service name has been unregistered by a server using the DDEML.
    /// </summary>
    /// <remarks>
    ///     This event will not be raised by servers that do not use the DDEML.
    /// </remarks>
    public event EventHandler<DdeRegistrationEventArgs> Unregister
    {
        add
        {
            lock (_LockObject)
            {
                _UnregisterEvent += value;
            }
        }
        remove
        {
            lock (_LockObject)
            {
                _UnregisterEvent -= value;
            }
        }
    }

    private void Dispose(bool disposing)
    {
        if (!disposing) return;

        void InnerDispose()
        {
            DdemlObject.Dispose();
        }

        try
        {
            Invoke(InnerDispose);

            // Dispose the synchronizer if it was created internally.
            var synchronizer = Synchronizer as DdeThread;
            synchronizer?.Dispose();
        }
        catch
        {
            // Swallow any exception that occurs.
        }
    }

    /// <summary>
    ///     This initializes the context.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the context is already initialized.
    /// </exception>
    /// <exception cref="DdeException">
    ///     This is thrown when the context could not be initialized.
    /// </exception>
    /// <remarks>
    ///     <para>
    ///         This class must be initialized before it can begin sending and receiving DDE messages.  This happens
    ///         automatically upon its first use by
    ///         a <c>DdeClient</c> or <c>DdeServer</c>.  An application can call <c>Initialize</c> to make the initialization
    ///         process occur immediately.
    ///         This is useful when a calling application expects this class to raise the <c>Register</c> and <c>Unregister</c>
    ///         events or invoke the
    ///         <c>ITransactionFilter.PreFilterTransaction</c> method before being used by a <c>DdeClient</c> or
    ///         <c>DdeServer</c>.
    ///     </para>
    ///     <para>
    ///         If you attempt to use a synchronizer that is not hosted on a thread running a windows message loop an exception
    ///         will be thrown.
    ///     </para>
    ///     <para>
    ///         Explicitly calling this method will allow added <c>ITransactionFilter</c> objects to begin intercepting the
    ///         DDEML callback function.
    ///     </para>
    /// </remarks>
    public void Initialize()
    {
        void InnerInitialize()
        {
            DdemlObject.Initialize();
            _InstanceId = DdemlObject.InstanceId;
            _IsInitialized = DdemlObject.IsInitialized;
        }

        try
        {
            Invoke(InnerInitialize);
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
    ///     This adds a transaction filter to monitor DDE transactions.
    /// </summary>
    /// <param name="filter">
    ///     The implementation of <c>ITransactionFilter</c> that you want to add.
    /// </param>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when filter is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the filter was already added.
    /// </exception>
    /// <remarks>
    ///     <para>
    ///         Transaction filters can be used to intercept the DDEML callback.
    ///     </para>
    ///     <para>
    ///         <note type="caution">
    ///             Incorrect usage of the DDEML can cause this library to function incorrectly and can lead to resource leaks.
    ///         </note>
    ///     </para>
    /// </remarks>
    public void AddTransactionFilter(IDdeTransactionFilter filter)
    {
        void InnerAddTransactionFilter()
        {
            IDdemlTransactionFilter tf = filter == null ? null : new DdemlTransactionFilter(filter);
            DdemlObject.AddTransactionFilter(tf);
        }

        try
        {
            Invoke(InnerAddTransactionFilter);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This removes a transaction filter and stops it from monitoring DDE transactions.
    /// </summary>
    /// <param name="filter">
    ///     The implementation of <c>ITransactionFilter</c> that you want to remove.
    /// </param>
    /// <exception cref="ArgumentNullException">
    ///     This is thrown when filter is a null reference.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     This is thrown when the filter was not previously added.
    /// </exception>
    /// <remarks>
    ///     <para>
    ///         Transaction filters can be used to intercept the DDEML callback.
    ///     </para>
    ///     <para>
    ///         <note type="caution">
    ///             Incorrect usage of the DDEML can cause this library to function incorrectly and can lead to resource leaks.
    ///         </note>
    ///     </para>
    /// </remarks>
    public void RemoveTransactionFilter(IDdeTransactionFilter filter)
    {
        void InnerRemoveTransactionFilter()
        {
            IDdemlTransactionFilter tf = filter == null ? null : new DdemlTransactionFilter(filter);
            DdemlObject.RemoveTransactionFilter(tf);
        }

        try
        {
            Invoke(InnerRemoveTransactionFilter);
        }
        catch (ObjectDisposedException e)
        {
            throw new ObjectDisposedException(GetType().ToString(), e);
        }
    }

    /// <summary>
    ///     This executes a ThreadStart delegate on the thread hosting this object.
    /// </summary>
    /// <param name="method">
    ///     The delegate to execute.
    /// </param>
    internal void Invoke(ThreadStart method)
    {
        Invoke(method, null);
    }

    private void OnRegister(object sender, DdemlRegistrationEventArgs internalArgs)
    {
        EventHandler<DdeRegistrationEventArgs> copy;

        // To make this thread-safe we need to hold a local copy of the reference to the invocation list.  This works because delegates are
        // immutable.
        lock (_LockObject)
        {
            copy = _RegisterEvent;
        }

        copy?.Invoke(this, new DdeRegistrationEventArgs(internalArgs));
    }

    private void OnUnregister(object sender, DdemlRegistrationEventArgs internalArgs)
    {
        EventHandler<DdeRegistrationEventArgs> copy;

        // To make this thread-safe we need to hold a local copy of the reference to the invocation list.  This works because delegates are
        // immutable.
        lock (_LockObject)
        {
            copy = _UnregisterEvent;
        }

        copy?.Invoke(this, new DdeRegistrationEventArgs(internalArgs));
    }

    private void OnStateChange(object sender, EventArgs args)
    {
        lock (_LockObject)
        {
            _InstanceId = _DdemlObject.InstanceId;
            _IsInitialized = _DdemlObject.IsInitialized;
        }
    }

    /// <threadsafety static="true" instance="true" />
    private sealed class DdemlTransactionFilter : IDdemlTransactionFilter
    {
        private readonly IDdeTransactionFilter _OuterFilter;

        public DdemlTransactionFilter(IDdeTransactionFilter filter)
        {
            _OuterFilter = filter;
        }

        public bool PreFilterTransaction(DdemlTransaction t)
        {
            return _OuterFilter.PreFilterTransaction(new DdeTransaction(t));
        }

        public override bool Equals(object obj)
        {
            if (obj is DdemlTransactionFilter target)
                return _OuterFilter.Equals(target._OuterFilter);
            return false;
        }

        public override int GetHashCode()
        {
            return _OuterFilter.GetHashCode();
        }
    } // class

    /// <threadsafety static="true" instance="true" />
    private sealed class DdeThread : IDisposable, ISynchronizeInvoke
    {
        private readonly Form _Form = new HiddenForm();

        private readonly ManualResetEvent _Initialized = new(false);

        private readonly object _LockObject = new();
        private readonly Thread _Thread;

        private int _ThreadId;

        public DdeThread()
        {
            _Form.Load += Form_Load;
            _Thread = new Thread(Run);
            _Thread.SetApartmentState(ApartmentState.STA);
            _Thread.Name = "DdeMessagePump";
            _Thread.IsBackground = true;
        }

        public void Dispose()
        {
            lock (_LockObject)
            {
                if ((_Thread.ThreadState & ThreadState.Unstarted) != 0)
                {
                    _Thread.Start();
                    _Initialized.WaitOne();
                }
            }

            if (InvokeRequired)
            {
                void InnerDispose()
                {
                    _Form.Dispose();
                }

                Invoke((ThreadStart) InnerDispose, null);
            }
            else
            {
                _Form.Dispose();
            }
        }

        public bool InvokeRequired
        {
            get
            {
                lock (_LockObject)
                {
                    return _ThreadId != GetCurrentThreadId();
                }
            }
        }

        public object Invoke(Delegate method, object[] args)
        {
            lock (_LockObject)
            {
                if ((_Thread.ThreadState & ThreadState.Unstarted) != 0)
                {
                    _Thread.Start();
                    _Initialized.WaitOne();
                }
            }

            if (!InvokeRequired) return method.DynamicInvoke(args);
            try
            {
                return _Form.Invoke(method, args);
            }
            catch (InvalidOperationException e)
            {
                if (!_Form.IsHandleCreated)
                    throw new ObjectDisposedException(GetType().ToString(), e);
                throw;
            }
        }

        public IAsyncResult BeginInvoke(Delegate method, object[] args)
        {
            lock (_LockObject)
            {
                if ((_Thread.ThreadState & ThreadState.Unstarted) != 0)
                {
                    _Thread.Start();
                    _Initialized.WaitOne();
                }
            }

            try
            {
                return _Form.BeginInvoke(method, args);
            }
            catch (InvalidOperationException e)
            {
                if (!_Form.IsHandleCreated)
                    throw new ObjectDisposedException(GetType().ToString(), e);
                throw;
            }
        }

        public object EndInvoke(IAsyncResult asyncResult)
        {
            return _Form.EndInvoke(asyncResult);
        }

        [DllImport("user32.dll")]
        private static extern void PostThreadMessage(int idThread, int Msg, IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        private void Run()
        {
            Thread.VolatileWrite(ref _ThreadId, GetCurrentThreadId());
            Application.ThreadException += Application_ThreadException;
            Application.Run(_Form);
        }

        private void Form_Load(object source, EventArgs e)
        {
            _Initialized.Set();
        }

        private void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            // This is here to prevent unhandled exceptions from appearing in a message box.
        } // class
    } // class
} // class
// namespace
