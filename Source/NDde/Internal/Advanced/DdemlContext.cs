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

using System.Runtime.InteropServices;
using System.Text;
using NDde.Internal.Client;
using NDde.Internal.Server;
using NDde.Internal.Utility;
using NDde.Properties;

namespace NDde.Internal.Advanced;

internal sealed class DdemlContext : IDisposable
{
    private static readonly WeakReferenceDictionary<int, DdemlContext> _Instances =
        new();

    private readonly Ddeml.DdeCallback _Callback; // DDEML callback function

    private readonly WeakReferenceDictionary<IntPtr, DdemlClient> _ClientTable =
        new(); // Active clients by conversation

    private readonly List<IDdemlTransactionFilter> _Filters = new(); // ITransactionFilter objects

    private readonly WeakReferenceDictionary<IntPtr, DdemlServer> _ServerTable1 =
        new(); // Active servers by conversation

    private readonly WeakReferenceDictionary<string, DdemlServer> _ServerTable2 =
        new(); // Active servers by service

    public DdemlContext()
    {
        // Create the callback that will be used by the DDEML.
        _Callback = OnDdeCallback;
    }

    public int InstanceId { get; private set; }

    public bool IsInitialized => InstanceId != 0;

    public Encoding Encoding { get; set; } = Encoding.ASCII;

    internal bool IsDisposed { get; private set; }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    internal static DdemlContext GetDefault()
    {
        lock (_Instances)
        {
            var context = _Instances[Ddeml.GetCurrentThreadId()];
            if (context != null) return context;
            context = new DdemlContext();
            _Instances.Add(Ddeml.GetCurrentThreadId(), context);

            return context;
        }
    }

    public event EventHandler<DdemlRegistrationEventArgs> Register;

    public event EventHandler<DdemlRegistrationEventArgs> Unregister;

    internal event EventHandler StateChange;

    ~DdemlContext()
    {
        Dispose(false);
    }

    private void Dispose(bool disposing)
    {
        if (IsDisposed) return;
        IsDisposed = true;
        if (disposing)
        {
            // Dispose all clients.
            foreach (var client in _ClientTable.Values)
                client.Dispose();

            // Dispose all servers.
            foreach (var server in _ServerTable2.Values)
                server.Dispose();

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
        }

        if (!IsInitialized) return;
        // Uninitialize this DDEML instance.
        InstanceManager.Uninitialize(InstanceId);

        // Indicate that this object is no longer initialized.
        InstanceId = 0;
    }

    public void Initialize()
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (IsInitialized)
            throw new InvalidOperationException(Resources.AlreadyInitializedMessage);

        Initialize(Ddeml.APPCLASS_STANDARD);

        // Raise the StateChange event.
        StateChange?.Invoke(this, EventArgs.Empty);
    }

    public void AddTransactionFilter(IDdemlTransactionFilter filter)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (filter == null)
            throw new ArgumentNullException(nameof(filter));
        if (_Filters.Contains(filter))
            throw new InvalidOperationException(Resources.FilterAlreadyAddedMessage);

        _Filters.Add(filter);
    }

    public void RemoveTransactionFilter(IDdemlTransactionFilter filter)
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().ToString());
        if (filter == null)
            throw new ArgumentNullException(nameof(filter));
        if (!_Filters.Contains(filter))
            throw new InvalidOperationException(Resources.FilterNotAddedMessage);

        _Filters.Remove(filter);
    }

    internal void Initialize(int afCmd)
    {
        // Initialize a DDEML instance.
        InstanceId = InstanceManager.Initialize(_Callback, afCmd);

        // If the instance identifier is null then the DDEML could not be initialized.
        if (InstanceId != 0) return;
        var error = Ddeml.DdeGetLastError(InstanceId);
        throw new DdemlException(Resources.InitializeFailedMessage, error);
    }

    internal void RegisterClient(DdemlClient client)
    {
        _ClientTable[client.Handle] = client;
    }

    internal void RegisterServer(DdemlServer server)
    {
        _ServerTable2[server.Service] = server;
    }

    internal void UnregisterClient(DdemlClient client)
    {
        _ClientTable[client.Handle] = null;
    }

    internal void UnregisterServer(DdemlServer server)
    {
        _ServerTable2[server.Service] = null;
    }

    private IntPtr OnDdeCallback(int uType, int uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData,
        IntPtr dwData1, IntPtr dwData2)
    {
        // Create a new transaction object that will be dispatched to a DdemlClient, DdemlServer, or ITransactionFilter.
        var t = new DdemlTransaction(uType, uFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2);

        // Run each transaction filter.
        if (_Filters.Any(filter => filter.PreFilterTransaction(t)))
        {
            return t.dwRet;
        }

        // Dispatch the transaction.
        switch (uType)
        {
            case Ddeml.XTYP_ADVDATA:
            {
                var client = _ClientTable[hConv];
                if (client != null)
                    if (client.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_ADVREQ:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_ADVSTART:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_ADVSTOP:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_CONNECT:
            {
                // Get the service name from the hsz2 string handle.
                var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
                var length = Ddeml.DdeQueryString(InstanceId, hsz2, psz, psz.Capacity,
                    Ddeml.CP_WINANSI);
                var service = psz.ToString();

                var server = _ServerTable2[service];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_CONNECT_CONFIRM:
            {
                // Get the service name from the hsz2 string handle.
                var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
                var length = Ddeml.DdeQueryString(InstanceId, hsz2, psz, psz.Capacity,
                    Ddeml.CP_WINANSI);
                var service = psz.ToString();

                var server = _ServerTable2[service];
                if (server != null)
                {
                    _ServerTable1[hConv] = server;
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                }

                break;
            }
            case Ddeml.XTYP_DISCONNECT:
            {
                var client = _ClientTable[hConv];
                if (client != null)
                {
                    _ClientTable[hConv] = null;
                    if (client.ProcessCallback(t))
                        return t.dwRet;
                }

                var server = _ServerTable1[hConv];
                if (server != null)
                {
                    _ServerTable1[hConv] = null;
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                }

                break;
            }
            case Ddeml.XTYP_EXECUTE:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_POKE:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_REQUEST:
            {
                var server = _ServerTable1[hConv];
                if (server != null)
                    if (server.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_XACT_COMPLETE:
            {
                var client = _ClientTable[hConv];
                if (client != null)
                    if (client.ProcessCallback(t))
                        return t.dwRet;
                break;
            }
            case Ddeml.XTYP_WILDCONNECT:
            {
                // This library does not support wild connects.
                return IntPtr.Zero;
            }
            case Ddeml.XTYP_MONITOR:
            {
                // Monitors are handled separately in DdemlMonitor.
                return IntPtr.Zero;
            }
            case Ddeml.XTYP_ERROR:
            {
                // Get the error code, but do nothing with it at this time.
                var error = dwData1.ToInt32();

                return IntPtr.Zero;
            }
            case Ddeml.XTYP_REGISTER:
            {
                if (Register == null) return IntPtr.Zero;
                // Get the service name from the hsz1 string handle.
                var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
                var length =
                    Ddeml.DdeQueryString(InstanceId, hsz1, psz, psz.Capacity,
                        Ddeml.CP_WINANSI);
                var service = psz.ToString();

                Register(this, new DdemlRegistrationEventArgs(service));

                return IntPtr.Zero;
            }
            case Ddeml.XTYP_UNREGISTER:
            {
                if (Unregister == null) return IntPtr.Zero;
                // Get the service name from the hsz1 string handle.
                var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
                var length =
                    Ddeml.DdeQueryString(InstanceId, hsz1, psz, psz.Capacity,
                        Ddeml.CP_WINANSI);
                var service = psz.ToString();

                Unregister(this, new DdemlRegistrationEventArgs(service));

                return IntPtr.Zero;
            }
        }

        return IntPtr.Zero;
    }

    /// <summary>
    ///     This class is needed to dispose of DDEML resources correctly since the DDEML is thread specific.
    /// </summary>
    private sealed class InstanceManager : IMessageFilter
    {
        private const int WM_APP = 0x8000;

        private static readonly string DataSlot = typeof(InstanceManager).FullName;

        private static readonly IDictionary<int, int> _Table = new Dictionary<int, int>();

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_APP + 1)
                Ddeml.DdeUninitialize(m.WParam.ToInt32());
            return false;
        }

        [DllImport("user32.dll")]
        private static extern void PostThreadMessage(int idThread, int Msg, IntPtr wParam,
            IntPtr lParam);

        public static int Initialize(Ddeml.DdeCallback pfnCallback, int afCmd)
        {
            lock (_Table)
            {
                // Initialize a DDEML instance.
                var instanceId = 0;
                Ddeml.DdeInitialize(ref instanceId, pfnCallback, afCmd, 0);

                if (instanceId == 0) return instanceId;
                // Make sure this thread has an IMessageFilter on it.
                var slot = Thread.GetNamedDataSlot(DataSlot);
                if (Thread.GetData(slot) == null)
                {
                    var filter = new InstanceManager();
                    Application.AddMessageFilter(filter);
                    Thread.SetData(slot, filter);
                }

                // Add an entry to the table that maps the instance identifier to the current thread.
                _Table.Add(instanceId, Ddeml.GetCurrentThreadId());

                return instanceId;
            }
        }

        public static void Uninitialize(int instanceId)
        {
            // This method could be called by the GC finalizer thread.  If it is then a direct call to the DDEML will fail since the DDEML is
            // thread specific.  A message will be posted to the DDEML thread instead.
            lock (_Table)
            {
                if (!_Table.ContainsKey(instanceId)) return;
                // Determine if the current thread matches what is in the table.
                var threadId = _Table[instanceId];
                if (threadId == Ddeml.GetCurrentThreadId())
                    Ddeml.DdeUninitialize(instanceId);
                else
                    PostThreadMessage(threadId, WM_APP + 1, new IntPtr(instanceId),
                        IntPtr.Zero);

                // Remove the instance identifier from the table because it is no longer in use.
                _Table.Remove(instanceId);
            }
        }
    } // class
} // class
// namespace