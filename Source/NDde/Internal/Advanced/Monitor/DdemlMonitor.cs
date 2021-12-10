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

namespace NDde.Internal.Advanced.Monitor;

[Flags]
internal enum DdemlMonitorFlags
{
    Callback = Ddeml.MF_CALLBACKS,

    Conversation = Ddeml.MF_CONV,

    Error = Ddeml.MF_ERRORS,

    String = Ddeml.MF_HSZ_INFO,

    Link = Ddeml.MF_LINKS,

    Message = Ddeml.MF_POSTMSGS | Ddeml.MF_SENDMSGS,
} // enum

internal sealed class DdemlMonitor : IDisposable
{
    private readonly DdemlContext _Context;
    private bool _Disposed;

    public DdemlMonitor()
        : this(DdemlContext.GetDefault())
    {
    }

    public DdemlMonitor(DdemlContext context)
    {
        _Context = context;
    }

    public void Dispose()
    {
        Dispose(true);
    }

    public event EventHandler<DdemlCallbackActivityEventArgs> CallbackActivity;
    public event EventHandler<DdemlConversationActivityEventArgs> ConversationActivity;
    public event EventHandler<DdemlErrorActivityEventArgs> ErrorActivity;
    public event EventHandler<DdemlLinkActivityEventArgs> LinkActivity;
    public event EventHandler<DdemlMessageActivityEventArgs> MessageActivity;
    public event EventHandler<DdemlStringActivityEventArgs> StringActivity;

    private void Dispose(bool disposing)
    {
        if (_Disposed) return;
        _Disposed = true;
        if (disposing)
            _Context.Dispose();
    }

    public void Start(DdemlMonitorFlags flags)
    {
        _Context.AddTransactionFilter(new TransactionFilter(this));
        _Context.Initialize(Ddeml.APPCLASS_STANDARD | (int) flags);
    }

    private void OnCallback(Ddeml.MONCBSTRUCT mon)
    {
        var args = new DdemlCallbackActivityEventArgs(
            mon.wType,
            mon.wFmt,
            mon.hConv,
            mon.hsz1,
            mon.hsz2,
            mon.hData,
            mon.dwData1,
            mon.dwData2,
            mon.dwRet,
            mon.hTask);

        CallbackActivity?.Invoke(this, args);
    }

    private void OnConversation(Ddeml.MONCONVSTRUCT mon)
    {
        // Get the service name from the hszSvc string handle.
        var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        var length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hszSvc, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var service = psz.ToString();

        // Get the topic name from the hszTopic string handle.
        psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hszTopic, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var topic = psz.ToString();

        var args = new DdemlConversationActivityEventArgs(
            service,
            topic,
            mon.fConnect,
            mon.hConvClient,
            mon.hConvServer,
            mon.hTask);

        ConversationActivity?.Invoke(this, args);
    }

    private void OnError(Ddeml.MONERRSTRUCT mon)
    {
        var args = new DdemlErrorActivityEventArgs(mon.wLastError, mon.hTask);

        ErrorActivity?.Invoke(this, args);
    }

    private void OnLink(Ddeml.MONLINKSTRUCT mon)
    {
        // Get the service name from the hszSvc string handle.
        var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        var length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hszSvc, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var service = psz.ToString();

        // Get the topic name from the hszTopic string handle.
        psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hszTopic, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var topic = psz.ToString();

        // Get the item name from the hszItem string handle.
        psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hszItem, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var item = psz.ToString();

        var args = new DdemlLinkActivityEventArgs(
            service,
            topic,
            item,
            mon.wFmt,
            !mon.fNoData,
            mon.fEstablished,
            mon.fServer,
            mon.hConvClient,
            mon.hConvServer,
            mon.hTask);

        LinkActivity?.Invoke(this, args);
    }

    private void OnPost(Ddeml.MONMSGSTRUCT mon)
    {
        var m = new Message {HWnd = mon.hwndTo, Msg = mon.wMsg, LParam = mon.lParam, WParam = mon.wParam};

        var args = new DdemlMessageActivityEventArgs(DdemlMessageActivityKind.Post, m, mon.hTask);

        MessageActivity?.Invoke(this, args);
    }

    private void OnSend(Ddeml.MONMSGSTRUCT mon)
    {
        var m = new Message {HWnd = mon.hwndTo, Msg = mon.wMsg, LParam = mon.lParam, WParam = mon.wParam};

        var args = new DdemlMessageActivityEventArgs(DdemlMessageActivityKind.Send, m, mon.hTask);

        MessageActivity?.Invoke(this, args);
    }

    private void OnString(Ddeml.MONHSZSTRUCT mon)
    {
        // Get the string from the hsz string handle.
        // TODO: For some reason this does not work correctly.
        var psz = new StringBuilder(Ddeml.MAX_STRING_SIZE);
        var length = Ddeml.DdeQueryString(_Context.InstanceId, mon.hsz, psz, psz.Capacity,
            Ddeml.CP_WINANSI);
        var str = psz.ToString();

        var action = mon.fsAction switch
        {
            Ddeml.MH_CLEANUP => DdemlStringActivityType.CleanUp,
            Ddeml.MH_CREATE => DdemlStringActivityType.Create,
            Ddeml.MH_DELETE => DdemlStringActivityType.Delete,
            Ddeml.MH_KEEP => DdemlStringActivityType.Keep,
            _ => DdemlStringActivityType.CleanUp
        };

        var args = new DdemlStringActivityEventArgs(str, action, mon.hTask);

        StringActivity?.Invoke(this, args);
    }

    private sealed class TransactionFilter : IDdemlTransactionFilter
    {
        private readonly DdemlMonitor _Parent;

        public TransactionFilter(DdemlMonitor parent)
        {
            _Parent = parent;
        }

        public bool PreFilterTransaction(DdemlTransaction t)
        {
            if (t.uType != Ddeml.XTYP_MONITOR) return true;
            var length = 0;
            var phData = Ddeml.DdeAccessData(t.hData, ref length);
            Ddeml.DdeUnaccessData(t.hData);
            switch (t.dwData2.ToInt32())
            {
                case Ddeml.MF_CALLBACKS:
                {
                    // Get the MONCBSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONCBSTRUCT>(phData);

                    _Parent.OnCallback(mon);
                    break;
                }
                case Ddeml.MF_CONV:
                {
                    // Get the MONCONVSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONCONVSTRUCT>(phData);
                    _Parent.OnConversation(mon);
                    break;
                }
                case Ddeml.MF_ERRORS:
                {
                    // Get the MONERRSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONERRSTRUCT>(phData);
                    _Parent.OnError(mon);
                    break;
                }
                case Ddeml.MF_HSZ_INFO:
                {
                    // Get the MONHSZSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONHSZSTRUCT>(phData);
                    _Parent.OnString(mon);
                    break;
                }
                case Ddeml.MF_LINKS:
                {
                    // Get the MONLINKSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONLINKSTRUCT>(phData);
                    _Parent.OnLink(mon);
                    break;
                }
                case Ddeml.MF_POSTMSGS:
                {
                    // Get the MONMSGSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONMSGSTRUCT>(phData);
                    _Parent.OnPost(mon);
                    break;
                }
                case Ddeml.MF_SENDMSGS:
                {
                    // Get the MONMSGSTRUCT object.
                    var mon = Marshal.PtrToStructure<Ddeml.MONMSGSTRUCT>(phData);
                    _Parent.OnSend(mon);
                    break;
                }
            }

            return true;
        }
    } // class
} // class
// namespace