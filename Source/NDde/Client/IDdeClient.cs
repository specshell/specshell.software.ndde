using System;
using System.Threading;
using System.Threading.Tasks;
using NDde.Advanced;

namespace NDde.Client
{
    public interface IDdeClient
    {
        /// <summary>
        ///     This gets the context associated with this instance.
        /// </summary>
        DdeContext Context { get; }

        /// <summary>
        ///     This gets the service name associated with this conversation.
        /// </summary>
        string Service { get; }

        /// <summary>
        ///     This gets the topic name associated with this conversation.
        /// </summary>
        string Topic { get; }

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
        IntPtr Handle { get; }

        /// <summary>
        ///     This gets a bool indicating whether this conversation is paused.
        /// </summary>
        bool IsPaused { get; }

        /// <summary>
        ///     This gets a bool indicating whether the conversation is established.
        /// </summary>
        /// <remarks>
        ///     <note type="caution">
        ///         Do not assume that the conversation is still established after checking this property.  The conversation can
        ///         terminate at any time.
        ///     </note>
        /// </remarks>
        bool IsConnected { get; }

        /// <summary>
        ///     This terminates the current conversation and releases all resources held by this instance.
        /// </summary>
        void Dispose();

        /// <summary>
        ///     This is raised when the data has changed for an item name that has an advise loop.
        /// </summary>
        event EventHandler<DdeAdviseEventArgs> Advise;

        /// <summary>
        ///     This is raised when the client has been disconnected.
        /// </summary>
        event EventHandler<DdeDisconnectedEventArgs> Disconnected;

        /// <summary>
        ///     This establishes a conversation with a server that supports the specified service name and topic name pair.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the client is already connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thrown when the client could not connect to the server.
        /// </exception>
        void Connect();

        /// <summary>
        ///     This establishes a conversation with a server that supports the specified service name and topic name pair.
        /// </summary>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the client is already connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thrown when the client could not connect to the server.
        /// </exception>
        Task ConnectAsync(CancellationToken cancellationToken = default);

        /// <summary>
        ///     This establishes a conversation with a server that supports the specified service name and topic name pair.
        /// </summary>
        /// <returns>
        ///     Zero if the operation succeed or non-zero if the operation failed.
        /// </returns>
        int TryConnect();

        /// <summary>
        ///     This terminates the current conversation.
        /// </summary>
        /// <event cref="DdeClient.Disconnected" />
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the client was not previously connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thown when the client could not disconnect from the server.
        /// </exception>
        void Disconnect();

        /// <summary>
        ///     This terminates the current conversation.
        /// </summary>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <event cref="DdeClient.Disconnected" />
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the client was not previously connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thown when the client could not disconnect from the server.
        /// </exception>
        Task DisconnectAsync(CancellationToken cancellationToken = default);

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
        void Pause();

        /// <summary>
        ///     This pauses the current conversation.
        /// </summary>
        /// <param name="cancellationToken">For cancellening the operation</param>
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
        Task PauseAsync(CancellationToken cancellationToken = default);

        /// <summary>
        ///     This resumes the current conversation.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the conversation was not previously paused or when the client is not connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thrown when the conversation could not be resumed.
        /// </exception>
        void Resume();

        /// <summary>
        ///     This resumes the current conversation.
        /// </summary>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <exception cref="InvalidOperationException">
        ///     This is thrown when the conversation was not previously paused or when the client is not connected.
        /// </exception>
        /// <exception cref="DdeException">
        ///     This is thrown when the conversation could not be resumed.
        /// </exception>
        Task ResumeAsync(CancellationToken cancellationToken = default);

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
        void Abandon(IAsyncResult asyncResult);

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
        void Execute(string command, int timeout = 500);

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
        int TryExecute(string command, int timeout);

        /// <summary>
        ///     This begins an asynchronous operation to send a command to the server application.
        /// </summary>
        /// <param name="command">
        ///     The command to be sent to the server application.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
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
        Task ExecuteAsync(string command, CancellationToken cancellationToken = default);

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
        IAsyncResult BeginExecute(string command, AsyncCallback callback, object state);

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
        void EndExecute(IAsyncResult asyncResult);

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
        void Poke(string item, string data, int timeout);

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
        void Poke(string item, byte[] data, int format, int timeout);

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
        int TryPoke(string item, byte[] data, int format, int timeout);

        /// <summary>
        ///     This begins an asynchronous operation to send data to the server application.
        /// </summary>
        /// <param name="item">
        ///     An item name supported by the current conversation.
        /// </param>
        /// <param name="data">
        ///     The data to send.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
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
        Task PokeAsync(string item, string data, CancellationToken cancellationToken = default);

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
        /// <param name="cancellationToken">For cancellening the operation</param>
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
        Task PokeAsync(string item, byte[] data, int format, CancellationToken cancellationToken = default);

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
        IAsyncResult BeginPoke(string item, byte[] data, int format, AsyncCallback callback,
            object state);

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
        void EndPoke(IAsyncResult asyncResult);

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
        string Request(string item, int timeout);

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
        byte[] Request(string item, int format, int timeout);

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
        int TryRequest(string item, int format, int timeout, out byte[] data);

        /// <summary>
        ///     This begins an asynchronous operation to request data using the specified item name.
        /// </summary>
        /// <param name="item">
        ///     An item name supported by the current conversation.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
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
        Task<string> RequestAsync(string item, CancellationToken cancellationToken = default);

        /// <summary>
        ///     This begins an asynchronous operation to request data using the specified item name.
        /// </summary>
        /// <param name="item">
        ///     An item name supported by the current conversation.
        /// </param>
        /// <param name="format">
        ///     The format of the data to return.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
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
        Task<byte[]> RequestAsync(string item, int format, CancellationToken cancellationToken = default);

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
        IAsyncResult BeginRequest(string item, int format, AsyncCallback callback, object state);

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
        byte[] EndRequest(IAsyncResult asyncResult);

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
        /// <event cref="DdeClient.Advise" />
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
        void StartAdvise(string item, int format, bool hot, int timeout);

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
        /// <event cref="DdeClient.Advise" />
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
        void StartAdvise(string item, int format, bool hot, bool acknowledge, int timeout,
            object adviseState);

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
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
        /// </returns>
        /// <event cref="DdeClient.Advise" />
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
        Task StartAdviseAsync(string item, int format, bool hot, CancellationToken cancellationToken = default);

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
        /// <param name="adviseState">
        ///     An application defined data object to associate with this advise loop.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
        /// </returns>
        /// <event cref="DdeClient.Advise" />
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
        Task StartAdviseAsync(string item, int format, bool hot, bool acknowledge, object adviseState, CancellationToken cancellationToken = default);

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
        /// <event cref="DdeClient.Advise" />
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
        IAsyncResult BeginStartAdvise(string item, int format, bool hot, AsyncCallback callback,
            object asyncState);

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
        /// <event cref="DdeClient.Advise" />
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
        IAsyncResult BeginStartAdvise(string item, int format, bool hot, bool acknowledge,
            AsyncCallback callback, object asyncState, object adviseState);

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
        void EndStartAdvise(IAsyncResult asyncResult);

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
        void StopAdvise(string item, int timeout);

        /// <summary>
        ///     This begins an asynchronous operation to terminate the advise loop for the specified item name.
        /// </summary>
        /// <param name="item">
        ///     An item name that has an active advise loop.
        /// </param>
        /// <param name="cancellationToken">For cancellening the operation</param>
        /// <returns>
        ///     An <c>Task</c> object for this operation.
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
        Task StopAdviseAsync(string item, CancellationToken cancellationToken = default);

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
        IAsyncResult BeginStopAdvise(string item, AsyncCallback callback, object state);

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
        void EndStopAdvise(IAsyncResult asyncResult);
    }
}
