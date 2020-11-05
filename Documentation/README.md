# NDde

Copyright © 2005-2006 by Brian Gideon (briangideon@yahoo.com)

This library provides a convenient and easy way to integrate .NET applications with legacy applications that use Dynamic Data Exchange (DDE). DDE is an older interprocess communication protocol that relies heavily on passing windows messages back and forth between applications. Other, more modern and robust, techniques for interprocess communication are available and should be used when available. This library is only intended to be used when no other alternatives exist. In other words, do not use this library to read or write data in Excel.

## Changelog

* 2.01.0563 – 8/03/2006

1. The default implementation of DdeServer.OnBeforeConnect will now accept all connections.
2. The default implementation of DdeServer.OnStartAdvise will now accept all advise loops.
3. The DdeContext.Encoding property was added to allow callers to specify the default character encoding used.
4. The last byte of the command string in the DdeServer.OnExecute method may be stripped. This has been fixed.
5. DdeMonitor and associated classes have been added.
6. DdeClient.StopAdvise was throwing a DdeException instead of an InvalidOperationException when an advise loop did not exist. This has been fixed.
7. DdeServer can now return zero length byte arrays. Previously, a zero length byte array would send a DDE_FNOTPROCESSED message to the client.
8. A parameter was added to DdeClient.StartAdvise and DdeClient.BeginStartAdvise that allows the caller to specify an application defined state object that is associated with the advise loop.
9. DdeServer.IsRegistered does not return the correct value. This has been fixed.
10. TryConnect, TryExecute, TryPoke, and TryRequest methods were added to DdeClient to provide an interface that does not throw exceptions.
11. A parameter was added to DdeClient.StartAdvise and DdeClient.BeginStartAdvise to control whether the client will acknowledge the server after each advisory.
12. DdeMessageLoop was added to provide an ISynchronizeInvoke object that can be used to easily create a message loop on any thread.

• 2.00.0410 - 3/03/2006 - .NET Framework 2.0 Release

1. Initial release.

• 1.04.0282 - 10/26/2005

1. DdeClient, DdeServer, and DdeContext instance methods are now thread-safe.
2. If DdeContext created its own thread for message pumping it will now create the hidden window using the HWND_MESSAGE parameter. That will make it a message only window. This is only done on WinXP or higher.
3. DdeContext will throw an exception if it cannot marshal an operation in a timely manner. The default timeout is set to 60 seconds.
4. DdeException is now serializable so that it can cross application domain boundaries.

• 1.03.0171 - 7/07/2005

1. During a warm advise notification the DdeClient.AdviseEventArgs.Data property returns a 0 length byte array. It should be returning null according to the online documentation. This has been fixed.
2. The DdeClient.Conversation property has been removed. The Service, Topic, Handle, and IsPaused properties are now directly available on DdeClient.
3. DdeClient.Connect and DdeClient.Disconnect methods have been added. The Connect method must be called to established a conversation with the server. The DdeClient constructor no longer does this automatically.
4. The DdeClient.Disconnected event is now raised when DdeClient.Disconnect or DdeClient.Dispose are called.
5. The DdeClient.Disconnected event now contains the DisconnectedEventArgs parameter which has the IsServerInitiated and IsDisposed properties that indicate whether the server terminated the conversation and whether the DdeClient object has been disposed respectively.
6. The DdeClient.InstanceId, DdeContext.Transaction.uType, and DdeContext.Transaction.uFmt properties are now Int32 instead of UInt32. This was done so that the library can be CLS compliant.
7. The DdeClient will now initiate an advise loop with a flag that instructs the server to wait for an acknowledgement from the client before sending another notification. This will prevent the server from sending data faster than the client can process it.
8. DdeServer.Register and DdeServer.Unregister methods have been added. The Register method must be called to register the service. The DdeServer constructor no longer does this automatically.
9. The DdeServer.IsRegistered property has been added.
10. The DdeServer.OnAdvise method signature has changed. The first parameter is the topic name instead of a conversation. The method now only executes once per call to Advise regardless of how many conversations are involved.
11. A DdeServer can now return a TooBusy value if the server is too busy to process the OnExecute or OnPoke events.
12. The DdeServer.Advise method now accepts a single asterix for the topic name or item name. When an asterix is used for the topic name, item name, or both all active advise loops will be notified.
13. The DdeContext.Initialize method has been added. This method typically does not need to be called explicitly.
14. The DdeContext.RemoveTransactionFilter method has been added.
15. The DdeContext.Transaction.dwData1 and DdeContext.Transaction.dwData2 properties are now IntPtr instead of UInt32. This was done to correctly match the types as they are declared in the DDEML callback function.
16. The DdeContext.Transaction.Result property has been renamed to dwRet.
17. DdeConversation has been renamed to Conversation and moved to a nested class under DdeServer.
18. The DdeException.Code property has been added. It contains an error code if one is returned by the DDEML.
19. DdeException now subclasses Exception instead of ApplicationException.
20. The wording on some DdeException messages has changed.
21. The online documentation is significantly enhanced.

• 1.02.0089 - 4/16/2005

1. Developers can now intercept the DDEML callback function by adding transaction filters to a DdeContext. When used in conjunction with the DdeContext.InstanceId and DdeConversation.Handle properties transaction filters give the developer complete control over the DDEML and the ability to call any DDEML function so that more advanced DDE functionality can be used.

• 1.01.0070 - 3/28/2005

1. An exception is thrown when a DDE server disconnects after the garbage collector finalizes a DdeClient. This has been fixed.

• 1.01.0047 - 3/05/2005

1. The DdeClient.EndXXX methods could deadlock if the the object gets disconnected. This has been fixed.
2. A DDEML string handle was not being freed after a DdeServer object was disposed. This has been fixed.
3. The Dispose(bool disposing) method on DdeServer was changed from private to protected virtual so that the IDisposable pattern can be implemented correctly in subclasses.
4. The DdeClient.Conversation property returned a new reference on each use. This could create subtle problems depending on how the caller used the reference. This has been fixed.
5. The DdeClient.Pause and DdeClient.Resume methods will now throw an exception if the object is not connected. This makes these methods consistent with the others.
6. The samples have been updated so that they append a null character to any text data sent through the library.
7. The DdeConversation.Handle property was added to get the DDEML handle of the conversation.
8. The DdeContext.InstanceId property was added to get the DDEML instance identifier.
9. The ToString method of the DdeConversation will now return information about the object.
10. The installer will now put the library in the Global Assembly Cache (GAC).
11. The installer will now create a shortcut to the sample projects in the start menu.

• 1.00.0000 - 1/17/2005

1. Initial release.
