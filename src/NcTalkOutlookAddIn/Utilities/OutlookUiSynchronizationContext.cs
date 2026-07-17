// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    // Marshals asynchronous continuations back to Outlook's STA thread without relying on a host-provided context.
    internal sealed class OutlookUiSynchronizationContext : SynchronizationContext, IDisposable
    {
        private readonly Control _marshalControl;
        private readonly int _threadId;
        private bool _disposed;

        internal OutlookUiSynchronizationContext()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException("The Outlook UI synchronization context must be created on an STA thread.");
            }

            _threadId = Thread.CurrentThread.ManagedThreadId;
            _marshalControl = new Control();

            // Outlook does not always install a SynchronizationContext for COM callbacks, so keep a hidden
            // WinForms handle on its message-pump thread as the stable dispatch target.
            if (_marshalControl.Handle == IntPtr.Zero)
            {
                throw new InvalidOperationException("The Outlook UI dispatch handle could not be created.");
            }
        }

        internal int ThreadId
        {
            get { return _threadId; }
        }

        internal Task<T> InvokeAsync<T>(Func<T> callback)
        {
            if (callback == null)
            {
                throw new ArgumentNullException("callback");
            }

            if (Thread.CurrentThread.ManagedThreadId == _threadId)
            {
                return InvokeOnCurrentThread(callback);
            }

            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                Post(
                    _ =>
                    {
                        try
                        {
                            VerifyAccess();
                            completion.TrySetResult(callback());
                        }
                        catch (Exception ex)
                        {
                            completion.TrySetException(ex);
                        }
                    },
                    null);
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }

            return completion.Task;
        }

        public override void Post(SendOrPostCallback callback, object state)
        {
            if (callback == null)
            {
                throw new ArgumentNullException("callback");
            }
            if (_disposed || _marshalControl.IsDisposed)
            {
                throw new ObjectDisposedException("OutlookUiSynchronizationContext");
            }

            _marshalControl.BeginInvoke(callback, state);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            VerifyAccess();
            _disposed = true;
            _marshalControl.Dispose();
        }

        private Task<T> InvokeOnCurrentThread<T>(Func<T> callback)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                VerifyAccess();
                completion.SetResult(callback());
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }

            return completion.Task;
        }

        private void VerifyAccess()
        {
            if (_disposed || _marshalControl.IsDisposed)
            {
                throw new ObjectDisposedException("OutlookUiSynchronizationContext");
            }
            if (Thread.CurrentThread.ManagedThreadId != _threadId
                || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException("The operation must run on the Outlook STA UI thread.");
            }
        }
    }
}
