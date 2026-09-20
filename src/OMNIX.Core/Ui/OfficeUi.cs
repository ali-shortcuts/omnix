using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace OMNIX.Core.Ui
{
    // VSTO can enter WPF without a SynchronizationContext or Application.Current.
    // Install a context only while starting the operation; its awaits retain the Office dispatcher.
    internal static class OfficeUi
    {
        public static Task RunAsync(Dispatcher dispatcher, Func<Task> operation)
        {
            return dispatcher.InvokeAsync(() =>
            {
                var previous = SynchronizationContext.Current;
                SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(dispatcher));
                try { return operation(); }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }).Task.Unwrap();
        }
    }
}
