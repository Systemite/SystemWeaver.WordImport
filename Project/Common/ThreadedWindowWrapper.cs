using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace SystemWeaver.WordImport.Common
{
    public class ThreadedWindowWrapper : IDisposable
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetActiveWindow();

        public interface IThreadedWindow
        {
            Window Window { get; }
            void SetStatus(string status);
            void SetTitle(string title);
            void SetCanClose(bool canClose);
            void SetProgress(string value);
        }

        private static readonly int MaxThreadWaitMilliseconds;

        private IThreadedWindow _launchedWindow;
        private ManualResetEventSlim _event;

        static ThreadedWindowWrapper()
        {
            MaxThreadWaitMilliseconds = 5000;
        }

        public ThreadedWindowWrapper()
        {
            _launchedWindow = null;
            _event = new ManualResetEventSlim(false);
        }

        public void LaunchThreadedWindow<T>(bool modal, bool cancelButton = false) where T : Window, IThreadedWindow, new()
        {
            if (_launchedWindow != null)
                throw new Exception("Threaded window already launched.");


            IntPtr ownerHandle = GetActiveWindow();

            Thread thread = new Thread(() =>
            {
                T w = new T();
                _launchedWindow = w;
                w.Window.ShowInTaskbar = true;
                SetOwnerWindow(w, ownerHandle);

                _event.Set();

                if (modal)
                {
                    // ShowDialog automatically starts the event queue for the new windows in the new thread.
                    // The window isn't modal though.
                    w.ShowDialog();
                }
                else
                {
                    w.Show();
                    w.Closed += (sender2, e2) => w.Dispatcher.InvokeShutdown();
                    System.Windows.Threading.Dispatcher.Run();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            _event.Wait(MaxThreadWaitMilliseconds);
            _event.Reset();
        }

        public void Close()
        {
            if (_launchedWindow == null)
                return;

            InvokeSetter<bool>(new SetterInfo<bool>("SetCanClose", true));

            if (_launchedWindow.Window.Dispatcher.CheckAccess())
                _launchedWindow.Window.Close();
            else
                _launchedWindow.Window.Dispatcher.Invoke(DispatcherPriority.Normal, new ThreadStart(Close));

            _launchedWindow = null;
        }

        public void SetStatus(string status)
        {
            InvokeSetter<string>(new SetterInfo<string>("SetStatus", status));
        }

        public void SetTitle(string title)
        {
            InvokeSetter<string>(new SetterInfo<string>("SetTitle", title));
        }
        public void SetProgress(string value)
        {
            InvokeSetter<string>(new SetterInfo<string>("SetProgress", value));
        }

        private void InvokeSetter<T>(object info)
        {
            SetterInfo<T> setterInfo = info as SetterInfo<T>;
            if (_launchedWindow == null || setterInfo == null || setterInfo.MethodName == null)
                return;

            var method = typeof(IThreadedWindow).GetMethod(setterInfo.MethodName);
            if (method == null)
                throw new Exception("Method not found: " + setterInfo.MethodName);

            var parameters = method.GetParameters();
            if (parameters.Length != 1)
                throw new Exception("Incorrect parameter count for method: " + parameters.Length + ". Method must take a single parameter.");

            if (parameters[0].ParameterType != typeof(T))
                throw new Exception("Wrong parameter type: " + typeof(T).Name + ". Expected: " + parameters[0].ParameterType.Name + ".");

            if (_launchedWindow.Window.Dispatcher.CheckAccess())
                method.Invoke(_launchedWindow, new object[] { setterInfo.Parameter });
            else
                _launchedWindow.Window.Dispatcher.Invoke(DispatcherPriority.Normal, new ParameterizedThreadStart(InvokeSetter<T>), setterInfo);
        }

        public void Dispose()
        {
            Close();
        }

        private const int GWL_HWNDPARENT = -8; // Owner --> not the parent

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);
        private static void SetOwnerWindow(Window owned, IntPtr intPtrOwner)
        {
            IntPtr windowHandleOwned = new WindowInteropHelper(owned).Handle;
            if (windowHandleOwned != IntPtr.Zero && intPtrOwner != IntPtr.Zero)
            {
                SetWindowLong(windowHandleOwned, GWL_HWNDPARENT, intPtrOwner.ToInt32());
            }
        }

        private class SetterInfo<T>
        {
            public SetterInfo(string methodName, T parameter) { MethodName = methodName; Parameter = parameter; }
            public string MethodName { get; set; }
            public T Parameter { get; set; }
        }
        public IThreadedWindow Window { get { return _launchedWindow; } }
    }
}
