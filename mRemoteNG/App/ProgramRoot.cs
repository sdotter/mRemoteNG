using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.Versioning;
using System.Threading;
using System.Windows.Forms;
using mRemoteNG.Config.Settings;
using mRemoteNG.UI.Forms;

namespace mRemoteNG.App
{
    [SupportedOSPlatform("windows")]
    public static class ProgramRoot
    {
        private static Mutex _mutex;
        private static FrmSplashScreenNew _frmSplashScreen = null;
        private static string customResourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Languages");

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            Trace.WriteLine("!!!!!!=============== TEST ==================!!!!!!!!!!!!!");
            // Forcing to load System.Configuration.ConfigurationManager before any other assembly to be able to check settings 
            try
            {
                string assemblyFile = "System.Configuration.ConfigurationManager" + ".dll";
                string assemblyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assemblies", assemblyFile);


                if (File.Exists(assemblyPath))
                {
                    Assembly.LoadFrom(assemblyPath);
                }
            }
            catch (FileNotFoundException ex)
            {
               Trace.WriteLine("Error occured: " + ex.Message);
            }

            //Subscribe to AssemblyResolve event
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            //Check if local settings DB exist or accessible
            CheckLockalDB();

            Lazy<bool> singleInstanceOption = new Lazy<bool>(() => Properties.OptionsStartupExitPage.Default.SingleInstance);

            if (singleInstanceOption.Value)
            {
                StartApplicationAsSingleInstance();
            }
            else
            {
                StartApplication();
            }
        }

        private static void CheckLockalDB()
        {
            LocalSettingsDBManager settingsManager = new LocalSettingsDBManager(dbPath: "mRemoteNG.appSettings", useEncryption: false, schemaFilePath: "");
        }
        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs resolveArgs)
        {
            string assemblyName = new AssemblyName(resolveArgs.Name).Name.Replace(".resources", string.Empty);
            string assemblyFile = assemblyName + ".dll";
            string assemblyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assemblies", assemblyFile);


            if (File.Exists(assemblyPath))
            {
                return Assembly.LoadFrom(assemblyPath);
            }
            return null;
        }
        
        private static void StartApplication()
        {
            CatchAllUnhandledExceptions();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            _frmSplashScreen = FrmSplashScreenNew.GetInstance();

            Screen currentMonitor = Screen.FromHandle(new System.Windows.Interop.WindowInteropHelper(_frmSplashScreen).Handle);

            System.Windows.PresentationSource source = System.Windows.PresentationSource.FromVisual(_frmSplashScreen);
            double dpiScaling = (source != null && source.CompositionTarget != null ? source.CompositionTarget.TransformFromDevice.M11 : 1);

            Rectangle workArea = currentMonitor.WorkingArea;
            var workAreaWidth = (int)Math.Floor(workArea.Width * dpiScaling);
            var workAreaHeight = (int)Math.Floor(workArea.Height * dpiScaling);

            _frmSplashScreen.Left = (((workAreaWidth - (_frmSplashScreen.Width * dpiScaling)) / 2) + (workArea.Left * dpiScaling));
            _frmSplashScreen.Top = (((workAreaHeight - (_frmSplashScreen.Height * dpiScaling)) / 2) + (workArea.Top * dpiScaling));
            _frmSplashScreen.ShowInTaskbar = false;
            _frmSplashScreen.Show();

            Application.Run(FrmMain.Default);
        }

        public static void CloseSingletonInstanceMutex()
        {
            _mutex?.Close();
        }

        private static void StartApplicationAsSingleInstance()
        {
            const string mutexID = "mRemoteNG_SingleInstanceMutex";
            _mutex = new Mutex(false, mutexID, out bool newInstanceCreated);
            if (!newInstanceCreated)
            {
                SwitchToCurrentInstance();
                return;
            }

            StartApplication();
            GC.KeepAlive(_mutex);
        }

        private static void SwitchToCurrentInstance()
        {
            IntPtr singletonInstanceWindowHandle = GetRunningSingletonInstanceWindowHandle();
            if (singletonInstanceWindowHandle == IntPtr.Zero) return;
            if (NativeMethods.IsIconic(singletonInstanceWindowHandle) != 0)
                _ = NativeMethods.ShowWindow(singletonInstanceWindowHandle, (int)NativeMethods.SW_RESTORE);
            NativeMethods.SetForegroundWindow(singletonInstanceWindowHandle);
        }

        private static IntPtr GetRunningSingletonInstanceWindowHandle()
        {
            IntPtr windowHandle = IntPtr.Zero;
            Process currentProcess = Process.GetCurrentProcess();
            foreach (Process enumeratedProcess in Process.GetProcessesByName(currentProcess.ProcessName))
            {
                if (enumeratedProcess.Id != currentProcess.Id &&
                    enumeratedProcess.MainModule.FileName == currentProcess.MainModule.FileName &&
                    enumeratedProcess.MainWindowHandle != IntPtr.Zero)
                    windowHandle = enumeratedProcess.MainWindowHandle;
            }

            return windowHandle;
        }

        private static void CatchAllUnhandledExceptions()
        {
            System.Windows.Forms.Application.ThreadException += ApplicationOnThreadException;
            System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += CurrentDomainOnUnhandledException;
        }

        private static void ApplicationOnThreadException(object sender, ThreadExceptionEventArgs e)
        {
            // if (PresentationSource.FromVisual(FrmSplashScreenNew))
            FrmSplashScreenNew.GetInstance().Close();

            if (FrmMain.Default.IsDisposed) return;

            FrmUnhandledException window = new(e.Exception, false);
            window.ShowDialog(FrmMain.Default);
        }

        private static void CurrentDomainOnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            //TODO: Check if splash closed properly
            //if (!FrmSplashScreenNew.GetInstance().IsDisposed)
            //    FrmSplashScreenNew.GetInstance().Close();

            FrmUnhandledException window = new(e.ExceptionObject as Exception, e.IsTerminating);
            window.ShowDialog(FrmMain.Default);
        }
        
    }
}