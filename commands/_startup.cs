using Autodesk.AutoCAD.Runtime;
using System;

[assembly: ExtensionApplication(typeof(AutoCADBallet.AutoCADBalletStartup))]

namespace AutoCADBallet
{
    /// <summary>
    /// AutoCAD Ballet extension application startup.
    /// This is the single entry point for all initialization.
    /// Only ONE ExtensionApplication attribute is allowed per assembly.
    /// </summary>
    public class AutoCADBalletStartup : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                // Initialize network server (auto-starts on each AutoCAD session)
                AutoCADBalletServer.InitializeServer();

                // Initialize view/layout switch logging
                SwitchViewLogging.InitializeLogging();
            }
            catch (System.Exception ex)
            {
                // Silently handle initialization errors to not interfere with AutoCAD startup
                System.Diagnostics.Debug.WriteLine($"AutoCAD Ballet initialization error: {ex.Message}");
            }
        }

        public void Terminate()
        {
            try
            {
                // Terminate network server
                AutoCADBalletServer.TerminateServer();

                // Terminate view/layout switch logging
                SwitchViewLogging.TerminateLogging();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoCAD Ballet termination error: {ex.Message}");
            }
        }
    }
}
