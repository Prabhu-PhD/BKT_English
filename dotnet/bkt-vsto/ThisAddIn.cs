using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;

namespace BktVsto
{
    /// <summary>
    /// Thin VSTO host for the BKT (Business Kasper Toolbox) engine.
    ///
    /// BKT ships as a legacy managed-COM add-in (mscoree shim), which recent
    /// Office builds do not load. VSTO add-ins, however, do load. This shim is a
    /// minimal VSTO add-in whose only job is to load the existing BKT.AddIn
    /// engine, drive its connection lifecycle, and expose its ribbon.
    ///
    /// BKT.dll is loaded by reflection from the translated install's bin folder
    /// (rather than referenced at build time) for two reasons:
    ///   1. BKT.dll and its dependencies are not strong-named, so VSTO refuses to
    ///      list them as ClickOnce manifest prerequisites (error MSB3188).
    ///   2. BKT locates its config.txt relative to BKT.dll's own location, so
    ///      loading it from the install bin makes it pick up the translated
    ///      config.txt + features automatically.
    /// </summary>
    public partial class ThisAddIn
    {
        // Translated BKT install (created by the official setup + apply.py).
        private const string BktBin = @"C:\Users\prabh\AppData\Local\bkt-toolbox\bin";

        private object _bkt; // BKT.AddIn instance (implements Office.IRibbonExtensibility)

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Resolve BKT's dependencies (IronPython, Fluent, etc.) from the install bin.
            AppDomain.CurrentDomain.AssemblyResolve += ResolveFromBktBin;

            Assembly bktAsm = Assembly.LoadFrom(Path.Combine(BktBin, "BKT.dll"));
            Type addinType = bktAsm.GetType("BKT.AddIn");
            _bkt = Activator.CreateInstance(addinType);

            // Self-sufficient connect routine; takes the live PowerPoint Application.
            addinType.GetMethod("OnConnection2").Invoke(_bkt, new object[] { this.Application });
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                MethodInfo onDisconnect = _bkt.GetType().GetMethod("OnDisconnection");
                // First parameter is Extensibility.ext_DisconnectMode; ext_dm_HostShutdown == 0.
                Type disconnectMode = onDisconnect.GetParameters()[0].ParameterType;
                object hostShutdown = Enum.ToObject(disconnectMode, 0);
                object[] args = new object[] { hostShutdown, null };
                onDisconnect.Invoke(_bkt, args);
            }
            catch
            {
                // best-effort cleanup; never block Office shutdown
            }
        }

        /// <summary>
        /// Hand BKT's own IRibbonExtensibility implementation to Office. BKT
        /// generates its CustomUI XML dynamically and routes every ribbon
        /// callback to IronPython, exactly as in the COM-hosted case.
        /// </summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return (IRibbonExtensibility)_bkt;
        }

        private Assembly ResolveFromBktBin(object sender, ResolveEventArgs args)
        {
            try
            {
                string dll = new AssemblyName(args.Name).Name + ".dll";
                string path = Path.Combine(BktBin, dll);
                return File.Exists(path) ? Assembly.LoadFrom(path) : null;
            }
            catch
            {
                return null;
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
