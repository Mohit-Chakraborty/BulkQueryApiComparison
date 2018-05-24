using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Microsoft.Win32;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject1
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class VSPackage : AsyncPackage, IVsSolutionEvents
    {
        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "2bd52308-b877-447a-a067-6252432ee5a9";

        private uint solutionEventsCookie;

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage"/> class.
        /// </summary>
        public VSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var solution = await this.GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            solution.AdviseSolutionEvents(this, out this.solutionEventsCookie);

            var projects = await GetProjectsAsync(solution, 10);

            // Switch to a background thread.
            // Run code that does not need the UI thread.
            await TaskScheduler.Default;

            await RunTestAsync(projects[0]);
        }

        #endregion

        private async Task<IVsHierarchy[]> GetProjectsAsync(IVsSolution solution, int numberOfProjects)
        {
            var projects = new IVsHierarchy[numberOfProjects];

            solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_ALLINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies);

            IVsHierarchy[] hierarchies = new IVsHierarchy[1];
            uint hierarchiesRetrieved = 0;
            int i = 0;

            while (enumHierarchies.Next(1, hierarchies, out hierarchiesRetrieved) == 0)
            {
                if (i >= numberOfProjects)
                {
                    break;
                }

                hierarchies[0].GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out object name);
                await this.WriteToOutputWindowAsync("Project name: " + name + Environment.NewLine);
                await this.WriteToOutputWindowAsync(Environment.NewLine);

                projects[i++] = hierarchies[0];
            }

            return projects;
        }

        private async Task RunTestAsync(IVsHierarchy hierarchy)
        {
            Stopwatch sw = new Stopwatch();

            // Old API optimistic scenario : We know upfront all properties we ever need to read for an item
            sw.Start();
            var values = await IVsHierarchyCalls.GetIVsHierarchyPropertiesAsync(hierarchy, this.JoinableTaskFactory);
            sw.Stop();
            await this.WriteToOutputWindowAsync("\tIVsHierarchyCalls (optimistic) time:   " + sw.ElapsedTicks + Environment.NewLine);

            // Old API pessimistic scenario : We make individual calls to read properties of an item
            sw.Restart();
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_Name, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_CanBuildFromMemory, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_Caption, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_DefaultEnableBuildProjectCfg, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_DefaultNamespace, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_DesignerFunctionVisibility, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_EditLabel, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_Expandable, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_Expanded, this.JoinableTaskFactory);
            await IVsHierarchyCalls.GetIVsHierarchyPropertyAsync(hierarchy, (int)__VSHPROPID.VSHPROPID_ExtObject, this.JoinableTaskFactory);
            sw.Stop();
            await this.WriteToOutputWindowAsync("\tIVsHierarchyCalls (pessimistic) time:  " + sw.ElapsedTicks + Environment.NewLine);

            // New API scenario : We collect properties we need to read and make one call via the COM layer
            sw.Restart();
            values = await IVsHierarchy2Calls.GetIVsHierarchyPropertiesAsync(hierarchy, this.JoinableTaskFactory);
            sw.Stop();
            await this.WriteToOutputWindowAsync("\tIVsHierarchy2Calls time:               " + sw.ElapsedTicks + Environment.NewLine);
        }

        private async System.Threading.Tasks.Task WriteToOutputWindowAsync(string output)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            var outputWindow = await this.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;

            Guid paneGuid = VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            int hr = outputWindow.GetPane(ref paneGuid, out IVsOutputWindowPane pane);

            if (ErrorHandler.Failed(hr) || (pane == null))
            {
                if (ErrorHandler.Succeeded(outputWindow.CreatePane(ref paneGuid, "General", fInitVisible: 1, fClearWithSolution: 1)))
                {
                    hr = outputWindow.GetPane(ref paneGuid, out pane);
                }
            }

            if (ErrorHandler.Succeeded(hr))
            {
                pane?.Activate();
                pane?.OutputString(output);
            }

            await TaskScheduler.Default;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            JoinableTaskFactory.Run(async delegate {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync();

                var solution = await this.GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
                //solution.AdviseSolutionEvents(this, out this.solutionEventsCookie);

                var projects = await GetProjectsAsync(solution, 10);

                // Switch to a background thread.
                // Run code that does not need the UI thread.
                await TaskScheduler.Default;

                await RunTestAsync(projects[0]);
            });

            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            //JoinableTaskFactory.Run(async delegate {
            //    await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            //    var solution = await this.GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            //    solution.UnadviseSolutionEvents(this.solutionEventsCookie);
            //});

            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }
    }
}
