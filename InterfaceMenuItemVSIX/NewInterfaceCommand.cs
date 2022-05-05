using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Composition;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

namespace InterfaceMenuItemVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class NewInterfaceCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0001;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("e8ad8e3e-0865-46d7-a1a8-2e4626c1bc0f");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        [Import]
        internal SVsServiceProvider SVsServiceProvider { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="NewInterfaceCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private NewInterfaceCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this._package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.CreateNewInterfaceFile, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static NewInterfaceCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this._package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in NewInterfaceCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new NewInterfaceCommand(package, commandService);
        }

        /// <summary>
        /// Create new interface file at context location
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void CreateNewInterfaceFile(object sender, EventArgs e)
        {
            // Must be on main UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            // Create new interface file
            try
            {
                // Get DTE2 service
                var dte2 = this._package.GetService<SDTE, DTE2>() as DTE2;
                if (dte2 == null)
                {
                    Debug.WriteLine("Could not obtain VS SDTE service.", ">> DEBUG-ME");
                    return;
                }

                // Get
                var addProjItemDlg = this._package.GetService<SVsAddProjectItemDlg, IVsAddProjectItemDlg>() as IVsAddProjectItemDlg;
                if (addProjItemDlg == null)
                {
                    Debug.WriteLine("Could not obtain VS SVsAddProjectItemDlg service.", ">> DEBUG-ME");
                    return;
                }

                // Init used in AddProjectItemDlg(...) call; keep same name
                UInt32 itemIdLoc = 0;
                Guid rguidProject = Guid.Empty;
                IVsProject pProject = null;
                UInt32 grfAddFlags = (uint)(__VSADDITEMFLAGS.VSADDITEM_AddNewItems |
                    __VSADDITEMFLAGS.VSADDITEM_SuggestTemplateName |
                    __VSADDITEMFLAGS.VSADDITEM_AllowHiddenTreeView);
                String lpszExpand = "Code";
                String lpszSelect = "Interface";
                String pbstrLocation = null;
                String pbstrFilter = "";
                Int32 pfDontShowAgain;

                // Get item ID location from VS hierarchy selection
                IVsHierarchy hierarchy = GetCurrentVsHierarchySelection(out itemIdLoc);
                if (hierarchy == null)
                {
                    Debug.WriteLine("Could not obtain VS selected item data.", ">> DEBUG-ME");
                    return;
                }

                // Get project
                Project project = VsHierarchyToProject(hierarchy);
                if (project == null)
                {
                    Debug.WriteLine("Could not obtain project data", ">> DEBUG-ME");
                    return;
                }

                // Get project guid
                rguidProject = new Guid(project.Kind);

                // Get converted VS project from project
                pProject = ProjectToVsProject(project);
                if (pProject == null)
                {
                    Debug.WriteLine("Could not convert project data to VS project.", ">> DEBUG-ME");
                    return;
                }

                // Get location
                hierarchy.GetCanonicalName(itemIdLoc, out pbstrLocation);

                // Open project item dialog
                var rtn = addProjItemDlg.AddProjectItemDlg(
                    itemIdLoc,
                    ref rguidProject,
                    pProject,
                    grfAddFlags,
                    lpszExpand,
                    lpszSelect,
                    ref pbstrLocation,
                    ref pbstrFilter,
                    out pfDontShowAgain);
                Debug.WriteLine($"Don't show again: {pfDontShowAgain}", ">> DEBUG-ME");

                // Is it OK?
                if (rtn != VSConstants.S_OK)
                {
                    Debug.WriteLine("Add project item dialog closed or canceled.", ">> DEBUG-ME");
                    return;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);

                VsShellUtilities.ShowMessageBox(
                    this._package,
                    $"Unable to create new interface file: {ex.Message}",
                    "Error",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        /// <summary>
        /// Get Solution Explorer's selected item directory 
        /// </summary>
        /// <param name="dte2">DTE2 service object</param>
        /// <param name="dir">Directory of selected item; out</param>
        /// <returns>True if directory of selected item is null/empty/white-space.</returns>
        private bool GetSolutionExplorerSelectedProjectItemDir(DTE2 dte2, out string dir)
        {
            // Must be on main UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            // The path
            dir = "";

            // Process
            if (dte2 != null &&
                dte2.SelectedItems.Count == 1)
            {
                // Get selected item
                var selectedItem = dte2.SelectedItems.Item(1);
                if (selectedItem != null)
                {
                    Debug.WriteLine(string.Format("Item name: {0}", selectedItem.Name), ">> DEBUG-ME");

                    // Is a directory selected
                    var projectItem = selectedItem.ProjectItem; // If selectedItem is a project, then .ProjectItem will be null,
                                                                // and .Project will not be null
                    if (projectItem != null &&
                        projectItem.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFolder)
                    {
                        try
                        {
                            var fullPathProp = projectItem.Properties.Item("FullPath");
                            dir = fullPathProp.Value.ToString();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message, ">> DEBUG-ME");
                        }
#if DEBUG
                        for (var i = 1; i <= projectItem.Properties.Count; i++)
                        {
                            try
                            {
                                var prop = projectItem.Properties.Item(i);
                                Debug.WriteLine(string.Format("Project Item: Property: {0} = {1}", prop.Name, prop.Value), ">> DEBUG-ME");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message, ">> DEBUG-ME");
                            }
                        }
#endif
                    }

                    // Is a project selected
                    if (string.IsNullOrWhiteSpace(dir))
                    {
                        var project = selectedItem.Project; // If selectedItem is a project, then .ProjectItem will be null,
                                                            // and .Project will not be null
                        if (project != null)
                        {
                            dir = Path.GetDirectoryName(project.FullName);
                        }
                    }

                    // Do we have a directory yet? USE IT!
                    if (!string.IsNullOrWhiteSpace(dir))
                    {
                        Debug.WriteLine(string.Format("Directory: {0}", dir), ">> DEBUG-ME");
                    }
                    else
                    {
                        Debug.WriteLine("Could not obtain directory path for selected item.", ">> DEBUG-ME");
                    }
                }
                else
                {
                    Debug.WriteLine("Selected item is null... ?", ">> DEBUG-ME");
                }
            }
            else
            {
                Debug.WriteLine("DTE2 ref is null or nothing is selected in solution explorer.", ">> DEBUG-ME");
            }

            return !string.IsNullOrWhiteSpace(dir);
        }

        /// <summary>
        /// Get current VS hierarchy selection
        /// </summary>
        /// <param name="projectItemId">Project item ID of selected item; out</param>
        /// <returns><see cref="IVsHierarchy"/> instance of selection</returns>
        public IVsHierarchy GetCurrentVsHierarchySelection(out uint projectItemId)
        {
            // Must be on main UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            // Set output var
            projectItemId = 0;

            // Declare vars
            IntPtr intPtrHierarchy;
            IntPtr intPtrSelectionContainer;
            IVsMultiItemSelect vsMultiItemSelect;

            // Get shell monitor selection service; do cast here, instead of 'as', to produce error if service is not available 
            var vsMonitorSelection = (IVsMonitorSelection)Package.GetGlobalService(typeof(SVsShellMonitorSelection));

            // Get current selection from service
            vsMonitorSelection.GetCurrentSelection(out intPtrHierarchy, out projectItemId, out vsMultiItemSelect, out intPtrSelectionContainer);

            // Convert multi item select to hierarchy object
            var vsHierarchy = Marshal.GetTypedObjectForIUnknown(intPtrHierarchy, typeof(IVsHierarchy)) as IVsHierarchy;

            // Return hierarchy object
            return vsHierarchy;
        }

        /// <summary>
        /// Converts instance of <see cref="IVsHierarchy"/> to instance of <see cref="Project"/>
        /// </summary>
        /// <param name="vsHierarchy">Instance of <see cref="IVsHierarchy"/> to convert</param>
        /// <returns>Converted instance of <see cref="Project"/></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        private Project VsHierarchyToProject(IVsHierarchy vsHierarchy)
        {
            if (vsHierarchy is null)
            {
                throw new ArgumentNullException(nameof(vsHierarchy));
            }

            // Must be on main UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            // Convert IVsHierarchy object to Project object
            object project = null;
            if (vsHierarchy.GetProperty(0xfffffffe, (int)__VSHPROPID.VSHPROPID_ExtObject, out project) == VSConstants.S_OK)
            {
                return project as Project;
            }

            throw new ArgumentException("Hierarchy is not a project.");
        }

        /// <summary>
        /// Converts instance of <see cref="Project"/> to instance of <see cref="IVsProject"/>
        /// </summary>
        /// <param name="project">INstance of <see cref="Project"/> to convert</param>
        /// <returns>Converted instance of <see cref="IVsProject"/></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        private IVsProject ProjectToVsProject(Project project)
        {
            if (project is null)
            {
                throw new ArgumentNullException(nameof(project));
            }

            // Must be on main UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            // Vet SVsSolution service
            var vsSolution = this._package.GetService<SVsSolution, IVsSolution>() as IVsSolution;
            if (vsSolution != null)
            {
                // Convert Project into IVsProject
                IVsProject vsProject = null;
                IVsHierarchy vsHierarchy;
                vsSolution.GetProjectOfUniqueName(project.UniqueName, out vsHierarchy);
                vsProject = vsHierarchy as IVsProject;
                if (vsProject != null)
                {
                    return vsProject;
                }
            }

            throw new ArgumentException("Project is not a VS project.");
        }
    }
}
