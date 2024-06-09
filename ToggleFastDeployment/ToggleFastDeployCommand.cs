using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Task = System.Threading.Tasks.Task;

namespace ToggleFastDeployment
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ToggleFastDeployCommand
    {
        public const int ToggleCheckIconCommandId = 0x0100;
        public const int ToggleCrossIconCommandId = 0x0101;
        public const int ToggleButtonCommandId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("22c0c606-14b1-46d9-bb97-604e7893465c");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        readonly AsyncPackage package;

        readonly OleMenuCommand checkMenuItem;
        readonly OleMenuCommand crossMenuItem;
        readonly OleMenuCommand buttonMenuItem;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleFastDeployCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ToggleFastDeployCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            //var checkCommandId = new CommandID(CommandSet, ToggleCheckIconCommandId);
            //this.checkMenuItem = new OleMenuCommand(this.Execute, checkCommandId);
            //this.checkMenuItem.BeforeQueryStatus += UpdateButtonStatus;
            //commandService.AddCommand(checkMenuItem);

            //var crossCommandID = new CommandID(CommandSet, ToggleCrossIconCommandId);
            //this.crossMenuItem = new OleMenuCommand(this.Execute, crossCommandID);
            //this.crossMenuItem.BeforeQueryStatus += UpdateButtonStatus;
            //commandService.AddCommand(crossMenuItem);

            var menuCommandID = new CommandID(CommandSet, ToggleButtonCommandId);
            this.buttonMenuItem = new OleMenuCommand(this.Execute, menuCommandID);
            this.buttonMenuItem.BeforeQueryStatus += UpdateButtonStatus;
            commandService.AddCommand(buttonMenuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ToggleFastDeployCommand Instance
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
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ToggleFastDeployCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ToggleFastDeployCommand(package, commandService);
        }

        async void UpdateButtonStatus(object sender, EventArgs e)
        {
            var dte = await package.GetServiceAsync(typeof(DTE)).ConfigureAwait(false) as DTE2;
            var sb = (SolutionBuild2)dte.Solution.SolutionBuild;
            var startupProject = ((object[])sb.StartupProjects).FirstOrDefault().ToString();
            var project = dte.Solution.Item(startupProject);
            var projectFilePath = project.FullName;

            var projectDoc = XDocument.Load(projectFilePath);
            var useMauiElement = projectDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "UseMaui");
            if(useMauiElement == null || !bool.Parse(useMauiElement.Value))
            {
                buttonMenuItem.Enabled = false;
                return;
            }

            var embedElement = GetEmbedElement(projectDoc);
            var newText = embedElement is null ? "Fast Deploy (not found)" : $"Fast Deploy ({bool.Parse(embedElement.Value)})";
            buttonMenuItem.Text = newText;
        }

        async void Execute(object sender, EventArgs e)
        {
            var dte = await package.GetServiceAsync(typeof(DTE)).ConfigureAwait(false) as DTE2;
            var startupProject = dte.Solution.Projects.Item(1);
            var projectFilePath = startupProject.FullName;

            var projectDoc = XDocument.Load(projectFilePath);
            var embedElement = GetEmbedElement(projectDoc);
            if(embedElement != null)
            {
                bool currentValue = bool.Parse(embedElement.Value);
                embedElement.Value = (!currentValue).ToString();
                projectDoc.Save(projectFilePath);
            }
        }

        XElement GetEmbedElement(XDocument doc) => doc.Descendants().FirstOrDefault(x => x.Name.LocalName == "EmbedAssembliesIntoApk");
    }
}
