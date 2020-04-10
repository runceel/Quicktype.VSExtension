using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace Quicktype
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class QuicktypeCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("af9e10ab-4080-4727-a2e4-f2fc2559365b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private static readonly string[] LanguageNames = { "c++", "cpp", "cplusplus", "cs", "csharp", "elm", "go", "golang", "java", "objc", "objective-c", "objectivec", "swift", "typescript", "ts", "tsx" };

        /// <summary>
        /// Initializes a new instance of the <see cref="QuicktypeCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private QuicktypeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static QuicktypeCommand Instance
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
            // Switch to the main thread - the call to AddCommand in QuicktypeCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new QuicktypeCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var jsonText = Clipboard.GetText().Trim();
            if (string.IsNullOrWhiteSpace(jsonText))
            {
                ShowMessage("Cannot paste - the clipboard is empty");
                return;
            }

            var dte = (DTE)Package.GetGlobalService(typeof(DTE));
            if (!(dte.ActiveDocument.Object() is TextDocument doc))
            {
                ShowMessage("Cannot paste - please open a document");
                return;
            }

            var language = doc.Language.ToLower();
            if (!LanguageNames.Contains(language))
            {
                ShowMessage($"Language \"{language}\" not supported");
                return;
            }

            var topLevelName = Path.GetFileNameWithoutExtension(dte.ActiveDocument.FullName);

            var jsonFileName = Path.GetTempFileName();
            File.WriteAllText(jsonFileName, jsonText);

            try
            {
                var p = new System.Diagnostics.Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.FileName = "quicktype.cmd";
                p.StartInfo.Arguments = "--telemetry disable --lang \"" + language + "\" --top-level \"" + topLevelName + "\" \"" + jsonFileName + "\"";
                p.Start();
                var output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode != 0)
                {
                    var error = p.StandardError.ReadToEnd();
                    ShowMessage($"quicktype could not process you JSON:\n\n{error}");
                    return;
                }

                doc.Selection.Insert(output);
            }
            catch (Win32Exception)
            {
                ShowMessage("Cannnot paste - Cannot find quicktype.cmd, please install quicktype usin `npm install -g quicktype` command");
                return;
            }
        }

        private void ShowMessage(string message) => VsShellUtilities.ShowMessageBox(
            this.package,
            message,
            "Quicktype",
            OLEMSGICON.OLEMSGICON_INFO,
            OLEMSGBUTTON.OLEMSGBUTTON_OK,
            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
    }
}
