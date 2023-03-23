using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Design;
using System.Linq;
using BJVSExtension.Utilities;
using Microsoft.VisualStudio.Text.Editor;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Package;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Projection;

namespace BJVSExtension
{
    class ReIndexCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("{3a251d82-e8ce-442f-9e42-5285653a5e8a}");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// The top-level object in the Visual Studio automation object model
        /// </summary>
        private static DTE2 dte;
        
        
        List<CodeRegin> Regions = new List<CodeRegin>();

        private Object[] _tokensFull;
        private Object[] _tokens;
        
        private EnvDTE.Properties properties;

        /// <summary>
        /// Contains a list of all comment tokens as objects on this instance of Visual Studio
        /// </summary>
        public Object[] TokensFull
        {
            get { return _tokensFull; }
            set
            {
                _tokensFull = value;
            }
        }

        /// <summary>
        /// Contains a slightly adjust readable version of the tokens, removing the :2 at the end
        /// Also adding "ALL" as an option and removing the last option "UnresolvedMergeConflict
        /// </summary>
        public Object[] Tokens
        {
            get { return _tokens; }
            set
            {
                _tokens = value;
                //OnPropertyChanged("Tokens");
            }
        }
        
        private ReIndexCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ReIndexCommand Instance
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
            // Switch to the main thread - the call to AddCommand in AlignCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ReIndexCommand(package, commandService);

            // Get DTE object
            if (await package.GetServiceAsync(typeof(DTE)) is DTE2 dte2)
                dte = dte2;
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

            // TODO：测试代码
            // Get all the comment tokens from the application object
            GetCommentTokens();
            // Get the readable list of comment tokens
            GetReadableCommentTokens();
            
            // TODO: 测试从Task窗口获取Task
            TaskList theTasks = dte.ToolWindows.TaskList;
            TaskItems2 TLItems = (TaskItems2)theTasks.TaskItems;
            //IVsTaskList3 taskList;
            //ITaskList _taskList = serviceProvider.GetService(typeof(SVsTaskList)) as ITaskList;
            // VS2015开始 用户不能添加自己的Task The user task feature was removed in Visual Studio 2015.
            // https://docs.microsoft.com/en-us/visualstudio/ide/using-the-task-list?view=vs-2022
            //TLItems.Add2("TODOTask", "LIJUN", "Test task 1.", 2);


            IWpfTextView textView = TextViewHelper.GetActiveTextView();
            (int startLineNo, int endLineNo) = TextViewHelper.GetSelectedLineNumbers(textView);
            
            SelectedLines selectedLines = new SelectedLines(textView.TextSnapshot, startLineNo, endLineNo);

            if (selectedLines.Lines.Count() >= 2)
            {   // 手工选择范围
                int tabSize = dte.ActiveDocument.TabSize;
                CommentReIndexer commentAligner = new CommentReIndexer(selectedLines.Lines, tabSize, selectedLines.LineEnding);
                string newText = commentAligner.GetText();

                try
                {
                    dte.UndoContext.Open("ReIndex comments");
                    TextViewHelper.ReplaceText(textView, selectedLines.StartPosition, selectedLines.Length, newText);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }
                finally
                {
                    dte.UndoContext.Close();
                }
            }
            else
            {   // 自动分析光标所在函数范围
                
                ITextBuffer buffer = textView.TextBuffer;
                var snapshot = buffer.CurrentSnapshot;

                IComponentModel componentModel = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
                if (componentModel == null)
                    return;

                IClassifier classifier = componentModel.GetService<IClassifierAggregatorService>().GetClassifier(buffer);
                ITextEditorFactoryService textEditorFactory = componentModel.GetService<ITextEditorFactoryService>();
                IProjectionBufferFactoryService projectionBufferFactory = componentModel.GetService<IProjectionBufferFactoryService>();

                RegionFinder finder = new RegionFinder(snapshot, classifier, textEditorFactory, projectionBufferFactory);

                Regions = finder.FindAll();

                foreach (var region in Regions)
                {
                    if(region.RegionType == CodeRegionType.Method && region.Complete)
                    {
                        if(startLineNo >= region.StartLine.LineNumber && startLineNo < region.EndLine.LineNumber)
                        {
                            selectedLines = new SelectedLines(textView.TextSnapshot, region.StartLine.LineNumber, region.EndLine.LineNumber-1);
                           
                            int tabSize = dte.ActiveDocument.TabSize;
                            CommentReIndexer commentAligner = new CommentReIndexer(selectedLines.Lines, tabSize, selectedLines.LineEnding);
                            string newText = commentAligner.GetText();

                            try
                            {
                                dte.UndoContext.Open("ReIndex comments");
                                TextViewHelper.ReplaceText(textView, selectedLines.StartPosition, selectedLines.Length, newText);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.Write(ex);
                            }
                            finally
                            {
                                dte.UndoContext.Close();
                            }
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Read all the comment tokens from the application object
        /// </summary>
        public void GetCommentTokens()
        {
            // Represents the Task List Node under the Enviroment node.
            properties = dte.get_Properties("Environment", "TaskList");

            // Represents the items in the comment Token list and their priorities (1-3/low-high).
            EnvDTE.Property commentProperties = properties.Item("CommentTokens");
            TokensFull = (Object[])commentProperties.Value;
        }

        public void GetReadableCommentTokens()
        {
            // Create an array of strings based on the list of token objects
            Tokens = new String[TokensFull.Length + 1];

            // Add an item for ALL
            Tokens[0] = "ALL";

            for (int i = 1; i < Tokens.Length; i++)
            {
                // Split the name TODO:2 into two sections and just take the first TODO
                Tokens[i] = TokensFull[i - 1].ToString().Split(new char[] { ':' })[0];
            }
        }
    }
}
