using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Threading;
using Nav.Common.VSPackages.VSFormatFile.JSBeautifier;
using DefGuidList = Microsoft.VisualStudio.Editor.DefGuidList;
using IServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Task = System.Threading.Tasks.Task;

namespace Nav.Common.VSPackages.VSFormatFile
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}", PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuids.GuidVSFormatFilePkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Format File", "Settings", 0, 0, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public class VSFormatFilePackage : AsyncPackage
    {
        public DTE2 Dte { get; private set; }
        public OptionsPage OptionsPage { get; private set; }
        public OleMenuCommandService MenuCommandService { get; private set; }

        private RunningDocumentTable _runningDocumentTable;
        private ServiceProvider _serviceProvider;
        private ITextUndoHistoryRegistry _undoHistoryRegistry;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _runningDocumentTable = new RunningDocumentTable(this);
            OptionsPage = (OptionsPage)GetDialogPage(typeof(OptionsPage));

            Dte = await GetServiceAsync(typeof(SDTE)) as DTE2;

            _serviceProvider = new ServiceProvider((IServiceProvider)Dte);
            var componentModel = (IComponentModel)GetGlobalService(typeof(SComponentModel));

            _undoHistoryRegistry = componentModel.DefaultExportProvider.GetExportedValue<ITextUndoHistoryRegistry>();
            MenuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            new SolutionExplorerFormatContextMenu(this);

            new SolutionExplorerJSBeautifierContextMenu(this);
        }

        Document FindDocument(uint docCookie)
        {
            var documentInfo = _runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            return Dte.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentPath);
        }

        public void Format(uint docCookie)
        {
            var document = FindDocument(docCookie);
            Format(document);
        }

        public bool Format(Document document)
        {
            if (document == null || document.Type != "Text" || document.Language == null ||
                document.Language == "Plain Text")
                return false;

            var oldActiveDocument = Dte.ActiveDocument;

            try
            {
                document.Activate();

                var languageOptions = Dte.Properties["TextEditor", document.Language];
                var insertTabs = (bool)languageOptions.Item("InsertTabs").Value;
                var isFilterAllowed = OptionsPage.AllowDenyFilter.IsAllowed(document.Name);

                var vsTextView = GetIVsTextView(document.FullName);
                if (vsTextView == null)
                    return false;
                var wpfTextView = GetWpfTextView(vsTextView);
                if (wpfTextView == null)
                    return false;

                _undoHistoryRegistry.TryGetHistory(wpfTextView.TextBuffer, out var history);

                using (var undo = history?.CreateTransaction("Format File"))
                {
                    vsTextView.GetCaretPos(out var oldCaretLine, out var oldCaretColumn);
                    vsTextView.SetCaretPos(oldCaretLine, 0);

                    // Do TabToSpace before FormatDocument, since VS format may break the tab formatting.
                    if (OptionsPage.EnableTabToSpace && isFilterAllowed && !insertTabs)
                        TabToSpace(wpfTextView, document.TabSize);

                    if (OptionsPage.EnableRemoveAndSort && IsCsFile(document))
                    {
                        if (!OptionsPage.EnableSmartRemoveAndSort || !HasIfCompilerDirective(wpfTextView))
                            RemoveAndSort();
                    }

                    if (OptionsPage.EnableFormatDocument &&
                        OptionsPage.AllowDenyFormatDocumentFilter.IsAllowed(document.Name))
                        FormatDocument();

                    // Do TabToSpace again after FormatDocument, since VS2017 may stick to tab. Should remove this after VS2017 fix the bug.
                    if (OptionsPage.EnableTabToSpace && isFilterAllowed && !insertTabs && Dte.Version == "15.0" &&
                        document.Language == "C/C++")
                        TabToSpace(wpfTextView, document.TabSize);

                    if (OptionsPage.EnableUnifyLineBreak && isFilterAllowed)
                        UnifyLineBreak(wpfTextView);

                    if (OptionsPage.EnableUnifyEndOfFile && isFilterAllowed)
                        UnifyEndOfFile(wpfTextView);

                    if (OptionsPage.EnableForceUtf8WithoutBom &&
                        OptionsPage.AllowDenyForceUtf8WithoutBomFilter.IsAllowed(document.Name))
                        ForceUtf8WithoutBom(wpfTextView);

                    if (OptionsPage.EnableRemoveTrailingSpaces && isFilterAllowed &&
                        (Dte.Version == "11.0" || !OptionsPage.EnableFormatDocument))
                        RemoveTrailingSpaces(wpfTextView);

                    vsTextView.GetCaretPos(out var newCaretLine, out var newCaretColumn);
                    vsTextView.SetCaretPos(newCaretLine, oldCaretColumn);

                    undo?.Complete();
                }
            }
            finally
            {
                oldActiveDocument?.Activate();
            }

            return true;
        }

        static bool IsCsFile(Document document)
        {
            return document.FullName.EndsWith(".cs", StringComparison.OrdinalIgnoreCase);
        }

        static bool HasIfCompilerDirective(ITextView wpfTextView)
        {
            return wpfTextView.TextSnapshot.GetText().Contains("#if");
        }

        void RemoveAndSort()
        {
            try
            {
                Dte.ExecuteCommand("Edit.RemoveAndSort", string.Empty);
            }
            catch (COMException)
            {
            }
        }

        void FormatDocument()
        {
            try
            {
                Dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
            }
            catch (COMException)
            {
            }
        }

        void UnifyLineBreak(ITextView wpfTextView)
        {
            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                string defaultLineBreak;
                switch (OptionsPage.LineBreak)
                {
                    case OptionsPage.LineBreakStyle.Unix:
                        defaultLineBreak = "\n";
                        break;
                    case OptionsPage.LineBreakStyle.Windows:
                        defaultLineBreak = "\r\n";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                foreach (var line in snapshot.Lines)
                {
                    if (line.LineNumber == (snapshot.LineCount - 1)
                        || line.LineNumber == snapshot.LineCount)
                    {
                        edit.Delete(line.End.Position, line.LineBreakLength);
                    }

                    // if line break is defaultLineBreak or the line is the last => continue;
                    if (line.GetLineBreakText() == defaultLineBreak
                        || line.LineNumber == (snapshot.LineCount - 1)
                        || line.LineNumber == snapshot.LineCount
                    )
                    {
                        continue;
                    }

                    edit.Delete(line.End.Position, line.LineBreakLength);
                    edit.Insert(line.End.Position, defaultLineBreak);
                }

                edit.Apply();
            }
        }

        static void UnifyEndOfFile(ITextView textView)
        {
            var snapshot = textView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var notEmptyLineNumber = snapshot.LineCount - 1;
                while (notEmptyLineNumber >= 0 && string.IsNullOrEmpty(snapshot.GetLineFromLineNumber(notEmptyLineNumber).GetText().Trim()))
                {
                    notEmptyLineNumber--;
                }

                var hasModified = false;

                if (notEmptyLineNumber < snapshot.LineCount - 1)
                {
                    var startPosition = snapshot.GetLineFromLineNumber(notEmptyLineNumber).End.Position;

                    var endPosition = snapshot.GetLineFromLineNumber(snapshot.LineCount - 1).EndIncludingLineBreak.Position;

                    edit.Delete(startPosition, endPosition - startPosition);
                    hasModified = true;
                }

                if (hasModified)
                    edit.Apply();
            }
        }

        static void ForceUtf8WithoutBom(ITextView wpfTextView)
        {
            try
            {
                ITextDocument textDocument;
                wpfTextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument),
                    out textDocument);

                var encoding = new UTF8Encoding(false);

                if (textDocument.Encoding != encoding)
                {
                    textDocument.Encoding = encoding;
                }
            }
            catch (Exception)
            {
            }
        }

        static void RemoveTrailingSpaces(ITextView textView)
        {
            var snapshot = textView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var hasModified = false;

                for (var i = 0; i < snapshot.LineCount; i++)
                {
                    var line = snapshot.GetLineFromLineNumber(i);
                    var lineText = line.GetText();

                    var trimmedLength = lineText.TrimEnd().Length;
                    if (trimmedLength == lineText.Length)
                        continue;

                    var spaceLength = lineText.Length - trimmedLength;
                    var endPosition = line.End.Position;
                    edit.Delete(endPosition - spaceLength, spaceLength);
                    hasModified = true;
                }

                if (hasModified)
                    edit.Apply();
            }
        }

        class SpaceStringPool
        {
            readonly string[] _stringCache = new string[8];

            public string GetString(int spaceCount)
            {
                if (spaceCount <= 0)
                    throw new ArgumentOutOfRangeException();

                var index = spaceCount - 1;

                if (spaceCount > _stringCache.Length)
                    return new string(' ', spaceCount);
                if (_stringCache[index] == null)
                {
                    _stringCache[index] = new string(' ', spaceCount);
                    return _stringCache[index];
                }
                return _stringCache[index];
            }
        }

        readonly SpaceStringPool _spaceStringPool = new SpaceStringPool();

        void TabToSpace(ITextView wpfTextView, int tabSize)
        {
            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var hasModifed = false;

                foreach (var line in snapshot.Lines)
                {
                    var lineText = line.GetText();

                    if (!lineText.Contains('\t'))
                        continue;

                    var positionOffset = 0;

                    for (var i = 0; i < lineText.Length; i++)
                    {
                        var currentChar = lineText[i];
                        if (currentChar == '\t')
                        {
                            var absTabPosition = line.Start.Position + i;
                            edit.Delete(absTabPosition, 1);
                            var spaceCount = tabSize - (i + positionOffset) % tabSize;
                            edit.Insert(absTabPosition, _spaceStringPool.GetString(spaceCount));
                            positionOffset += spaceCount - 1;
                            hasModifed = true;
                        }
                        else if (IsCjkCharacter(currentChar))
                            positionOffset++;
                    }
                }

                if (hasModifed)
                    edit.Apply();
            }
        }

        readonly Regex _cjkRegex = new Regex(
            @"\p{IsHangulJamo}|" +
            @"\p{IsCJKRadicalsSupplement}|" +
            @"\p{IsCJKSymbolsandPunctuation}|" +
            @"\p{IsEnclosedCJKLettersandMonths}|" +
            @"\p{IsCJKCompatibility}|" +
            @"\p{IsCJKUnifiedIdeographsExtensionA}|" +
            @"\p{IsCJKUnifiedIdeographs}|" +
            @"\p{IsHangulSyllables}|" +
            @"\p{IsCJKCompatibilityForms}|" +
            @"\p{IsHalfwidthandFullwidthForms}");

        bool IsCjkCharacter(char character)
        {
            return _cjkRegex.IsMatch(character.ToString());
        }

        static IWpfTextView GetWpfTextView(IVsTextView vTextView)
        {
            IWpfTextView view = null;
            var userData = (IVsUserData)vTextView;

            if (userData != null)
            {
                var guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var holder);
                var viewHost = (IWpfTextViewHost)holder;
                view = viewHost.TextView;
            }

            return view;
        }

        IVsTextView GetIVsTextView(string filePath)
        {
            return VsShellUtilities.IsDocumentOpen(_serviceProvider, filePath, Guid.Empty, out var uiHierarchy,
                out var itemId, out var windowFrame)
                ? VsShellUtilities.GetTextView(windowFrame)
                : null;
        }

        IVsOutputWindowPane _outputWindowPane;

        public void OutputString(string message)
        {
            Guid _outputWindowPaneGuid = PackageGuids.GuidVSFormatFileOutputWindowPane;

            if (_outputWindowPane == null)
            {
                var outWindow = (IVsOutputWindow)GetGlobalService(typeof(SVsOutputWindow));
                outWindow.CreatePane(ref _outputWindowPaneGuid, "VSFormatFile", 1, 1);
                outWindow.GetPane(ref _outputWindowPaneGuid, out _outputWindowPane);
            }

            _outputWindowPane.OutputString(message + Environment.NewLine);
            _outputWindowPane.Activate(); // Brings this pane into view
        }

        public void JSBeautifier(uint docCookie)
        {
            var document = FindDocument(docCookie);
            JSBeautifier(document);
        }

        public bool JSBeautifier(Document document)
        {
            if (document == null
                || document.Type != "Text"
                || document.Language == null
                || document.Language == "Plain Text"
                )
            {
                return false;
            }

            var oldActiveDocument = Dte.ActiveDocument;

            try
            {
                document.Activate();

                var languageOptions = Dte.Properties["TextEditor", document.Language];
                var insertTabs = (bool)languageOptions.Item("InsertTabs").Value;
                var isFilterAllowed = OptionsPage.AllowDenyFilter.IsAllowed(document.Name);

                var vsTextView = GetIVsTextView(document.FullName);
                if (vsTextView == null)
                    return false;

                var wpfTextView = GetWpfTextView(vsTextView);
                if (wpfTextView == null)
                    return false;

                _undoHistoryRegistry.TryGetHistory(wpfTextView.TextBuffer, out var history);

                using (var undo = history?.CreateTransaction("JSBeautify"))
                {
                    var res = vsTextView.GetCaretPos(out var oldCaretLine, out var oldCaretColumn);
                    res = vsTextView.SetCaretPos(oldCaretLine, 0);

                    var snapshot = wpfTextView.TextSnapshot;

                    using (var edit = snapshot.TextBuffer.CreateEdit())
                    {
                        string text = snapshot.GetText();

                        var beautifier = new Beautifier();

                        beautifier.Opts.IndentSize = 4;
                        beautifier.Opts.IndentChar = ' ';

                        beautifier.Opts.PreserveNewlines = true;
                        beautifier.Opts.JslintHappy = true;

                        beautifier.Opts.KeepArrayIndentation = true;
                        beautifier.Opts.KeepFunctionIndentation = false;

                        beautifier.Opts.BraceStyle = BraceStyle.Collapse;
                        beautifier.Opts.BreakChainedMethods = false;

                        beautifier.Flags.IndentationLevel = 0;

                        text = beautifier.Beautify(text);

                        edit.Replace(0, snapshot.Length, text);

                        edit.Apply();
                    }

                    res = vsTextView.GetCaretPos(out var newCaretLine, out var newCaretColumn);
                    res = vsTextView.SetCaretPos(newCaretLine, oldCaretColumn);

                    undo?.Complete();
                }
            }
            finally
            {
                oldActiveDocument?.Activate();
            }

            return true;
        }
    }
}
