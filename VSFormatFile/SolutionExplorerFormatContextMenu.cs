using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;

namespace Nav.Common.VSPackages.VSFormatFile
{
    class SolutionExplorerFormatContextMenu
    {
        readonly VSFormatFilePackage _package;

        public SolutionExplorerFormatContextMenu(VSFormatFilePackage package)
        {
            _package = package;

            var mcs = _package.MenuCommandService;

            var menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetCode, PackageIds.CmdIdVSFormatFileCode);
            var menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetCode, PackageIds.CmdIdVSFormatFileOpenDocuments);
            menuItem = new OleMenuCommand(VSFormatFileOpenedDocumentsEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            menuItem.BeforeQueryStatus += oleMenuItemDocuments_BeforeQueryStatus;
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetFile, PackageIds.CmdIdVSFormatFileFile);
            menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetFolder, PackageIds.CmdIdVSFormatFileFolder);
            menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetProject, PackageIds.CmdIdVSFormatFileProject);
            menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetSolution, PackageIds.CmdIdVSFormatFileSolution);
            menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetSolution, PackageIds.CmdIdVSFormatFileSolutionFolder);
            menuItem = new OleMenuCommand(VSFormatFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);
        }

        private void oleMenuItemDocuments_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null)
            {
                return;
            }

            var visible = false;

            if (_package.Dte != null
                && _package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeDocument
                && _package.Dte.ActiveWindow.Document != null
                && _package.Dte.Documents != null
                )
            {
                foreach (var doc in _package.Dte.Documents.OfType<EnvDTE.Document>().Where(d => d != _package.Dte.ActiveWindow.Document))
                {
                    if (doc.ActiveWindow == null
                      || doc.ActiveWindow.Type != EnvDTE.vsWindowType.vsWindowTypeDocument
                      || doc.ActiveWindow.Visible == false
                      )
                    {
                        continue;
                    }

                    if (_package.Dte.ItemOperations.IsFileOpen(doc.FullName, EnvDTE.Constants.vsViewKindTextView)
                      || _package.Dte.ItemOperations.IsFileOpen(doc.FullName, EnvDTE.Constants.vsViewKindCode)
                      )
                    {
                        visible = true;
                        break;
                    }
                }
            }

            menuCommand.Enabled = menuCommand.Visible = visible;
        }

        void VSFormatFileEventHandler(object sender, EventArgs e)
        {
            FormatSelectedItems();
        }

        void VSFormatFileOpenedDocumentsEventHandler(object sender, EventArgs e)
        {
            if (_package.Dte != null
                && _package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeDocument
                && _package.Dte.Documents != null
                )
            {
                var list = _package.Dte.Documents.OfType<EnvDTE.Document>().ToList();

                foreach (var doc in list)
                {
                    if (doc.ActiveWindow == null
                      || doc.ActiveWindow.Type != EnvDTE.vsWindowType.vsWindowTypeDocument
                      || doc.ActiveWindow.Visible == false
                      )
                    {
                        continue;
                    }

                    if (_package.Dte.ItemOperations.IsFileOpen(doc.FullName, EnvDTE.Constants.vsViewKindTextView)
                        || _package.Dte.ItemOperations.IsFileOpen(doc.FullName, EnvDTE.Constants.vsViewKindCode)
                        )
                    {
                        _package.Format(doc);
                    }
                }
            }
        }

        void FormatSelectedItems()
        {
            if (_package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeDocument
                && _package.Dte.ActiveWindow.Document != null
                )
            {
                _package.Format(_package.Dte.ActiveWindow.Document);
            }
            else if (_package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeSolutionExplorer
                && _package.Dte.SelectedItems != null
                )
            {
                foreach (EnvDTE.UIHierarchyItem selectedItem in (object[])_package.Dte.ToolWindows.SolutionExplorer.SelectedItems)
                    FormatItem(selectedItem.Object);
            }
        }

        void FormatItem(object item)
        {
            var solution = item as EnvDTE.Solution;
            if (solution != null)
            {
                foreach (EnvDTE.Project subProject in solution.Projects)
                    FormatItem(subProject);
                return;
            }

            if (item is EnvDTE.Project project)
            {
                if (project.Kind == EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder)
                    foreach (EnvDTE.ProjectItem projectSubItem in project.ProjectItems)
                        FormatItem(projectSubItem.SubProject);
                else
                    foreach (EnvDTE.ProjectItem projectSubItem in project.ProjectItems)
                        FormatItem(projectSubItem);
                return;
            }

            if (item is EnvDTE.ProjectItem projectItem)
                if (projectItem.ProjectItems != null && projectItem.ProjectItems.Count > 0)
                    foreach (EnvDTE.ProjectItem subProjectItem in projectItem.ProjectItems)
                        FormatItem(subProjectItem);
                else
                    FormatProjectItem(projectItem);
        }

        void FormatProjectItem(EnvDTE.ProjectItem item)
        {
            if (!_package.OptionsPage.AllowDenyFilter.IsAllowed(item.Name))
                return;

            EnvDTE.Window documentWindow = null;
            try
            {
                if (!item.IsOpen[EnvDTE.Constants.vsViewKindTextView])
                {
                    documentWindow = item.Open(EnvDTE.Constants.vsViewKindTextView);
                    if (documentWindow == null)
                        return;
                }

                if (_package.Format(item.Document))
                    item.Document.Save();
            }
            catch (COMException)
            {
                _package.OutputString($"Failed to process {item.Name}.");
            }
            finally
            {
                documentWindow?.Close();
            }
        }
    }
}