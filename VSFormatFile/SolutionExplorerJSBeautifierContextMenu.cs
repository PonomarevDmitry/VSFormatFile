using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;

namespace Nav.Common.VSPackages.VSFormatFile
{
    class SolutionExplorerJSBeautifierContextMenu
    {
        readonly VSFormatFilePackage _package;

        public SolutionExplorerJSBeautifierContextMenu(VSFormatFilePackage package)
        {
            _package = package;

            var mcs = _package.MenuCommandService;

            var menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetCode, PackageIds.CmdIdJSBeautifierOnSaveCode);
            var menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetCode, PackageIds.CmdIdJSBeautifierOnSaveOpenDocuments);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveOpenedDocumentsEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            menuItem.BeforeQueryStatus += oleMenuItemDocuments_BeforeQueryStatus;
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetFile, PackageIds.CmdIdJSBeautifierOnSaveFile);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetFolder, PackageIds.CmdIdJSBeautifierOnSaveFolder);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetProject, PackageIds.CmdIdJSBeautifierOnSaveProject);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetSolution, PackageIds.CmdIdJSBeautifierOnSaveSolution);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(PackageGuids.GuidVSFormatFileCmdSetSolution, PackageIds.CmdIdJSBeautifierOnSaveSolutionFolder);
            menuItem = new OleMenuCommand(JSBeautifierOnSaveEventHandler, menuCommandId)
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

        void JSBeautifierOnSaveEventHandler(object sender, EventArgs e)
        {
            JSBeautifierSelectedItems();
        }

        void JSBeautifierOnSaveOpenedDocumentsEventHandler(object sender, EventArgs e)
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
                        _package.JSBeautifier(doc);
                    }
                }
            }
        }

        void JSBeautifierSelectedItems()
        {
            if (_package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeDocument
                && _package.Dte.ActiveWindow.Document != null
                )
            {
                _package.JSBeautifier(_package.Dte.ActiveWindow.Document);
            }
            else if (_package.Dte.ActiveWindow != null
                && _package.Dte.ActiveWindow.Type == EnvDTE.vsWindowType.vsWindowTypeSolutionExplorer
                && _package.Dte.SelectedItems != null
                )
            {
                foreach (EnvDTE.UIHierarchyItem selectedItem in (object[])_package.Dte.ToolWindows.SolutionExplorer.SelectedItems)
                    JSBeautifierItem(selectedItem.Object);
            }
        }

        void JSBeautifierItem(object item)
        {
            var solution = item as EnvDTE.Solution;
            if (solution != null)
            {
                foreach (EnvDTE.Project subProject in solution.Projects)
                    JSBeautifierItem(subProject);
                return;
            }

            if (item is EnvDTE.Project project)
            {
                if (project.Kind == EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder)
                    foreach (EnvDTE.ProjectItem projectSubItem in project.ProjectItems)
                        JSBeautifierItem(projectSubItem.SubProject);
                else
                    foreach (EnvDTE.ProjectItem projectSubItem in project.ProjectItems)
                        JSBeautifierItem(projectSubItem);
                return;
            }

            if (item is EnvDTE.ProjectItem projectItem)
                if (projectItem.ProjectItems != null && projectItem.ProjectItems.Count > 0)
                    foreach (EnvDTE.ProjectItem subProjectItem in projectItem.ProjectItems)
                        JSBeautifierItem(subProjectItem);
                else
                    JSBeautifierProjectItem(projectItem);
        }

        void JSBeautifierProjectItem(EnvDTE.ProjectItem item)
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

                if (_package.JSBeautifier(item.Document))
                {
                    item.Document.Save();
                }
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