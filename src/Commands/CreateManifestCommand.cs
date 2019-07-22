using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace MonikerManifestTools
{
    internal sealed class CreateManifestCommand
    {
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Assumes.Present(commandService);

            var cmdId = new CommandID(PackageGuids.guidPackageCmdSet, PackageIds.CreateManifest);
            var cmd = new MenuCommand((s, e) => Execute(package), cmdId)
            {
                Supported = false
            };

            commandService.AddCommand(cmd);
        }

        private static void Execute(IServiceProvider package)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = package.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(dte);

            ProjectItem item = dte.SelectedItems.Item(1)?.ProjectItem;
            Project project = item?.ContainingProject;

            if (project == null)
                return;

            string projectPath = project.Properties.Item("FullPath").Value.ToString().TrimEnd('\\');
            string manifestFilePath = Path.Combine(projectPath, "Monikers.imagemanifest");

            dte.CheckFileOutOfSourceControl(manifestFilePath);

            if (TryGenerateManifest(item, manifestFilePath))
            {
                string folder = item?.FileNames[1];

                IncludeManifestInProjectAndVsix(item.ContainingProject, manifestFilePath);
                SetInputImagesAsResource(item.ContainingProject, folder);

                VsShellUtilities.OpenDocument(package, manifestFilePath);

                dte.ExecuteCommand("SolutionExplorer.SyncWithActiveDocument");
            }
        }

        private static void SetInputImagesAsResource(Project project, string folder)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (string file in Directory.EnumerateFiles(folder, "*.png"))
            {
                ProjectItem item = project.DTE.Solution.FindProjectItem(file);

                if (item != null)
                {
                    item.SetItemType("Resource");
                }
                else
                {
                    project.DTE.AddFileToProject(project, file, "Resource");
                }
            }
        }

        private static bool TryGenerateManifest(ProjectItem item, string manifestFilePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Project project = item?.ContainingProject;
            string folder = item?.FileNames[1]?.TrimEnd('\\');

            if (project == null)
                return false;

            try
            {
                string assembly = Assembly.GetExecutingAssembly().Location;
                string root = Path.GetDirectoryName(assembly);
                string projectPath = project.Properties.Item("FullPath").Value.ToString().TrimEnd('\\');
                string toolsDir = Path.Combine(root, "Resources");

                string assemblyName = project.Properties.Item("AssemblyName").Value.ToString();
                string manifestName = Path.GetFileName(manifestFilePath);


                string args = $"/manifest:\"{manifestName}\" /assembly:\"{assemblyName}\" /resources:\"{folder}\" /guidName:{Path.GetFileNameWithoutExtension(manifestFilePath)}Guid /rootPath:\"{projectPath}\"";

                var start = new ProcessStartInfo
                {
                    WorkingDirectory = projectPath,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    FileName = Path.Combine(toolsDir, "ManifestFromResources.exe"),
                    Arguments = args
                };

                using (var p = new System.Diagnostics.Process())
                {
                    p.StartInfo = start;
                    p.Start();
                    p.WaitForExit();
                }

                return File.Exists(manifestFilePath);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }

            return false;
        }

        private static void IncludeManifestInProjectAndVsix(Project project, string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ProjectItem item = project.DTE.AddFileToProject(project, filePath, itemType: "content");

            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));

            solution.GetProjectOfUniqueName(item.ContainingProject.UniqueName, out IVsHierarchy hierarchy);

            if (hierarchy is IVsBuildPropertyStorage buildPropertyStorage)
            {
                hierarchy.ParseCanonicalName(filePath, out uint itemId);
                buildPropertyStorage.SetItemAttribute(itemId, "IncludeInVSIX", "true");
            }
        }
    }
}
