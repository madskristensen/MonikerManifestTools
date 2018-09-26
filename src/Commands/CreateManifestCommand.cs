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
            string folder = item?.FileNames[1];

            if (!TryGetFileName(folder, out string manifestFileName))
            {
                return;
            }

            dte.CheckFileOutOfSourceControl(manifestFileName);

            if (TryGenerateManifest(item, manifestFileName))
            {
                IncludeManifestInProjectAndVsix(item.ContainingProject, manifestFileName);
                SetInputImagesAsResource(item.ContainingProject, folder);

                VsShellUtilities.OpenDocument(package, manifestFileName);

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

        private static bool TryGenerateManifest(ProjectItem item, string manifestFileName)
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
                string manifestName = Path.GetFileName(manifestFileName);


                string args = $"/manifest:\"{manifestName}\" /assembly:\"{assemblyName}\" /resources:\"{folder}\" /guidName:{Path.GetFileNameWithoutExtension(manifestFileName)}Guid /rootPath:\"{projectPath}\"";

                var start = new ProcessStartInfo
                {
                    WorkingDirectory = folder,
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

                return File.Exists(manifestFileName);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }

            return false;
        }

        private static bool TryGetFileName(string initialDirectory, out string fileName)
        {
            fileName = null;

            using (var dialog = new SaveFileDialog())
            {
                dialog.InitialDirectory = initialDirectory;
                dialog.FileName = "Monikers";
                dialog.DefaultExt = ".imagemanifest";
                dialog.Filter = "Image Manifest files | *.imagemanifest";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return false;

                fileName = dialog.FileName;
            }

            return true;
        }

        private static void IncludeManifestInProjectAndVsix(Project project, string fileName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ProjectItem item = project.DTE.AddFileToProject(project, fileName);

            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));

            solution.GetProjectOfUniqueName(item.ContainingProject.UniqueName, out IVsHierarchy hierarchy);

            if (hierarchy is IVsBuildPropertyStorage buildPropertyStorage)
            {
                hierarchy.ParseCanonicalName(fileName, out uint itemId);
                buildPropertyStorage.SetItemAttribute(itemId, "IncludeInVSIX", "true");
            }
        }
    }
}
