using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Reflection;
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
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

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
            {
                return;
            }

            var projectPath = project.Properties.Item("FullPath").Value.ToString().TrimEnd('\\');
            var manifestFilePath = Path.Combine(projectPath, "Monikers.imagemanifest");

            dte.CheckFileOutOfSourceControl(manifestFilePath);

            if (TryGenerateManifest(item, manifestFilePath))
            {
                var folder = item?.FileNames[1];

                IncludeManifestInProjectAndVsix(dte, item.ContainingProject, manifestFilePath);
                SetInputImagesAsResource(dte, item.ContainingProject, folder);

                VsShellUtilities.OpenDocument(package, manifestFilePath);
            }
        }

        private static void SetInputImagesAsResource(DTE2 dte, Project project, string folder)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var file in Directory.EnumerateFiles(folder, "*.png"))
            {
                ProjectItem item = dte.Solution.FindProjectItem(file);

                if (item != null)
                {
                    item.SetItemType("Resource");
                }
                else
                {
                    dte.AddFileToProject(project, file, "Resource");
                }
            }
        }

        private static bool TryGenerateManifest(ProjectItem item, string manifestFilePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Project project = item?.ContainingProject;
            var folder = item?.FileNames[1]?.TrimEnd('\\');

            if (project == null)
            {
                return false;
            }

            try
            {
                var assembly = Assembly.GetExecutingAssembly().Location;
                var root = Path.GetDirectoryName(assembly);
                var projectPath = project.Properties.Item("FullPath").Value.ToString().TrimEnd('\\');
                var toolsDir = Path.Combine(root, "Resources");

                var assemblyName = project.Properties.Item("AssemblyName").Value.ToString();
                var manifestName = Path.GetFileName(manifestFilePath);


                var args = $"/manifest:\"{manifestName}\" /assembly:\"{assemblyName}\" /resources:\"{folder}\" /guidName:{Path.GetFileNameWithoutExtension(manifestFilePath)}Guid /rootPath:\"{projectPath}\"";

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

        private static void IncludeManifestInProjectAndVsix(DTE2 dte, Project project, string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ProjectItem item = dte.AddFileToProject(project, filePath, itemType: "content");

            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));

            solution.GetProjectOfUniqueName(item.ContainingProject.UniqueName, out IVsHierarchy hierarchy);

            if (hierarchy is IVsBuildPropertyStorage buildPropertyStorage)
            {
                hierarchy.ParseCanonicalName(filePath, out var itemId);
                buildPropertyStorage.SetItemAttribute(itemId, "IncludeInVSIX", "true");
            }
        }
    }
}
